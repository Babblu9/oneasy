import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { buildHyperFormulaFromWorkbook, getHfCellValue } from '@/lib/hyperFormulaEngine';
import { resolveWorksheet } from '@/lib/templateGuards';

// /tmp is the only writable directory on Vercel
const WORK_EXCEL = '/tmp/active_working.xlsx';
const SOURCE_EXCEL = path.join(process.cwd(), 'excel-templates', 'active_working.xlsx');
const TEMPLATE_FALLBACKS = [
    path.join(process.cwd(), 'templates', 'healthcare_master.xlsx'),
    path.join(process.cwd(), 'templates', 'saas_master.xlsx'),
    path.join(process.cwd(), 'templates', 'edtech_master.xlsx'),
];

export const maxDuration = 60;

/** Seed /tmp with the master template on cold-starts or when /tmp has been cleared */
function ensureWorkingFile() {
    if (!fs.existsSync(WORK_EXCEL)) {
        if (fs.existsSync(SOURCE_EXCEL)) {
            fs.copyFileSync(SOURCE_EXCEL, WORK_EXCEL);
        }
    }
}

// Map UI tab names → real Excel sheet names (Standardized for Professional models)
const TAB_TO_SHEET = {
    '1. Basics': '1. Basics',
    'Basics': '1. Basics',
    'Revenue Model': 'A.I Revenue Streams - P1',
    'Revenue': 'A.I Revenue Streams - P1',
    'OPEX': 'A.IIOPEX',
    'A.II OPEX': 'A.IIOPEX',
    'Project Cost': '2.Total Project Cost',
    'P&L': '4. P&L',
    'Balance Sheet': '5. Balance Sheet',
    '5. Balance sheet': '5. Balance sheet',
    '4. P&L': '4. P&L',
    '5. Balance Sheet': '5. Balance Sheet',
};

function colToLetter(colNumber) {
    let n = Number(colNumber);
    let s = '';
    while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s || 'A';
}

async function loadWorkbookWithRecovery() {
    // Try /tmp working copy first, then project bundle, then template fallbacks
    ensureWorkingFile();
    const candidates = [WORK_EXCEL, SOURCE_EXCEL, ...TEMPLATE_FALLBACKS].filter(p => fs.existsSync(p));
    const errors = [];

    for (const filePath of candidates) {
        try {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(filePath);
            return { wb, filePath, recovered: filePath !== WORK_EXCEL, errors };
        } catch (err) {
            errors.push(`${filePath}: ${err.message}`);
        }
    }

    throw new Error(`Unable to load workbook from any candidate. ${errors.join(' | ')}`);
}

/**
 * GET /api/excel-workbook?sheet=<tabName>
 * Returns cell data for a single sheet as JSON for lightweight rendering.
 * Reads from the working copy (patched) if it exists, else the source.
 */
export async function GET(request) {
    const { searchParams } = new URL(request.url);
    const tabName = searchParams.get('sheet') || '2. Basics';
    const sheetName = TAB_TO_SHEET[tabName] || tabName;

    try {
        const { wb, filePath, recovered, errors } = await loadWorkbookWithRecovery();
        if (recovered && filePath && fs.existsSync(filePath) && filePath !== WORK_EXCEL) {
            try {
                // Copy recovered file to /tmp so subsequent requests use the patched version
                fs.copyFileSync(filePath, WORK_EXCEL);
            } catch {
                // non-fatal; we can still serve the recovered workbook for this request
            }
        }


        const ws = resolveWorksheet(wb, sheetName) || wb.worksheets[0];

        if (!ws) {
            // Return list of available sheets if absolutely no sheets are found
            const sheetNames = wb.worksheets.map(s => s.name);
            return Response.json({ error: `Sheet "${sheetName}" not found`, available: sheetNames }, { status: 404 });
        }

        // Use the actual resolved name in case we fell back to the first sheet
        const actualSheetName = ws.name;

        // Extract cell data — limit to first 60 rows × 20 cols for performance on heavy sheets
        const isHeavySheet = ['B.I Sales - P1', 'B.II OPEX - P1', 'B.I Sales - P1 (2)', 'B.II - OPEX'].includes(actualSheetName);
        const maxRows = isHeavySheet ? 50 : 120;
        const maxCols = isHeavySheet ? 20 : 25;

        const rows = [];
        const colWidths = [];
        let hfContext = null;
        const ensureHfContext = () => {
            if (!hfContext) hfContext = buildHyperFormulaFromWorkbook(wb);
            return hfContext;
        };

        // Get column widths
        if (ws.columns && Array.isArray(ws.columns)) {
            ws.columns.forEach((col, idx) => {
                if (idx < maxCols) {
                    colWidths.push(Math.min(Math.max(col.width || 10, 6), 40));
                }
            });
        }

        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber > maxRows) return;
            const cells = [];
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                if (colNumber > maxCols) return;
                const style = cell.style || {};
                const font = style.font || {};
                const fill = style.fill || {};
                const alignment = style.alignment || {};
                const border = style.border || {};

                // Compute display value
                let value = cell.value;
                if (value !== null && value !== undefined) {
                    if (typeof value === 'object' && value !== null) {
                        if (value.formula || value.sharedFormula) {
                            // The real calculation happens when the user clicks 'Download Excel' and opens the file in Native Excel.
                            // ExcelJS doesn't contain a calculation engine.
                            if (value.result !== undefined && value.result !== null) {
                                if (typeof value.result === 'object') {
                                    value = value.result.error ? '#ERROR!' : (value.result.value ?? '');
                                } else {
                                    value = value.result;
                                }
                            } else {
                                // Fallback to HyperFormula runtime evaluation when Excel cache is blank
                                const { hf, sheetMap } = ensureHfContext();
                                const cellAddr = `${colToLetter(colNumber)}${rowNumber}`;
                                const hv = getHfCellValue(hf, sheetMap, actualSheetName, cellAddr);
                                value = hv != null ? hv : '';
                            }
                        } else if (value instanceof Date) {
                            value = value.toLocaleDateString('en-IN');
                        } else if (value.richText) {
                            value = value.richText.map(rt => rt.text || '').join('');
                        } else if (value.text) {
                            value = value.text;
                        }
                    }
                } else {
                    value = '';
                }

                // Format numbers
                if (typeof value === 'number') {
                    const numFmt = cell.numFmt || '';
                    if (numFmt.includes('%')) {
                        value = (value * 100).toFixed(1) + '%';
                    } else if (numFmt.includes('0.00') || numFmt.includes('#,##0')) {
                        value = Number(value).toLocaleString('en-IN', { maximumFractionDigits: 2 });
                    } else if (Math.abs(value) > 1000) {
                        value = Number(value).toLocaleString('en-IN', { maximumFractionDigits: 0 });
                    } else {
                        value = Number(value);
                    }
                }

                cells.push({
                    col: colNumber,
                    value: String(value ?? ''),
                    rawFormula: (cell.value && typeof cell.value === 'object' && (cell.value.formula || cell.value.sharedFormula)) ? (cell.value.formula || cell.value.sharedFormula) : null,
                    bold: font.bold || false,
                    italic: font.italic || false,
                    color: font.color?.argb || null,
                    bgColor: fill.fgColor?.argb || (fill.pattern === 'solid' ? fill.bgColor?.argb : null) || null,
                    align: alignment.horizontal || 'left',
                    wrap: alignment.wrapText || false,
                    hasBorder: !!(border.bottom || border.top),
                });
            });
            rows.push({ row: rowNumber, cells });
        });

        return Response.json({
            sheet: actualSheetName,
            tab: tabName,
            rows,
            colWidths,
            totalRows: ws.rowCount,
            totalCols: ws.columnCount,
            truncated: ws.rowCount > maxRows || ws.columnCount > maxCols,
            recoveredFrom: recovered ? path.basename(filePath) : null,
            loadWarnings: errors.length > 0 ? errors : undefined,
        });

    } catch (err) {
        console.error('excel-workbook GET error:', err);
        return Response.json({ error: err.message }, { status: 500 });
    }
}
