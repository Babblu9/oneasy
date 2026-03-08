import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';

// Path to the real Excel file (in the Docty-Healthcare sub-folder)
const SOURCE_EXCEL = path.join(process.cwd(), 'Docty-Healthcare', 'Docty Healthcare - Business Plan.xlsx');
// Working copy — patched by /api/excel-fill and reset by /api/excel-reset
const WORK_EXCEL = path.join(process.cwd(), 'Docty-Healthcare', 'Docty_Healthcare_Working.xlsx');

// Map UI tab names → real Excel sheet names
const TAB_TO_SHEET = {
    '2. Basics': '2. Basics',
    'Branch': 'Branch',
    'A.I Revenue Streams': 'A.I Revenue Streams - P1',
    'A.II OPEX': 'A.IIOPEX',
    'A.III CAPEX': 'A.III CAPEX',
    'B.I Sales - P1': 'B.I Sales - P1',
    '1. P&L': '1. P&L',
    'B.II OPEX': 'B.II - OPEX',
    'FA Schedule': 'FA Schedule',
    '5. Balance Sheet': '5. Balance sheet',
    '6. Ratios': '6. Ratios',
    'DSCR': 'DSCR',
};

/**
 * GET /api/excel-workbook?sheet=<tabName>
 * Returns cell data for a single sheet as JSON for lightweight rendering.
 * Reads from the working copy (patched) if it exists, else the source.
 */
export async function GET(request) {
    const { searchParams } = new URL(request.url);
    const tabName = searchParams.get('sheet') || '2. Basics';
    const sheetName = TAB_TO_SHEET[tabName] || tabName;

    // ── KEY FIX: always prefer working copy (contains AI patches / reset data) ──
    const EXCEL_PATH = fs.existsSync(WORK_EXCEL) ? WORK_EXCEL : SOURCE_EXCEL;

    try {
        if (!fs.existsSync(EXCEL_PATH)) {
            return Response.json({ error: `Excel file not found at: ${EXCEL_PATH}` }, { status: 404 });
        }

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(EXCEL_PATH);


        const ws = wb.getWorksheet(sheetName);
        if (!ws) {
            // Return list of available sheets if sheet not found
            const sheetNames = wb.worksheets.map(s => s.name);
            return Response.json({ error: `Sheet "${sheetName}" not found`, available: sheetNames }, { status: 404 });
        }

        // Extract cell data — limit to first 60 rows × 20 cols for performance on heavy sheets
        const isHeavySheet = ['B.I Sales - P1', 'B.II OPEX - P1', 'B.I Sales - P1 (2)', 'B.II - OPEX'].includes(sheetName);
        const maxRows = isHeavySheet ? 50 : 120;
        const maxCols = isHeavySheet ? 20 : 25;

        const rows = [];
        const colWidths = [];

        // Get column widths
        ws.columns.forEach((col, idx) => {
            if (idx < maxCols) {
                colWidths.push(Math.min(Math.max(col.width || 10, 6), 40));
            }
        });

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
                            // Use cached result if available
                            value = value.result ?? value.formula ?? '';
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
            sheet: sheetName,
            tab: tabName,
            rows,
            colWidths,
            totalRows: ws.rowCount,
            totalCols: ws.columnCount,
            truncated: ws.rowCount > maxRows || ws.columnCount > maxCols,
        });

    } catch (err) {
        console.error('excel-workbook GET error:', err);
        return Response.json({ error: err.message }, { status: 500 });
    }
}
