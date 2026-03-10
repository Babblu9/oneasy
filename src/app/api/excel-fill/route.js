import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { isAllowedSheet, isFormulaCell, isValidCellAddress, normalizeInputValue, normalizeSheetName, resolveWorksheet } from '@/lib/templateGuards';
import { detectOpexColumnsFromExcelJs, detectRevenueColumnsFromExcelJs } from '@/lib/streamInjection';

// /tmp is the only writable directory on Vercel; the master template is read-only in the project bundle
const WORK_EXCEL = '/tmp/active_working.xlsx';
const SOURCE_EXCEL = path.join(process.cwd(), 'excel-templates', 'active_working.xlsx');

export const maxDuration = 60;

/** Seed /tmp with the master template on cold-starts or when /tmp has been cleared */
function ensureWorkingFile() {
    if (!fs.existsSync(WORK_EXCEL)) {
        if (!fs.existsSync(SOURCE_EXCEL)) throw new Error(`Source template not found: ${SOURCE_EXCEL}`);
        fs.copyFileSync(SOURCE_EXCEL, WORK_EXCEL);
    }
}

/**
 * POST /api/excel-fill
 * Body: { patches: [{sheet, cell, value}, ...] }
 * Returns: { success: true, patchedCount: N }
 *
 * Writes values into the working copy of the Excel.
 * Preserves all formulas — formula results update when opened in Excel.
 */
export async function POST(request) {
    try {
        const { patches } = await request.json();

        if (!Array.isArray(patches) || patches.length === 0) {
            return Response.json({ error: 'patches must be a non-empty array' }, { status: 400 });
        }

        // Seed /tmp with master template if needed (Vercel cold-start)
        try { ensureWorkingFile(); } catch (e) {
            return Response.json({ error: e.message }, { status: 404 });
        }

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(WORK_EXCEL);

        const revenueSheetNames = ['A.I Revenue Streams - P1', 'A.I Revenue Streams - P2'];
        const opexSheetNames = ['A.IIOPEX'];
        const columnDetection = {
            revenue: {},
            opex: {},
        };

        for (const s of revenueSheetNames) {
            const ws = resolveWorksheet(wb, s);
            if (ws) columnDetection.revenue[s] = detectRevenueColumnsFromExcelJs(ws);
        }
        for (const s of opexSheetNames) {
            const ws = resolveWorksheet(wb, s);
            if (ws) columnDetection.opex[s] = detectOpexColumnsFromExcelJs(ws);
        }

        const mapColumn = (sheetName, cellAddr) => {
            const m = String(cellAddr || '').match(/^([A-Z]{1,3})(\d{1,6})$/);
            if (!m) return cellAddr;
            const col = m[1];
            const row = m[2];
            const isRevenue = revenueSheetNames.includes(sheetName);
            const isOpex = opexSheetNames.includes(sheetName);

            if (isRevenue) {
                const d = columnDetection.revenue[sheetName] || {};
                if (col === 'G' || col === 'H') return `${d.qtyCol || 'H'}${row}`;
                if (col === 'I' || col === 'J') return `${d.valueCol || 'J'}${row}`;
                if (col === 'E' || col === 'F') return `${d.labelCol || 'E'}${row}`;
            }

            if (isOpex) {
                const d = columnDetection.opex[sheetName] || {};
                if (col === 'G') return `${d.qtyCol || 'G'}${row}`;
                if (col === 'I') return `${d.valueCol || 'I'}${row}`;
                if (col === 'D' || col === 'E') return `${d.labelCol || 'E'}${row}`;
            }
            return cellAddr;
        };

        const deduped = new Map();
        for (const patch of patches) {
            const sheet = normalizeSheetName(patch?.sheet);
            const cell = mapColumn(sheet, patch?.cell);
            const key = `${sheet}!${cell}`;
            deduped.set(key, { ...patch, sheet, cell });
        }
        const normalizedPatches = [...deduped.values()];

        let patchedCount = 0;
        const patchedSheets = new Set();
        const errors = [];
        const blockedFormulaCells = [];
        const formulaReroutes = [];
        const unresolvedRevenue = new Set();
        const unresolvedOpex = new Set();
        const patchedRows = new Set();

        const parseCell = (cellAddr) => {
            const m = String(cellAddr || '').match(/^([A-Z]{1,3})(\d{1,6})$/);
            if (!m) return null;
            return { col: m[1], row: m[2] };
        };

        const writeWithFormulaFallback = (ws, normalizedSheet, cell, value, patch) => {
            const primary = ws.getCell(cell);
            if (!isFormulaCell(primary)) {
                primary.value = normalizeInputValue(value);
                return { ok: true, cell };
            }

            const parsed = parseCell(cell);
            if (!parsed) {
                const block = `${normalizedSheet}!${cell}`;
                blockedFormulaCells.push(block);
                return { ok: false, blocked: block };
            }

            const isRevenue = revenueSheetNames.includes(normalizedSheet);
            const isOpex = opexSheetNames.includes(normalizedSheet);
            const candidates = [];

            if (isRevenue) {
                const d = columnDetection.revenue[normalizedSheet] || {};
                if (parsed.col === 'G' || parsed.col === 'H') {
                    candidates.push(`${d.qtyCol || 'H'}${parsed.row}`, `G${parsed.row}`, `H${parsed.row}`);
                } else if (parsed.col === 'I' || parsed.col === 'J') {
                    candidates.push(`${d.valueCol || 'J'}${parsed.row}`, `I${parsed.row}`, `J${parsed.row}`);
                } else if (parsed.col === 'E' || parsed.col === 'F') {
                    candidates.push(`${d.labelCol || 'E'}${parsed.row}`, `E${parsed.row}`, `F${parsed.row}`);
                }
            } else if (isOpex) {
                const d = columnDetection.opex[normalizedSheet] || {};
                if (parsed.col === 'G') {
                    candidates.push(`${d.qtyCol || 'G'}${parsed.row}`, `G${parsed.row}`);
                } else if (parsed.col === 'I') {
                    candidates.push(`${d.valueCol || 'I'}${parsed.row}`, `I${parsed.row}`, `H${parsed.row}`);
                } else if (parsed.col === 'D' || parsed.col === 'E') {
                    candidates.push(`${d.labelCol || 'E'}${parsed.row}`, `E${parsed.row}`, `D${parsed.row}`);
                }
            }

            const uniqueCandidates = [...new Set(candidates)].filter(c => c !== cell);
            for (const alt of uniqueCandidates) {
                if (!isValidCellAddress(alt)) continue;
                const altCell = ws.getCell(alt);
                if (isFormulaCell(altCell)) continue;
                altCell.value = normalizeInputValue(value);
                formulaReroutes.push({
                    from: `${normalizedSheet}!${cell}`,
                    to: `${normalizedSheet}!${alt}`,
                    name: patch?.sourceName || null,
                    type: patch?.sourceType || null,
                });
                return { ok: true, cell: alt, rerouted: true };
            }

            const block = `${normalizedSheet}!${cell}`;
            blockedFormulaCells.push(block);
            return { ok: false, blocked: block };
        };

        for (const patch of normalizedPatches) {
            const { sheet, cell, value } = patch;
            if (!sheet || !cell || value === undefined) {
                errors.push(`Invalid patch: ${JSON.stringify(patch)}`);
                continue;
            }

            if (!isAllowedSheet(sheet) || !isValidCellAddress(cell)) {
                errors.push(`Blocked patch outside allowed model input surface: ${sheet}!${cell}`);
                continue;
            }

            const normalizedSheet = normalizeSheetName(sheet);
            const ws = resolveWorksheet(wb, normalizedSheet);
            if (!ws) {
                errors.push(`Sheet not found: "${normalizedSheet}"`);
                continue;
            }

            try {
                const writeResult = writeWithFormulaFallback(ws, normalizedSheet, cell, value, patch);
                if (!writeResult.ok) {
                    errors.push(`Blocked formula cell write: ${writeResult.blocked}`);
                    if (patch?.sourceName) {
                        if (patch?.sourceType === 'opex') unresolvedOpex.add(String(patch.sourceName));
                        else unresolvedRevenue.add(String(patch.sourceName));
                    }
                    continue;
                }

                patchedCount++;
                patchedSheets.add(normalizedSheet);
                const rowNum = String(writeResult.cell).replace(/^[A-Z]{1,3}/, '');
                patchedRows.add(`${normalizedSheet}:${rowNum}`);
            } catch (cellErr) {
                errors.push(`Cell error at ${normalizedSheet}!${cell}: ${cellErr.message}`);
                if (patch?.sourceName) {
                    if (patch?.sourceType === 'opex') unresolvedOpex.add(String(patch.sourceName));
                    else unresolvedRevenue.add(String(patch.sourceName));
                }
            }
        }

        // Save working copy
        await wb.xlsx.writeFile(WORK_EXCEL);

        return Response.json({
            success: true,
            patchedCount,
            patchedSheets: [...patchedSheets],
            report: {
                patchedRows: [...patchedRows].slice(0, 60),
                blockedFormulaCells: [...new Set(blockedFormulaCells)].slice(0, 60),
                formulaReroutes: formulaReroutes.slice(0, 60),
                unresolvedRevenue: [...unresolvedRevenue].slice(0, 20),
                unresolvedOpex: [...unresolvedOpex].slice(0, 20),
                columnDetection,
            },
            errors: errors.length > 0 ? errors : undefined,
        });

    } catch (err) {
        console.error('excel-fill POST error:', err);
        return Response.json({ error: err.message }, { status: 500 });
    }
}

/**
 * GET /api/excel-fill
 * Returns the current working Excel as a downloadable file.
 */
export async function GET() {
    try {
        // Seed if needed then send the working file
        try { ensureWorkingFile(); } catch (e) {
            return Response.json({ error: e.message }, { status: 404 });
        }
        const buffer = fs.readFileSync(WORK_EXCEL);

        return new Response(buffer, {
            status: 200,
            headers: {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': 'attachment; filename="Docty_Healthcare_Financial_Model.xlsx"',
                'Content-Length': buffer.length.toString(),
            },
        });
    } catch (err) {
        return Response.json({ error: err.message }, { status: 500 });
    }
}
