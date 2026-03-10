import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { BUSINESS_INFO_CELLS, FUNDING_CELLS, ASSUMPTION_CELLS } from '../../../lib/excelCellMap.js';
import { revenueCatalog, opexCatalog } from '../../../lib/modelCatalog.js';

// /tmp is the only writable directory on Vercel
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
 * POST /api/excel-reset
 * Clears all AI-fillable input cells in the working copy, starting fresh.
 * Keeps all formulas and formatting intact.
 * Revenue/OPEX cell refs now sourced from model_inputs.json via modelCatalog.
 */
export async function POST() {
    try {
        // Seed /tmp with master template if needed (Vercel cold-start)
        try { ensureWorkingFile(); } catch (e) {
            return Response.json({ error: e.message }, { status: 404 });
        }

        // Always scrub the CURRENT active template
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(WORK_EXCEL);

        let cleared = 0;

        const clearCell = (sheetName, cellAddr) => {
            const ws = wb.getWorksheet(sheetName);
            if (!ws) return;
            const cell = ws.getCell(cellAddr);
            // Only clear value cells (not formula cells — keep formulas intact)
            if (!cell.formula && !cell.sharedFormula) {
                cell.value = null;
                cleared++;
            }
        };

        // Clear all Business Info cells
        Object.values(BUSINESS_INFO_CELLS).forEach(({ sheet, cell }) => clearCell(sheet, cell));

        // Clear Funding cells
        Object.values(FUNDING_CELLS).forEach(({ sheet, cell }) => clearCell(sheet, cell));

        // Clear Assumption cells
        Object.values(ASSUMPTION_CELLS).forEach(({ sheet, cell }) => clearCell(sheet, cell));

        // Clear only known Revenue/OPEX editable input cells from model catalog
        revenueCatalog.forEach(item => {
            if (item?.cellRefs?.qty) clearCell('A.I Revenue Streams - P1', item.cellRefs.qty);
            if (item?.cellRefs?.price) clearCell('A.I Revenue Streams - P1', item.cellRefs.price);
        });

        opexCatalog.forEach(item => {
            if (item?.cellRefs?.cost) clearCell('A.IIOPEX', item.cellRefs.cost);
        });

        // Reset branch multiplier inputs
        clearCell('A.I Revenue Streams - P1', 'H7');
        clearCell('A.IIOPEX', 'I8');

        // Wipe cached formula results only; never touch formula text or static labels/headers
        wb.worksheets.forEach(ws => {
            ws.eachRow({ includeEmpty: false }, (row) => {
                row.eachCell({ includeEmpty: false }, (cell) => {
                    if (!(cell.formula || cell.sharedFormula)) return;
                    if (cell.value && cell.value.result !== undefined) {
                        cell.value = { ...cell.value, result: undefined };
                    }
                });
            });
        });

        // Save to working copy
        await wb.xlsx.writeFile(WORK_EXCEL);

        console.log(`[excel-reset] Deep cleaned ${cleared} raw data fields from active memory.`);

        return Response.json({ success: true, clearedCells: cleared });
    } catch (err) {
        console.error('excel-reset error:', err);
        return Response.json({ error: err.message }, { status: 500 });
    }
}
