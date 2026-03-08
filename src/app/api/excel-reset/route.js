import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import {
    BUSINESS_INFO_CELLS,
    FUNDING_CELLS,
    ASSUMPTION_CELLS,
} from '../../../lib/excelCellMap.js';
import { revenueCatalog, opexCatalog } from '../../../lib/modelCatalog.js';

const SOURCE_EXCEL = path.join(process.cwd(), 'Docty-Healthcare', 'Docty Healthcare - Business Plan.xlsx');
const WORK_EXCEL = path.join(process.cwd(), 'Docty-Healthcare', 'Docty_Healthcare_Working.xlsx');

/**
 * POST /api/excel-reset
 * Clears all AI-fillable input cells in the working copy, starting fresh.
 * Keeps all formulas and formatting intact.
 * Revenue/OPEX cell refs now sourced from model_inputs.json via modelCatalog.
 */
export async function POST() {
    try {
        if (!fs.existsSync(SOURCE_EXCEL)) {
            return Response.json({ error: `Source Excel not found: ${SOURCE_EXCEL}` }, { status: 404 });
        }

        // Always start from the original source — fresh slate
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(SOURCE_EXCEL);

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

        // Clear revenue stream volume/price/growth cells (from model_inputs.json)
        const SHEET_REVENUE = 'A.I Revenue Streams - P1';
        revenueCatalog.forEach(product => {
            const refs = product.cellRefs;
            if (refs.qty) clearCell(SHEET_REVENUE, refs.qty);
            if (refs.price) clearCell(SHEET_REVENUE, refs.price);
            if (refs.growth_Y2_monthly) clearCell(SHEET_REVENUE, refs.growth_Y2_monthly);
            if (refs.growth_Y2_Y3) clearCell(SHEET_REVENUE, refs.growth_Y2_Y3);
            if (refs.growth_Y3_Y4) clearCell(SHEET_REVENUE, refs.growth_Y3_Y4);
            if (refs.growth_Y4_Y5) clearCell(SHEET_REVENUE, refs.growth_Y4_Y5);
            if (refs.growth_Y5_Y6) clearCell(SHEET_REVENUE, refs.growth_Y5_Y6);
            if (refs.growth_Y6_Y7) clearCell(SHEET_REVENUE, refs.growth_Y6_Y7);
        });

        // Clear OPEX cost cells (from model_inputs.json)
        const SHEET_OPEX = 'A.IIOPEX';
        opexCatalog.forEach(item => {
            const refs = item.cellRefs;
            if (refs.cost) clearCell(SHEET_OPEX, refs.cost);
            if (refs.growth_Y1_monthly) clearCell(SHEET_OPEX, refs.growth_Y1_monthly);
            if (refs.growth_Y1_Y2) clearCell(SHEET_OPEX, refs.growth_Y1_Y2);
            if (refs.growth_Y2_Y3) clearCell(SHEET_OPEX, refs.growth_Y2_Y3);
            if (refs.growth_Y3_Y4) clearCell(SHEET_OPEX, refs.growth_Y3_Y4);
            if (refs.growth_Y4_Y5) clearCell(SHEET_OPEX, refs.growth_Y4_Y5);
            if (refs.growth_Y5_Y6) clearCell(SHEET_OPEX, refs.growth_Y5_Y6);
        });

        // Save to working copy
        await wb.xlsx.writeFile(WORK_EXCEL);

        return Response.json({ success: true, clearedCells: cleared });
    } catch (err) {
        console.error('excel-reset error:', err);
        return Response.json({ error: err.message }, { status: 500 });
    }
}
