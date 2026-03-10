import { NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs';
import ExcelJS from 'exceljs';
import { isAllowedSheet, isFormulaCell, isValidCellAddress, normalizeInputValue, normalizeSheetName, resolveWorksheet } from '@/lib/templateGuards';

// /tmp is the only writable directory on Vercel
const WORK_EXCEL = '/tmp/active_working.xlsx';
const SOURCE_EXCEL = path.join(process.cwd(), 'excel-templates', 'active_working.xlsx');

export const maxDuration = 60;

/** Seed /tmp with the master template on cold-starts */
function ensureWorkingFile() {
    if (!fs.existsSync(WORK_EXCEL)) {
        if (!fs.existsSync(SOURCE_EXCEL)) throw new Error(`Source template not found: ${SOURCE_EXCEL}`);
        fs.copyFileSync(SOURCE_EXCEL, WORK_EXCEL);
    }
}

/**
 * POST /api/edit-cell
 * Body: { sheet: "A.I Revenue Streams - P2", cell: "G10", value: 10 }
 * Edits a cell only in the active working copy (never source template).
 * Blocks writes to formula cells to preserve template integrity.
 */
export async function POST(req) {
    try {
        const body = await req.json();
        const { sheet, cell, value } = body;

        if (!sheet || !cell || value === undefined) {
            return NextResponse.json(
                { error: 'Missing required fields: sheet, cell, value' },
                { status: 400 }
            );
        }

        // Seed /tmp with master template if needed (Vercel cold-start)
        try { ensureWorkingFile(); } catch (e) {
            return NextResponse.json({ error: e.message }, { status: 404 });
        }

        const workingFile = WORK_EXCEL;

        if (!isAllowedSheet(sheet) || !isValidCellAddress(cell)) {
            return NextResponse.json(
                { error: `Blocked edit outside allowed model input surface: ${sheet}!${cell}` },
                { status: 400 }
            );
        }
        const safeSheet = normalizeSheetName(sheet);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(workingFile);

        const worksheet = resolveWorksheet(workbook, safeSheet);
        if (!worksheet) {
            return NextResponse.json(
                { error: `Sheet "${safeSheet}" not found.` },
                { status: 404 }
            );
        }

        const targetCell = worksheet.getCell(cell);
        if (isFormulaCell(targetCell)) {
            return NextResponse.json(
                { error: `Blocked formula cell write: ${safeSheet}!${cell}` },
                { status: 400 }
            );
        }
        const before = targetCell.value;

        const parsedValue = normalizeInputValue(value);
        targetCell.value = parsedValue;

        await workbook.xlsx.writeFile(workingFile);

        console.log(`[edit-cell] ${safeSheet}!${cell}: "${before}" → "${parsedValue}" (working copy)`);

        return NextResponse.json({
            success: true,
            sheet: safeSheet,
            cell,
            before,
            after: parsedValue,
            message: `✅ Updated ${cell} in "${sheet}": ${before} → ${parsedValue}`
        });

    } catch (error) {
        console.error('Error editing cell:', error);
        return NextResponse.json({ error: 'Internal server error: ' + error.message }, { status: 500 });
    }
}
