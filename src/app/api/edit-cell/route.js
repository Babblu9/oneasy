import { NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs';
import ExcelJS from 'exceljs';
import { isAllowedSheet, isFormulaCell, isValidCellAddress, normalizeInputValue, normalizeSheetName, resolveWorksheet } from '@/lib/templateGuards';

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

        const workingFile = path.join(process.cwd(), 'excel-templates', 'active_working.xlsx');
        if (!fs.existsSync(workingFile)) {
            return NextResponse.json({ error: 'Working workbook not found. Select a template first.' }, { status: 404 });
        }

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
