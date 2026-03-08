import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';

const SOURCE_EXCEL = path.join(process.cwd(), 'Docty-Healthcare', 'Docty Healthcare - Business Plan.xlsx');
const WORK_EXCEL = path.join(process.cwd(), 'Docty-Healthcare', 'Docty_Healthcare_Working.xlsx');

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

        if (!fs.existsSync(SOURCE_EXCEL)) {
            return Response.json({ error: `Excel file not found: ${SOURCE_EXCEL}` }, { status: 404 });
        }

        // Work on a copy — preserve the original untouched
        const wb = new ExcelJS.Workbook();

        // Use the working copy if it exists, otherwise start from source
        const readFrom = fs.existsSync(WORK_EXCEL) ? WORK_EXCEL : SOURCE_EXCEL;
        await wb.xlsx.readFile(readFrom);

        let patchedCount = 0;
        const patchedSheets = new Set();
        const errors = [];

        for (const patch of patches) {
            const { sheet, cell, value } = patch;
            if (!sheet || !cell || value === undefined) {
                errors.push(`Invalid patch: ${JSON.stringify(patch)}`);
                continue;
            }

            const ws = wb.getWorksheet(sheet);
            if (!ws) {
                errors.push(`Sheet not found: "${sheet}"`);
                continue;
            }

            try {
                const targetCell = ws.getCell(cell);
                // Preserve formula structure — only overwrite value cells (not formula cells)
                // For formula cells, set the result cache so it shows the new value
                if (targetCell.formula || targetCell.sharedFormula) {
                    // Set the cached result value so the viewer can read it
                    targetCell.value = { formula: targetCell.formula || targetCell.sharedFormula, result: value };
                } else {
                    targetCell.value = value;
                }
                patchedCount++;
                patchedSheets.add(sheet);
            } catch (cellErr) {
                errors.push(`Cell error at ${sheet}!${cell}: ${cellErr.message}`);
            }
        }

        // Save working copy
        await wb.xlsx.writeFile(WORK_EXCEL);

        return Response.json({
            success: true,
            patchedCount,
            patchedSheets: [...patchedSheets],
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
        const filePath = fs.existsSync(WORK_EXCEL) ? WORK_EXCEL : SOURCE_EXCEL;
        const buffer = fs.readFileSync(filePath);

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
