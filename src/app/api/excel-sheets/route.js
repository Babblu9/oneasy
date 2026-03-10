import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';

// /tmp is the only writable directory on Vercel; reads still come from the bundled project files
const WORK_EXCEL = '/tmp/active_working.xlsx';
const SOURCE_EXCEL = path.join(process.cwd(), 'excel-templates', 'active_working.xlsx');

export const maxDuration = 30;

export async function GET() {
    const EXCEL_PATH = fs.existsSync(WORK_EXCEL) ? WORK_EXCEL : SOURCE_EXCEL;

    try {
        if (!fs.existsSync(EXCEL_PATH)) {
            return Response.json({ error: `Excel file not found at: ${EXCEL_PATH}` }, { status: 404 });
        }

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(EXCEL_PATH);

        const sheetNames = wb.worksheets.map(s => s.name);
        return Response.json({ sheets: sheetNames });
    } catch (err) {
        console.error('excel-sheets GET error:', err);
        return Response.json({ error: err.message }, { status: 500 });
    }
}
