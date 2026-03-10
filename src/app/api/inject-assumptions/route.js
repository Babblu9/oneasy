import { NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs';
import XlsxPopulate from 'xlsx-populate';
import { getTemplate, getAssumptionCells, getTemplatePath } from '@/lib/templateRegistry';

// /tmp is the only writable directory on Vercel
const WORK_EXCEL = '/tmp/active_working.xlsx';

export const maxDuration = 60;

/**
 * POST /api/inject-assumptions
 * Body: {
 *   templateId: "edtech",
 *   assumptions: {
 *     course_price: 5000,
 *     students_per_month: 200,
 *     employees: 5,
 *     funding: 2000000
 *   }
 * }
 *
 * Injects assumption values directly into the correct template cells,
 * keeping all Excel formulas intact. Writes only to the active working copy.
 */
export async function POST(req) {
    try {
        const body = await req.json();
        const { templateId, assumptions } = body;

        if (!templateId || !assumptions || typeof assumptions !== 'object') {
            return NextResponse.json(
                { error: 'templateId and assumptions object are required' },
                { status: 400 }
            );
        }

        const template = getTemplate(templateId);
        if (!template) {
            return NextResponse.json(
                { error: `Unknown template: "${templateId}"` },
                { status: 404 }
            );
        }

        const cellMap = getAssumptionCells(templateId);
        const sourceFile = getTemplatePath(templateId);

        if (!sourceFile || !fs.existsSync(sourceFile)) {
            return NextResponse.json(
                { error: `Template file not found for "${templateId}". Please upload the template first.` },
                { status: 404 }
            );
        }

        const workingFile = WORK_EXCEL;
        if (!fs.existsSync(workingFile)) {
            // Initialize a session-scoped working file from selected template source
            fs.copyFileSync(sourceFile, workingFile);
        }

        // Load workbook using xlsx-populate (preserves charts, macros, formatting)
        const workbook = await XlsxPopulate.fromFileAsync(workingFile);

        const injected = [];
        const skipped = [];

        for (const [key, value] of Object.entries(assumptions)) {
            const cellDef = cellMap[key];
            if (!cellDef) {
                skipped.push(key);
                continue;
            }

            const { sheet: sheetName, cell: cellRef } = cellDef;
            const sheet = workbook.sheet(sheetName);

            if (!sheet) {
                skipped.push(`${key} (sheet "${sheetName}" not found)`);
                continue;
            }

            const parsedValue = (value !== '' && !isNaN(Number(value))) ? Number(value) : value;
            sheet.cell(cellRef).value(parsedValue);
            injected.push({ key, sheetName, cellRef, value: parsedValue });
        }

        // Write only to the active working copy (live view)
        const buffer = await workbook.outputAsync();
        fs.writeFileSync(workingFile, buffer);

        console.log(`[inject-assumptions] ${templateId}: injected ${injected.length} values, skipped ${skipped.length}`);

        return NextResponse.json({
            success: true,
            templateId,
            templateName: template.name,
            injected,
            skipped,
            message: `✅ Injected ${injected.length} assumptions into ${template.name} template`
        });

    } catch (error) {
        console.error('inject-assumptions error:', error);
        return NextResponse.json({ error: 'Internal server error: ' + error.message }, { status: 500 });
    }
}
