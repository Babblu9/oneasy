import { NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs';
import { getTemplate, getTemplatePath } from '@/lib/templateRegistry';

/**
 * POST /api/select-template
 * Body: { templateId: "edtech" }
 * 
 * Sets the active template for the session:
 * - Copies the selected template to the working viewer path
 * - Returns template metadata and available sheets
 */
export async function POST(req) {
    try {
        const { templateId } = await req.json();

        if (!templateId) {
            return NextResponse.json({ error: 'templateId is required' }, { status: 400 });
        }

        const template = getTemplate(templateId);
        if (!template) {
            return NextResponse.json({ error: `Unknown template: "${templateId}"` }, { status: 404 });
        }

        const sourceFile = getTemplatePath(templateId);

        if (!sourceFile || !fs.existsSync(sourceFile)) {
            return NextResponse.json({
                success: false,
                templateId,
                templateName: template.name,
                warning: `Template file "${template.file}" not found. Please upload the template first.`,
                questions: template.questions
            });
        }

        // Copy selected template to the active working location for the viewer
        const workingFile = path.join(process.cwd(), 'Docty-Healthcare', 'active_working.xlsx');

        fs.copyFileSync(sourceFile, workingFile);

        // EXTRA: Get sheet list to help frontend navigate
        const ExcelJS = await import('exceljs');
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(workingFile);
        const sheets = wb.worksheets.map(s => s.name);

        return NextResponse.json({
            success: true,
            templateId,
            templateName: template.name,
            icon: template.icon,
            questions: template.questions,
            sheets,
            message: `✅ ${template.name} template loaded`
        });

    } catch (error) {
        console.error('select-template error:', error);
        return NextResponse.json({ error: 'Internal server error: ' + error.message }, { status: 500 });
    }
}

/** GET /api/select-template — returns metadata for all templates */
export async function GET() {
    const { getAllTemplates } = await import('@/lib/templateRegistry');
    const templates = getAllTemplates();
    return NextResponse.json({ templates });
}
