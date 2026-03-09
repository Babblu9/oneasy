import fs from 'fs';
import path from 'path';
import { NextResponse } from 'next/server';

export async function POST(req) {
    try {
        const formData = await req.formData();
        const file = formData.get('file');

        if (!file) {
            return NextResponse.json({ error: "No file received." }, { status: 400 });
        }

        const buffer = Buffer.from(await file.arrayBuffer());

        // Path 1: The template used by the Generator API
        const generatorPath = path.join(process.cwd(), 'templates', 'finance-model-final.xlsx');

        // Path 2: The source excel used by the live Spreadsheet Viewer
        const viewerPath = path.join(process.cwd(), 'Docty-Healthcare', 'Docty Healthcare - Business Plan.xlsx');

        // Write the fresh file to both locations to keep the entire engine in sync
        fs.writeFileSync(generatorPath, buffer);
        fs.writeFileSync(viewerPath, buffer);

        // Discard the working copy so the viewer is forced to render the new source template
        const workExcelPath = path.join(process.cwd(), 'Docty-Healthcare', 'active_working.xlsx');
        if (fs.existsSync(workExcelPath)) {
            fs.unlinkSync(workExcelPath);
        }

        return NextResponse.json({ message: "Template over-written and synced successfully." }, { status: 200 });

    } catch (error) {
        console.error("Error occurred while uploading custom template: ", error);
        return NextResponse.json({ error: "Failed to upload template." }, { status: 500 });
    }
}
