import fs from 'fs';
import path from 'path';
import { NextResponse } from 'next/server';

// NOTE: On Vercel, the project bundle is read-only — we cannot permanently write back to
// excel-templates/ or templates/ from a serverless function. Instead, we write the uploaded
// file to /tmp so subsequent API route calls use the fresh template for this instance.
// If permanent template storage is needed, use Vercel Blob or S3.
export const maxDuration = 60;

export async function POST(req) {
    try {
        const formData = await req.formData();
        const file = formData.get('file');

        if (!file) {
            return NextResponse.json({ error: "No file received." }, { status: 400 });
        }

        const buffer = Buffer.from(await file.arrayBuffer());

        // Write uploaded template directly as the active working copy in /tmp
        const workExcelPath = '/tmp/active_working.xlsx';
        fs.writeFileSync(workExcelPath, buffer);

        return NextResponse.json({ message: "Template uploaded and activated successfully." }, { status: 200 });

    } catch (error) {
        console.error("Error occurred while uploading custom template: ", error);
        return NextResponse.json({ error: "Failed to upload template." }, { status: 500 });
    }
}
