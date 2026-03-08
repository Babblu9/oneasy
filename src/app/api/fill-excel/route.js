import { exec } from 'child_process';
import { promisify } from 'util';
import fs from 'fs';
import path from 'path';

const execAsync = promisify(exec);

export async function POST(request) {
    try {
        const { branches, product, units, price } = await request.json();

        const timestamp = Date.now();
        const outputPath = path.join(process.cwd(), `tmp_output_${timestamp}.xlsx`);
        const scriptPath = path.join(process.cwd(), 'fill_excel_v2.js');

        // Run Node.js script (v2 with exceljs)
        const nodePath = '/usr/local/bin/node';
        const command = `"${nodePath}" "${scriptPath}" "${branches}" "${product}" "${units}" "${price}" "${outputPath}"`;

        await execAsync(command);

        if (!fs.existsSync(outputPath)) {
            throw new Error("Resulting file not found.");
        }

        const fileBuffer = fs.readFileSync(outputPath);
        const uint8Array = new Uint8Array(fileBuffer);

        // Cleanup generated file after reading
        fs.unlinkSync(outputPath);

        return new Response(uint8Array, {
            status: 200,
            headers: {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': `attachment; filename="Plan_${product.replace(/\s+/g, '_')}.xlsx"`,
                'Content-Length': uint8Array.byteLength.toString(),
            },
        });
    } catch (error) {
        console.error("API Error:", error);
        return new Response(JSON.stringify({ error: error.message }), {
            status: 500,
            headers: { 'Content-Type': 'application/json' },
        });
    }
}
