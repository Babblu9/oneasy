import { NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs';
import ExcelJS from 'exceljs';
import { applyUpdatesToHyperFormula, buildHyperFormulaFromWorkbook, extractDashboardMetrics } from '@/lib/hyperFormulaEngine';
import { isAllowedSheet, isFormulaCell, isValidCellAddress, normalizeInputValue, normalizeSheetName } from '@/lib/templateGuards';
import { getCachedEngine, setCachedEngine, workbookFingerprint } from '@/lib/hfEngineCache';

// /tmp is the only writable directory on Vercel
const WORK_EXCEL = '/tmp/active_working.xlsx';
const SOURCE_EXCEL = path.join(process.cwd(), 'excel-templates', 'active_working.xlsx');

export const maxDuration = 60;

/** Seed /tmp with the master template on cold-starts or when /tmp has been cleared */
function ensureWorkingFile() {
    if (!fs.existsSync(WORK_EXCEL)) {
        if (!fs.existsSync(SOURCE_EXCEL)) throw new Error(`Source template not found: ${SOURCE_EXCEL}`);
        fs.copyFileSync(SOURCE_EXCEL, WORK_EXCEL);
    }
}

/**
 * POST /api/recalculate
 * Body: { updates: [{ sheet, cell, value }] }
 *
 * Pipeline:
 * 1) Parse active working XLSX
 * 2) Apply safe editable-cell updates
 * 3) Build HyperFormula engine from workbook formulas/data
 * 4) Return dashboard metrics
 */
export async function POST(req) {
    const t0 = Date.now();

    try {
        const { updates = [] } = await req.json();

        // Seed /tmp with master template if needed (Vercel cold-start)
        try { ensureWorkingFile(); } catch (e) {
            return NextResponse.json({ error: e.message }, { status: 404 });
        }

        const workbookPath = WORK_EXCEL;
        const beforeFingerprint = workbookFingerprint(workbookPath);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(workbookPath);

        const applied = [];
        const blocked = [];

        for (const u of updates) {
            const sheet = String(u?.sheet || '');
            const cell = String(u?.cell || '');
            const value = u?.value;

            if (!sheet || !cell || value === undefined) {
                blocked.push({ reason: 'invalid_update', update: u });
                continue;
            }

            if (!isAllowedSheet(sheet) || !isValidCellAddress(cell)) {
                blocked.push({ reason: 'blocked_surface', update: u });
                continue;
            }

            const safeSheet = normalizeSheetName(sheet);
            const ws = workbook.getWorksheet(safeSheet);
            if (!ws) {
                blocked.push({ reason: 'sheet_not_found', update: { ...u, sheet: safeSheet } });
                continue;
            }

            const target = ws.getCell(cell);
            if (isFormulaCell(target)) {
                blocked.push({ reason: 'formula_cell', update: { ...u, sheet: safeSheet } });
                continue;
            }

            target.value = normalizeInputValue(value);
            applied.push({ sheet: safeSheet, cell, value: target.value });
        }

        await workbook.xlsx.writeFile(workbookPath);
        const afterFingerprint = workbookFingerprint(workbookPath);

        let hfContext = null;
        const cachedBefore = getCachedEngine(beforeFingerprint);

        if (cachedBefore) {
            const updatedForHf = applied.map(u => ({ ...u }));
            applyUpdatesToHyperFormula({
                hf: cachedBefore.hf,
                sheetMap: cachedBefore.sheetMap,
                updates: updatedForHf
            });
            hfContext = cachedBefore;
            setCachedEngine(afterFingerprint, hfContext);
        } else {
            hfContext = buildHyperFormulaFromWorkbook(workbook);
            setCachedEngine(afterFingerprint, hfContext);
        }

        const metrics = extractDashboardMetrics(hfContext);

        return NextResponse.json({
            success: true,
            applied,
            blocked,
            cache: {
                reused: !!cachedBefore,
                key: afterFingerprint
            },
            metrics: {
                ...metrics,
                calcTimeMs: Date.now() - t0,
            }
        });
    } catch (error) {
        console.error('recalculate error:', error);
        return NextResponse.json({ error: error.message || 'Recalculation failed' }, { status: 500 });
    }
}
