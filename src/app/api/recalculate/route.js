import { NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs';
import ExcelJS from 'exceljs';
import { applyUpdatesToHyperFormula, buildHyperFormulaFromWorkbook, extractDashboardMetrics } from '@/lib/hyperFormulaEngine';
import { isAllowedSheet, isFormulaCell, isValidCellAddress, normalizeInputValue, normalizeSheetName } from '@/lib/templateGuards';
import { getCachedEngine, setCachedEngine, workbookFingerprint } from '@/lib/hfEngineCache';

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

        const workbookPath = path.join(process.cwd(), 'Docty-Healthcare', 'active_working.xlsx');
        if (!fs.existsSync(workbookPath)) {
            return NextResponse.json({ error: 'Working workbook not found. Select template first.' }, { status: 404 });
        }
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
