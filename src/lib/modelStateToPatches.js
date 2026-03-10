/**
 * modelStateToPatches.js
 * =======================
 * Converts the full in-UI financial model state (revP1, opexP1, basics, loan)
 * into a flat array of Excel cell patches for /api/excel-fill.
 *
 * Row assignment for the revenue/OPEX sheets uses sequential slots:
 *   Revenue:  rows 10, 11, 12, ... up to 200 (per group item with a sub name and qty/price > 0)
 *   OPEX:     rows 10, 11, 12, ... up to 200 (per group item with a name and cost > 0)
 *
 * Basics sheet cell map comes from BUSINESS_INFO_CELLS / FUNDING_CELLS in excelCellMap.js.
 */

import { BUSINESS_INFO_CELLS, FUNDING_CELLS, SHEET } from './excelCellMap';

// ─── Revenue rows (write to generic slots from row 10) ────────────────────────
const REV_START = 10;
const OPEX_START = 10;

/**
 * Build patches for revP1 groups → Revenue sheet
 * Each filled item writes: label (E col), qty (G col), price (I col)
 */
function revGroupsToPatches(groups = [], sheetName = SHEET.REVENUE) {
    const patches = [];
    let row = REV_START;

    for (const group of groups) {
        // Write group header as a label row if it has a name
        if (String(group.header || '').trim()) {
            patches.push({ sheet: sheetName, cell: `E${row}`, value: String(group.header).trim() });
            row++;
        }

        for (const item of group.items || []) {
            const sub = String(item.sub || '').trim();
            const qty = Number(item.qty) || 0;
            const price = Number(item.price) || 0;
            if (!sub && qty === 0 && price === 0) continue; // skip empty rows
            if (row > 200) break;

            if (sub) patches.push({ sheet: sheetName, cell: `E${row}`, value: sub, sourceType: 'revenue', sourceName: sub });
            if (qty > 0) {
                patches.push({ sheet: sheetName, cell: `G${row}`, value: qty, sourceType: 'revenue', sourceName: sub });
                patches.push({ sheet: sheetName, cell: `H${row}`, value: qty, sourceType: 'revenue', sourceName: sub });
            }
            if (price > 0) {
                patches.push({ sheet: sheetName, cell: `I${row}`, value: price, sourceType: 'revenue', sourceName: sub });
                patches.push({ sheet: sheetName, cell: `J${row}`, value: price, sourceType: 'revenue', sourceName: sub });
            }
            // Growth rates (Y1-Y5)
            const growthCols = { gY1: 'K', gY2: 'L', gY3: 'M', gY4: 'N', gY5: 'O' };
            for (const [field, col] of Object.entries(growthCols)) {
                const g = Number(item[field]);
                if (Number.isFinite(g)) {
                    patches.push({ sheet: sheetName, cell: `${col}${row}`, value: g, sourceType: 'revenue', sourceName: sub });
                }
            }
            row++;
        }
    }
    return patches;
}

/**
 * Build patches for opexP1 groups → OPEX sheet
 * Each filled item writes: label (D col), qty (G col), cost per unit (I col)
 */
function opexGroupsToPatches(groups = [], sheetName = SHEET.OPEX) {
    const patches = [];
    let row = OPEX_START;

    for (const group of groups) {
        if (String(group.header || '').trim()) {
            patches.push({ sheet: sheetName, cell: `D${row}`, value: String(group.header).trim() });
            patches.push({ sheet: sheetName, cell: `E${row}`, value: String(group.header).trim() });
            row++;
        }

        for (const item of group.items || []) {
            const sub = String(item.sub || '').trim();
            const qty = Number(item.qty) || 0;
            const cost = Number(item.cost) || 0;
            if (!sub && qty === 0 && cost === 0) continue;
            if (row > 200) break;

            if (sub) {
                patches.push({ sheet: sheetName, cell: `D${row}`, value: sub, sourceType: 'opex', sourceName: sub });
                patches.push({ sheet: sheetName, cell: `E${row}`, value: sub, sourceType: 'opex', sourceName: sub });
            }
            if (qty > 0) patches.push({ sheet: sheetName, cell: `G${row}`, value: qty, sourceType: 'opex', sourceName: sub });
            if (cost > 0) {
                patches.push({ sheet: sheetName, cell: `H${row}`, value: cost, sourceType: 'opex', sourceName: sub });
                patches.push({ sheet: sheetName, cell: `I${row}`, value: cost, sourceType: 'opex', sourceName: sub });
            }
            // Growth rates
            const growthCols = { gY1: 'J', gY2: 'K', gY3: 'L', gY4: 'M', gY5: 'N' };
            for (const [field, col] of Object.entries(growthCols)) {
                const g = Number(item[field]);
                if (Number.isFinite(g)) {
                    patches.push({ sheet: sheetName, cell: `${col}${row}`, value: g, sourceType: 'opex', sourceName: sub });
                }
            }
            row++;
        }
    }
    return patches;
}

/**
 * Build patches for the "1. Basics" sheet from basics + loans data
 */
function basicsToPatches(basics = {}, loan1 = {}, loan2 = {}, totalProjectCost = {}) {
    const patches = [];

    const safe = (v) => (v == null || v === '') ? null : v;

    const fields = {
        legalName: basics.legalName,
        tradeName: basics.tradeName,
        address: basics.address,
        email: basics.email,
        phone: basics.contact || basics.phone,
        promoters: basics.promoters,
        startDate: basics.startDateP1 || basics.startDate,
        description: basics.description,
    };

    for (const [field, value] of Object.entries(fields)) {
        const loc = BUSINESS_INFO_CELLS[field];
        if (loc && safe(value)) {
            patches.push({ sheet: loc.sheet, cell: loc.cell, value: String(value) });
        }
    }

    // Loan 1 (term loan)
    if (loan1?.amount) {
        patches.push({ sheet: SHEET.BASICS, cell: 'D26', value: Number(loan1.amount) });
        patches.push({ sheet: SHEET.BASICS, cell: 'D28', value: Number(loan1.rate) / 100 });
        patches.push({ sheet: SHEET.BASICS, cell: 'D30', value: Number(loan1.duration) });
    }

    // Project cost
    if (totalProjectCost?.total) {
        patches.push({ sheet: SHEET.PROJECT_COST || '2.Total Project Cost', cell: 'C7', value: Number(totalProjectCost.total) });
        if (totalProjectCost.promoterContrib) {
            patches.push({ sheet: SHEET.PROJECT_COST || '2.Total Project Cost', cell: 'C8', value: Number(totalProjectCost.promoterContrib) });
        }
        if (totalProjectCost.termLoan) {
            patches.push({ sheet: SHEET.PROJECT_COST || '2.Total Project Cost', cell: 'C9', value: Number(totalProjectCost.termLoan) });
        }
    }

    return patches;
}

/**
 * Main export: convert full model state to Excel patches
 *
 * @param {{ basics, revP1, opexP1, loan1, loan2, totalProjectCost }} state
 * @returns {Array<{sheet, cell, value}>}
 */
export function modelStateToPatches(state = {}) {
    const { basics = {}, revP1 = [], opexP1 = [], loan1 = {}, loan2 = {}, totalProjectCost = {} } = state;

    return [
        ...basicsToPatches(basics, loan1, loan2, totalProjectCost),
        ...revGroupsToPatches(revP1, SHEET.REVENUE),
        ...opexGroupsToPatches(opexP1, SHEET.OPEX),
    ].filter(p => p && p.sheet && p.cell && p.value !== undefined && p.value !== null && p.value !== '');
}
