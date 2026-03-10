/**
 * modelStateToPatches.js
 * =======================
 * Converts the full in-UI financial model state into a flat array
 * of Excel cell patches for /api/excel-fill.
 *
 * Row/column layout is derived from inspecting the actual Excel template.
 *
 * ─── REVENUE SHEET (A.I Revenue Streams - P1) ────────────────────────────
 * Template uses column C as a row-ID (never write there).
 * Groups occupy fixed row ranges (5 groups × 5 items each):
 *
 *   Group  Header Row  Item Rows    Columns
 *   ─────  ──────────  ─────────────────────────────
 *     1        9       10-14        E=label  G=qty  I=price  K-O=growth
 *     2       18       19-23
 *     3       27       28-32
 *     4       36       37-41
 *     5       45       46-50
 *
 * ─── OPEX SHEET (A.IIOPEX) ───────────────────────────────────────────────
 * Groups occupy fixed row ranges (5 groups, up to 7 items each):
 *
 *   Group  Header Row  Item Rows      Columns
 *   ─────  ──────────  ──────────────────────────────────────────
 *     1        9       10-16          E=label  G=qty  H=cost/unit  J-N=growth
 *     2       18       19-23
 *     3       24       25-31
 *     4       32       33-38
 *     5       39       40-44
 *
 * ─── BASICS SHEET (1. Basics) ───────────────────────────────────────────
 * Column D contains input values. Row numbers from inspection.
 */

// ─── Sheet Names ──────────────────────────────────────────────────────────────
const SHEET_BASICS = '1. Basics';
const SHEET_REVENUE = 'A.I Revenue Streams - P1';
const SHEET_OPEX = 'A.IIOPEX';
const SHEET_PROJECT_COST = '2.Total Project Cost';

// ─── Revenue group row layout ─────────────────────────────────────────────────
// Derived from exact template inspection. Row 10 is group 1's '1a' which has a
// formula — skip it. Write to rows 11-15 (1b through Grand Total row).
const REVENUE_GROUP_ROWS = [
    [9, [11, 12, 13, 14, 15]],  // Group 1: header D9, items 1b-1e + row 15
    [18, [19, 20, 21, 22, 23]],  // Group 2: header D18, items 2a-2e
    [27, [28, 29, 30, 31, 32]],  // Group 3: header D27, items 3a-3e
    [36, [37, 38, 39, 40, 41]],  // Group 4: header D36, items 4a-4e
    [45, [46, 47, 48, 49, 50]],  // Group 5: header D45, items 5a-5e
];

// Revenue columns — confirmed from template header row 8 inspection:
// D=Major Header, E=Minor Header/label, G=Products/month, I=Sale Price, K=Growth Y2m, L=Growth Y2-Y3, M=Growth Y3-Y4, N=Growth Y4-Y5, O=Growth Y5-Y6
const REV_COL_LABEL = 'E';
const REV_COL_QTY = 'G';
const REV_COL_PRICE = 'I';
const REV_GROWTH_COLS = { gY1: 'K', gY2: 'L', gY3: 'M', gY4: 'N', gY5: 'O' };

// ─── OPEX group row layout ────────────────────────────────────────────────────
// Each entry: [groupHeaderRow, [item rows...]]
const OPEX_GROUP_ROWS = [
    [9, [10, 11, 12, 13, 14, 15, 16]],  // Group 1 — Technology (7 items)
    [18, [19, 20, 21, 22, 23]],           // Group 2 — Legal (5 items)
    [24, [25, 26, 27, 28, 29, 30, 31]],  // Group 3 — Utilities (7 items)
    [32, [33, 34, 35, 36, 37, 38]],      // Group 4 — People (6 items)
    [39, [40, 41, 42, 43, 44]],          // Group 5 — Marketing (5 items)
];

// OPEX columns (label E, qty G, cost-per-unit H, growth J-N)
const OPEX_COL_LABEL = 'E';
const OPEX_COL_QTY = 'G';
const OPEX_COL_COST = 'H';
const OPEX_GROWTH_COLS = { gY1: 'J', gY2: 'K', gY3: 'L', gY4: 'M', gY5: 'N' };

// ─── Basics sheet cell positions ─────────────────────────────────────────────
const BASICS_CELLS = {
    legalName: { cell: 'D2' },
    tradeName: { cell: 'D4' },
    address: { cell: 'D6' },
    email: { cell: 'D8' },
    contact: { cell: 'D10' },
    promoters: { cell: 'D12' },
    startDateP1: { cell: 'D16' },
    description: { cell: 'D18' },
    burningDesire: { cell: 'D22' },
};

// ─── Helper ───────────────────────────────────────────────────────────────────
function p(sheet, cell, value, extra = {}) {
    return { sheet, cell, value, ...extra };
}

// ─── Revenue patches ──────────────────────────────────────────────────────────
function revToPatches(groups = []) {
    const patches = [];
    const groupList = Array.isArray(groups) ? groups : [];

    groupList.forEach((group, gi) => {
        const groupLayout = REVENUE_GROUP_ROWS[gi];
        if (!groupLayout) return;
        const [headerRow, itemRows] = groupLayout;

        // Write group header label
        const hdr = String(group.header || '').trim();
        if (hdr) {
            patches.push(p(SHEET_REVENUE, `D${headerRow}`, hdr));
        }

        // Write each item
        const items = Array.isArray(group.items) ? group.items : [];
        items.forEach((item, ii) => {
            const row = itemRows[ii];
            if (!row) return;

            const sub = String(item.sub || '').trim();
            const qty = Number(item.qty) || 0;
            const price = Number(item.price) || 0;

            if (sub) patches.push(p(SHEET_REVENUE, `${REV_COL_LABEL}${row}`, sub));
            if (qty > 0) patches.push(p(SHEET_REVENUE, `${REV_COL_QTY}${row}`, qty));
            if (price > 0) patches.push(p(SHEET_REVENUE, `${REV_COL_PRICE}${row}`, price));

            // Growth rates
            Object.entries(REV_GROWTH_COLS).forEach(([field, col]) => {
                const g = Number(item[field]);
                if (Number.isFinite(g) && g !== 0) {
                    patches.push(p(SHEET_REVENUE, `${col}${row}`, g));
                }
            });
        });
    });
    return patches;
}

// ─── OPEX patches ─────────────────────────────────────────────────────────────
function opexToPatches(groups = []) {
    const patches = [];
    const groupList = Array.isArray(groups) ? groups : [];

    groupList.forEach((group, gi) => {
        const groupLayout = OPEX_GROUP_ROWS[gi];
        if (!groupLayout) return;
        const [headerRow, itemRows] = groupLayout;

        // Write group header into D (Major Header column)
        const hdr = String(group.header || '').trim();
        if (hdr) {
            patches.push(p(SHEET_OPEX, `D${headerRow}`, hdr));
        }

        const items = Array.isArray(group.items) ? group.items : [];
        items.forEach((item, ii) => {
            const row = itemRows[ii];
            if (!row) return;

            const sub = String(item.sub || '').trim();
            const qty = Number(item.qty) || 0;
            const cost = Number(item.cost) || 0;

            if (sub) patches.push(p(SHEET_OPEX, `${OPEX_COL_LABEL}${row}`, sub));
            // Only write qty if it's a real headcount/unit value (avoid date format issues from qty=1)
            // The template G column is typed as number — always pass as rounded integer
            if (qty > 0) {
                patches.push(p(SHEET_OPEX, `${OPEX_COL_QTY}${row}`, Math.round(qty)));
            }
            if (cost > 0) patches.push(p(SHEET_OPEX, `${OPEX_COL_COST}${row}`, cost));

            // Growth rates
            Object.entries(OPEX_GROWTH_COLS).forEach(([field, col]) => {
                const g = Number(item[field]);
                if (Number.isFinite(g) && g !== 0) {
                    patches.push(p(SHEET_OPEX, `${col}${row}`, g));
                }
            });
        });
    });
    return patches;
}

// ─── Basics patches ───────────────────────────────────────────────────────────
function basicsToPatches(basics = {}, loan1 = {}, loan2 = {}, totalProjectCost = {}) {
    const patches = [];

    Object.entries(BASICS_CELLS).forEach(([field, { cell }]) => {
        const val = basics[field];
        if (val != null && String(val).trim()) {
            patches.push(p(SHEET_BASICS, cell, String(val).trim()));
        }
    });

    // Loan 1
    if (loan1?.amount > 0) {
        patches.push(p(SHEET_BASICS, 'D26', Number(loan1.amount)));
        if (loan1.rate) patches.push(p(SHEET_BASICS, 'D28', Number(loan1.rate) / 100));
        if (loan1.duration) patches.push(p(SHEET_BASICS, 'D30', Number(loan1.duration)));
    }

    // Total Project Cost sheet
    if (totalProjectCost?.total > 0) {
        patches.push(p(SHEET_PROJECT_COST, 'C7', Number(totalProjectCost.total)));
        if (totalProjectCost.promoterContrib > 0)
            patches.push(p(SHEET_PROJECT_COST, 'C8', Number(totalProjectCost.promoterContrib)));
        if (totalProjectCost.termLoan > 0)
            patches.push(p(SHEET_PROJECT_COST, 'C9', Number(totalProjectCost.termLoan)));
    }

    return patches;
}

// ─── Main export ──────────────────────────────────────────────────────────────
/**
 * @param {{ basics, revP1, opexP1, loan1, loan2, totalProjectCost }} state
 * @returns {Array<{sheet, cell, value}>}
 */
export function modelStateToPatches(state = {}) {
    const {
        basics = {},
        revP1 = [],
        opexP1 = [],
        loan1 = {},
        loan2 = {},
        totalProjectCost = {},
    } = state;

    const all = [
        ...basicsToPatches(basics, loan1, loan2, totalProjectCost),
        ...revToPatches(revP1),
        ...opexToPatches(opexP1),
    ];

    // Filter out any patches with null/undefined/empty values
    return all.filter(patch =>
        patch &&
        patch.sheet &&
        patch.cell &&
        patch.value !== null &&
        patch.value !== undefined &&
        patch.value !== ''
    );
}
