/**
 * Excel Cell Map
 * ==============
 * Maps AI [DATA: {...}] action types to specific cells in the real Excel workbook.
 * 
 * This is a CLIENT-SIDE module (imported by FinancialContext.js).
 * It does NOT import model_inputs.json — cell references come from:
 *   1. Explicit cellRef in AI DATA tags (preferred, always used when available)
 *   2. Lightweight fallback lookup tables (for backwards compatibility)
 */

// ── Sheet Names (real Excel names) ──────────────────────────────────────────
export const SHEET = {
    BASICS: '2. Basics',
    BRANCH: 'Branch',
    REVENUE: 'A.I Revenue Streams - P1',
    OPEX: 'A.IIOPEX',
    CAPEX: 'A.III CAPEX',
    SALES: 'B.I Sales - P1',
    PNL: '1. P&L',
    FA: 'FA Schedule',
    BALANCE: '5. Balance sheet',
    RATIOS: '6. Ratios',
    DSCR: 'DSCR',
};

// ── Business Info → Basics Sheet ────────────────────────────────────────────
export const BUSINESS_INFO_CELLS = {
    legalName: { sheet: SHEET.BASICS, cell: 'D2' },
    tradeName: { sheet: SHEET.BASICS, cell: 'D4' },
    address: { sheet: SHEET.BASICS, cell: 'D6' },
    email: { sheet: SHEET.BASICS, cell: 'D8' },
    phone: { sheet: SHEET.BASICS, cell: 'D10' },
    promoters: { sheet: SHEET.BASICS, cell: 'D12' },
    companyType: { sheet: SHEET.BASICS, cell: 'D14' },
    startDate: { sheet: SHEET.BASICS, cell: 'D16' },
    description: { sheet: SHEET.BASICS, cell: 'D18' },
    equityShares: { sheet: SHEET.BASICS, cell: 'D20' },
    faceValue: { sheet: SHEET.BASICS, cell: 'D22' },
    paidUpCapital: { sheet: SHEET.BASICS, cell: 'D24' },
};

// ── Funding → Basics Sheet ───────────────────────────────────────────────────
export const FUNDING_CELLS = {
    loanAmount: { sheet: SHEET.BASICS, cell: 'D26' },
    interestRate: { sheet: SHEET.BASICS, cell: 'D28' },
    loanTenureMonths: { sheet: SHEET.BASICS, cell: 'D30' },
    moratoriumMonths: { sheet: SHEET.BASICS, cell: 'D32' },
    equityFromPromoters: { sheet: SHEET.BASICS, cell: 'D34' },
};

// ── Assumptions ──────────────────────────────────────────────────────────────
export const ASSUMPTION_CELLS = {
    initialInvestment: { sheet: SHEET.BASICS, cell: 'D36' },
    taxRate: { sheet: SHEET.BASICS, cell: 'D38' },
    inflationRate: { sheet: SHEET.BASICS, cell: 'D40' },
};

// ── Lightweight Revenue Fallback (top ~20 products for when AI omits cellRef)
// The AI is instructed to ALWAYS include cellRef, but just in case:
const REVENUE_FALLBACK = [
    { name: 'Root Canal', qty: 'H10', price: 'J10' },
    { name: 'Teeth Extraction', qty: 'H11', price: 'J11' },
    { name: 'Wisdom tooth Removal', qty: 'H12', price: 'J12' },
    { name: 'Implant', qty: 'H13', price: 'J13' },
    { name: 'X Ray', qty: 'H14', price: 'J14' },
    { name: 'Scaling', qty: 'H15', price: 'J15' },
    { name: 'Flap Surgery', qty: 'H16', price: 'J16' },
    { name: 'Filling', qty: 'H21', price: 'J21' },
    { name: 'Crowns', qty: 'H22', price: 'J22' },
    { name: 'Consultation', qty: 'H31', price: 'J31' },
    { name: 'Medical Certificate', qty: 'H32', price: 'J32' },
    { name: 'Ethical drugs', qty: 'H70', price: 'J70' },
    { name: 'Generic drugs', qty: 'H71', price: 'J71' },
    { name: 'Nurtraceuticals', qty: 'H72', price: 'J72' },
    { name: 'FMCG', qty: 'H73', price: 'J73' },
    { name: 'Surgicals', qty: 'H74', price: 'J74' },
    { name: 'CBC', qty: 'H78', price: 'J78' },
    { name: 'Blood Glucose', qty: 'H79', price: 'J79' },
];

// ── Lightweight OPEX Fallback
const OPEX_FALLBACK = [
    { name: 'Rent', cost: 'I11' },
    { name: 'Electricity', cost: 'I12' },
    { name: 'Internet & Telephone', cost: 'I13' },
    { name: 'Maintenance & Repairs', cost: 'I14' },
    { name: 'Petty cash expense', cost: 'I15' },
    { name: 'Management', cost: 'I22' },
    { name: 'Doctors', cost: 'I23' },
    { name: 'Clinical Staff', cost: 'I24' },
    { name: 'Admin & Receptionists', cost: 'I25' },
    { name: 'Housekeeping', cost: 'I31' },
    { name: 'Lab Aggregators', cost: 'I34' },
    { name: 'Digital Marketing', cost: 'I58' },
    { name: 'Promotions', cost: 'I60' },
    { name: 'Lease Deed Registration', cost: 'I66' },
    { name: 'Drug Licenses', cost: 'I67' },
];

// ── CAPEX ────────────────────────────────────────────────────────────────────
export const CAPEX_START_ROW = 5;
export const CAPEX_COLS = {
    name: 'B', cost: 'C', usefulLife: 'D', perBranch: 'E', depreciation: 'F'
};

// ── Branch ────────────────────────────────────────────────────────────────────
export const BRANCH_COLS = {
    monthOffset: 'B',
    branchName: 'A',
    cumulativeBranches: 'E',
};

// ── Fuzzy match helpers (fallback only) ─────────────────────────────────────
function fuzzyFind(list, nameField, searchName) {
    const n = (searchName || '').toLowerCase().trim();
    if (!n) return null;
    let found = list.find(r => r[nameField].toLowerCase() === n);
    if (found) return found;
    found = list.find(r =>
        n.includes(r[nameField].toLowerCase()) || r[nameField].toLowerCase().includes(n)
    );
    return found || null;
}

export function findRevenueRow(productName) {
    return fuzzyFind(REVENUE_FALLBACK, 'name', productName);
}

export function findOpexRow(subCategory) {
    return fuzzyFind(OPEX_FALLBACK, 'name', subCategory);
}

// ── Validate cell reference format ──────────────────────────────────────────
function isValidCellRef(ref) {
    return typeof ref === 'string' && /^[A-Z]{1,3}\d{1,4}$/.test(ref);
}

// ── Build patch list for a given DATA action ─────────────────────────────────
export function dataActionToPatches(action) {
    const patches = [];

    switch (action.type) {
        case 'setBusinessInfo': {
            Object.entries(BUSINESS_INFO_CELLS).forEach(([field, loc]) => {
                if (action[field] != null && action[field] !== '') {
                    let val = action[field];
                    if (Array.isArray(val)) val = val.join(', ');
                    patches.push({ sheet: loc.sheet, cell: loc.cell, value: val });
                }
            });
            break;
        }

        case 'setFunding': {
            Object.entries(FUNDING_CELLS).forEach(([field, loc]) => {
                if (action[field] != null) {
                    patches.push({ sheet: loc.sheet, cell: loc.cell, value: Number(action[field]) });
                }
            });
            break;
        }

        case 'setAssumptions': {
            Object.entries(ASSUMPTION_CELLS).forEach(([field, loc]) => {
                if (action[field] != null) {
                    patches.push({ sheet: loc.sheet, cell: loc.cell, value: Number(action[field]) });
                }
            });
            break;
        }

        case 'setBranchCount': {
            const count = Number(action.count);
            if (count > 0 && count <= 100) {
                // Write to master branch count cell H7 in Revenue sheet
                patches.push({ sheet: SHEET.REVENUE, cell: 'H7', value: count });
                // Also write to OPEX multiplier I8
                patches.push({ sheet: SHEET.OPEX, cell: 'I8', value: count });
            }
            break;
        }

        case 'addRevenueStream': {
            // Primary: use cellRef from AI DATA tag (exact cell reference from catalog)
            if (action.cellRef && typeof action.cellRef === 'object') {
                const { qty, price: priceCell } = action.cellRef;
                if (qty && isValidCellRef(qty) && action.units != null) {
                    patches.push({ sheet: SHEET.REVENUE, cell: qty, value: Number(action.units) });
                }
                if (priceCell && isValidCellRef(priceCell) && action.price != null) {
                    patches.push({ sheet: SHEET.REVENUE, cell: priceCell, value: Number(action.price) });
                }

                // Write growth rates if provided
                if (action.growthRates && typeof action.growthRates === 'object') {
                    const growthCells = {
                        'Y2_monthly': action.cellRef.growth_Y2_monthly,
                        'Y2_Y3': action.cellRef.growth_Y2_Y3,
                        'Y3_Y4': action.cellRef.growth_Y3_Y4,
                        'Y4_Y5': action.cellRef.growth_Y4_Y5,
                        'Y5_Y6': action.cellRef.growth_Y5_Y6,
                        'Y6_Y7': action.cellRef.growth_Y6_Y7,
                    };
                    for (const [grKey, cellAddr] of Object.entries(growthCells)) {
                        if (action.growthRates[grKey] != null && cellAddr && isValidCellRef(cellAddr)) {
                            patches.push({
                                sheet: SHEET.REVENUE,
                                cell: cellAddr,
                                value: Number(action.growthRates[grKey]) / 100,
                            });
                        }
                    }
                }
            } else {
                // Fallback: fuzzy match by product name
                const row = findRevenueRow(action.productName);
                if (row) {
                    if (row.qty && action.units != null) {
                        patches.push({ sheet: SHEET.REVENUE, cell: row.qty, value: Number(action.units) });
                    }
                    if (row.price && action.price != null) {
                        patches.push({ sheet: SHEET.REVENUE, cell: row.price, value: Number(action.price) });
                    }
                }
            }
            break;
        }

        case 'addOpex': {
            // Primary: use cellRef from AI DATA tag
            if (action.cellRef && typeof action.cellRef === 'object') {
                const { cost } = action.cellRef;
                if (cost && isValidCellRef(cost)) {
                    const totalCost = (Number(action.price) || 0) * (Number(action.units) || 1);
                    patches.push({ sheet: SHEET.OPEX, cell: cost, value: totalCost });
                }

                // Write growth rates if provided
                if (action.growthRates && typeof action.growthRates === 'object') {
                    const growthCells = {
                        'Y1_monthly': action.cellRef.growth_Y1_monthly,
                        'Y1_Y2': action.cellRef.growth_Y1_Y2,
                        'Y2_Y3': action.cellRef.growth_Y2_Y3,
                        'Y3_Y4': action.cellRef.growth_Y3_Y4,
                        'Y4_Y5': action.cellRef.growth_Y4_Y5,
                        'Y5_Y6': action.cellRef.growth_Y5_Y6,
                    };
                    for (const [grKey, cellAddr] of Object.entries(growthCells)) {
                        if (action.growthRates[grKey] != null && cellAddr && isValidCellRef(cellAddr)) {
                            patches.push({
                                sheet: SHEET.OPEX,
                                cell: cellAddr,
                                value: Number(action.growthRates[grKey]) / 100,
                            });
                        }
                    }
                }
            } else {
                // Fallback: fuzzy match by sub-category name
                const row = findOpexRow(action.subCategory);
                if (row) {
                    const totalCost = (Number(action.price) || 0) * (Number(action.units) || 1);
                    if (row.cost) {
                        patches.push({ sheet: SHEET.OPEX, cell: row.cost, value: totalCost });
                    }
                }
            }
            break;
        }

        case 'addCapex': {
            // CAPEX rows appended from CAPEX_START_ROW
            break;
        }

        default:
            break;
    }

    return patches;
}
