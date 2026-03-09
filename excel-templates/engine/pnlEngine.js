/**
 * P&L Engine
 * ==========
 * Converts 1. P&L sheet (175 formulas) into pure computation logic.
 *
 * Excel architecture (decoded):
 *   Rows 11-16: Revenue by stream (from B.I Sales yearly totals)
 *   Row 17:     Total Revenue = SUM(11:16)
 *   Row 18:     Revenue Growth % = (Year / PrevYear) - 1
 *   Rows 21-28: OPEX by category (from B.II OPEX yearly totals)
 *   Row 30:     Total OPEX = SUM(21:29)
 *   Row 33:     EBITDA = Revenue - OPEX
 *   Row 34:     EBITDA Margin = EBITDA / Revenue
 *   Row 36:     PBT (pre-depreciation) = EBITDA - Other Expenses
 *   Row 38:     Depreciation (from FA Schedule)
 *   Row 39:     PBT = Row36 - Depreciation
 *   Row 42:     Tax = IF(PBT > 0, PBT × 25.17%, 0)
 *   Row 43:     PAT = PBT - Tax
 *   Row 44:     PAT Margin = PAT / Revenue × 100
 *
 * This engine:
 *   computeYearlyPnL() → yearly P&L from revenue, opex, depreciation
 *   computeMonthlyPnL() → monthly P&L for granular view
 */

// ── Tax Configuration ────────────────────────────────────────────────
const DEFAULT_TAX_RATE = 0.2517; // 25.17% corporate tax (India)

/**
 * Compute yearly P&L statement.
 *
 * @param {Object} params
 * @param {number[]} params.yearlyRevenue - Total revenue per year
 * @param {number[]} params.yearlyOpex - Total OPEX per year
 * @param {number[]} params.yearlyDepreciation - Depreciation per year
 * @param {Object} [params.yearlyRevenueByStream] - Revenue breakdown by stream
 * @param {Object} [params.yearlyOpexByCategory] - OPEX breakdown by category
 * @param {number} [params.taxRate=0.2517]
 * @returns {Array<Object>} Yearly P&L rows
 */
function computeYearlyPnL({
    yearlyRevenue,
    yearlyOpex,
    yearlyDepreciation = [],
    yearlyRevenueByStream = {},
    yearlyOpexByCategory = {},
    taxRate = DEFAULT_TAX_RATE,
}) {
    const years = yearlyRevenue.length;
    const results = [];

    for (let y = 0; y < years; y++) {
        const revenue = yearlyRevenue[y];
        const opex = yearlyOpex[y];
        const depreciation = yearlyDepreciation[y] || 0;

        // ── Core P&L calculations ──
        const ebitda = revenue - opex;
        const ebitdaMargin = revenue > 0 ? ebitda / revenue : 0;
        const pbt = ebitda - depreciation;
        const tax = pbt > 0 ? pbt * taxRate : 0;
        const pat = pbt - tax;
        const patMargin = revenue > 0 ? (pat / revenue) * 100 : 0;

        // ── Growth metrics ──
        const revenueGrowth = y > 0 && yearlyRevenue[y - 1] > 0
            ? (revenue / yearlyRevenue[y - 1]) - 1
            : 0;
        const opexGrowth = y > 0 && yearlyOpex[y - 1] > 0
            ? (opex / yearlyOpex[y - 1]) - 1
            : 0;

        // ── Revenue breakdown ──
        const revenueBreakdown = {};
        for (const [stream, values] of Object.entries(yearlyRevenueByStream)) {
            revenueBreakdown[stream] = values[y] || 0;
        }

        // ── OPEX breakdown ──
        const opexBreakdown = {};
        for (const [category, values] of Object.entries(yearlyOpexByCategory)) {
            opexBreakdown[category] = values[y] || 0;
        }

        results.push({
            year: y + 1,
            fy: `Y${y + 1}`,
            revenue: round(revenue),
            opex: round(opex),
            ebitda: round(ebitda),
            ebitdaMargin: round(ebitdaMargin * 100, 2),
            depreciation: round(depreciation),
            pbt: round(pbt),
            tax: round(tax),
            pat: round(pat),
            patMargin: round(patMargin, 2),
            revenueGrowth: round(revenueGrowth * 100, 2),
            opexGrowth: round(opexGrowth * 100, 2),
            revenueBreakdown,
            opexBreakdown,
        });
    }

    return results;
}

/**
 * Compute monthly P&L (more granular than yearly).
 *
 * @param {Object} params
 * @param {number[]} params.monthlyRevenue - Revenue per month
 * @param {number[]} params.monthlyOpex - OPEX per month
 * @param {number} [params.taxRate=0.2517]
 * @returns {Array<{ month, revenue, opex, ebitda, tax, pat }>}
 */
function computeMonthlyPnL({
    monthlyRevenue,
    monthlyOpex,
    taxRate = DEFAULT_TAX_RATE,
}) {
    const months = Math.min(monthlyRevenue.length, monthlyOpex.length);
    const results = [];

    for (let m = 0; m < months; m++) {
        const revenue = monthlyRevenue[m];
        const opex = monthlyOpex[m];
        const ebitda = revenue - opex;
        const tax = ebitda > 0 ? ebitda * taxRate : 0;
        const pat = ebitda - tax;

        results.push({
            month: m,
            revenue: round(revenue),
            opex: round(opex),
            ebitda: round(ebitda),
            ebitdaMargin: revenue > 0 ? round((ebitda / revenue) * 100, 2) : 0,
            tax: round(tax),
            pat: round(pat),
        });
    }

    return results;
}

/**
 * Compute summary statistics across all years.
 *
 * @param {Array<Object>} yearlyPnL - Output from computeYearlyPnL
 * @returns {Object} Summary with totals, averages, and CAGR
 */
function computeSummary(yearlyPnL) {
    const totalRevenue = yearlyPnL.reduce((sum, y) => sum + y.revenue, 0);
    const totalOpex = yearlyPnL.reduce((sum, y) => sum + y.opex, 0);
    const totalPat = yearlyPnL.reduce((sum, y) => sum + y.pat, 0);
    const avgEbitdaMargin = yearlyPnL.reduce((sum, y) => sum + y.ebitdaMargin, 0) / yearlyPnL.length;

    // CAGR: (EndValue / StartValue)^(1/n) - 1
    const firstRevenue = yearlyPnL[0]?.revenue || 0;
    const lastRevenue = yearlyPnL[yearlyPnL.length - 1]?.revenue || 0;
    const n = yearlyPnL.length - 1;
    const revenueCagr = firstRevenue > 0 && n > 0
        ? Math.pow(lastRevenue / firstRevenue, 1 / n) - 1
        : 0;

    // Break-even month (first month where cumulative PAT >= 0)
    let cumulativePat = 0;
    let breakEvenYear = null;
    for (const year of yearlyPnL) {
        cumulativePat += year.pat;
        if (cumulativePat >= 0 && breakEvenYear === null) {
            breakEvenYear = year.year;
        }
    }

    return {
        totalRevenue: round(totalRevenue),
        totalOpex: round(totalOpex),
        totalEbitda: round(totalRevenue - totalOpex),
        totalPat: round(totalPat),
        avgEbitdaMargin: round(avgEbitdaMargin, 2),
        revenueCagr: round(revenueCagr * 100, 2),
        breakEvenYear,
        peakRevenue: round(Math.max(...yearlyPnL.map(y => y.revenue))),
        peakPat: round(Math.max(...yearlyPnL.map(y => y.pat))),
    };
}

// ── Utility ──────────────────────────────────────────────────────────

function round(value, decimals = 2) {
    const factor = Math.pow(10, decimals);
    return Math.round(value * factor) / factor;
}

module.exports = {
    DEFAULT_TAX_RATE,
    computeYearlyPnL,
    computeMonthlyPnL,
    computeSummary,
};
