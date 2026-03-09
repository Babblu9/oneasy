/**
 * OPEX Engine
 * ===========
 * Converts B.II OPEX - P1 (14,130 formulas) + B.II - OPEX (3,346 formulas)
 * into pure computation logic.
 *
 * Excel architecture (decoded):
 *   B.II - OPEX Row 10: H10 = F10 × G10 (base cost)
 *                        J10 = F10 × $J$8 (units × growth)
 *                        K10 = G10 × $K$8 (price × growth)
 *                        L10 = J10 × K10  (cost = grown_units × grown_price)
 *   Pattern repeats per month for 72 months.
 *   SUMIF aggregates into yearly totals.
 *
 * Two distinct growth profiles:
 *   1. STABLE (Categories 1-5): 1% annual inflation
 *   2. DECLINING (Categories 6-7): Variable declining rates
 *
 * This engine replaces all of that with:
 *   computeExpenseMonthly()  → one expense, all months
 *   computeAllOpex()         → all expenses, aggregated
 */

const { getYearIndex } = require('./branchEngine');
const { aggregateYearly } = require('./revenueEngine');

// ── Growth Profile Definitions ───────────────────────────────────────
// From A.IIOPEX: two patterns observed across 7 expense categories

const GROWTH_PROFILES = {
    /**
     * Stable: Utilities, Salaries, Vendor Payments, Supplies, Payouts
     * 100% of cost from day 1 (no ramp), 1% annual inflation
     */
    stable: {
        yearlyRates: {
            1: 0.01, // Y1→Y2
            2: 0.01, // Y2→Y3
            3: 0.01, // Y3→Y4
            4: 0.01, // Y4→Y5
            5: 0.01, // Y5→Y6
        },
    },

    /**
     * Declining: Marketing & Promotions, Licenses & Registration
     * Higher initial growth that tapers off as brand establishes
     */
    declining: {
        yearlyRates: {
            1: 0.01, // Y1→Y2 (minimal first year)
            2: 0.45, // Y2→Y3 (aggressive marketing push)
            3: 0.40, // Y3→Y4
            4: 0.25, // Y4→Y5 (tapering)
            5: 0.15, // Y5→Y6 (maintenance)
        },
    },
};

// ── Category → Profile Mapping ───────────────────────────────────────

const CATEGORY_PROFILES = {
    'Utilities': 'stable',
    'Salaries': 'stable',
    'Vendor Payments': 'stable',
    'Supplies': 'stable',
    'Payouts': 'stable',
    'Marketing & Promotions': 'declining',
    'Licenses & Registration': 'declining',
};

/**
 * Get the growth profile for an expense category.
 * @param {string} category
 * @returns {Object} Growth profile with yearlyRates
 */
function getGrowthProfile(category) {
    const profileName = CATEGORY_PROFILES[category] || 'stable';
    return GROWTH_PROFILES[profileName];
}

/**
 * Compute 72-month projection for a single expense.
 *
 * @param {Object} expense
 * @param {number} expense.amount - Base monthly cost (₹)
 * @param {boolean} expense.perBranch - If true, cost scales with branch count
 * @param {string} expense.category - Expense category name
 * @param {boolean} expense.active - Whether this expense is active
 * @param {Array} branchSchedule - From branchEngine.buildBranchSchedule()
 * @param {number} [totalMonths=72]
 * @returns {Array<{ month, cost }>}
 */
function computeExpenseMonthly(expense, branchSchedule, totalMonths = 72) {
    if (!expense.active || expense.amount <= 0) {
        return Array.from({ length: totalMonths }, (_, m) => ({ month: m, cost: 0 }));
    }

    const profile = getGrowthProfile(expense.category);
    const results = [];
    let baseCost = expense.amount; // Cost per branch (or flat)

    for (let m = 0; m < totalMonths; m++) {
        const yearIdx = getYearIndex(m);
        const branches = branchSchedule[m]?.branches ?? branchSchedule[branchSchedule.length - 1].branches;

        // ── Apply yearly growth at year boundary ──
        if (m > 0) {
            const prevYearIdx = getYearIndex(m - 1);
            if (yearIdx !== prevYearIdx) {
                const rate = profile.yearlyRates[yearIdx] ?? 0;
                baseCost = baseCost * (1 + rate);
            }
        }

        // ── Compute monthly cost ──
        const cost = expense.perBranch
            ? baseCost * branches
            : baseCost;

        results.push({
            month: m,
            cost: Math.round(cost * 100) / 100,
        });
    }

    return results;
}

/**
 * Compute OPEX for ALL expenses, grouped by category.
 *
 * @param {Array<Object>} expenses - All expense items with:
 *   { name, category, amount, perBranch, active }
 * @param {Array} branchSchedule - From branchEngine.buildBranchSchedule()
 * @param {number} [totalMonths=72]
 * @returns {{
 *   categories: Object,        // Per-category monthly totals
 *   monthlyGrand: number[],    // Grand total per month
 *   yearlyGrand: number[],     // Grand total per year
 *   yearlyByCategory: Object   // Per-category yearly totals
 * }}
 */
function computeAllOpex(expenses, branchSchedule, totalMonths = 72) {
    const categories = {};

    for (const expense of expenses) {
        const monthly = computeExpenseMonthly(expense, branchSchedule, totalMonths);

        if (!categories[expense.category]) {
            categories[expense.category] = {
                name: expense.category,
                items: [],
                monthlyTotal: new Array(totalMonths).fill(0),
            };
        }

        categories[expense.category].items.push({
            name: expense.name,
            monthly,
        });

        for (let m = 0; m < totalMonths; m++) {
            categories[expense.category].monthlyTotal[m] += monthly[m].cost;
        }
    }

    // ── Grand total per month ──
    const monthlyGrand = new Array(totalMonths).fill(0);
    for (const cat of Object.values(categories)) {
        for (let m = 0; m < totalMonths; m++) {
            monthlyGrand[m] += cat.monthlyTotal[m];
        }
    }

    // ── Yearly aggregation ──
    const yearlyGrand = aggregateYearly(monthlyGrand);
    const yearlyByCategory = {};
    for (const [name, cat] of Object.entries(categories)) {
        yearlyByCategory[name] = aggregateYearly(cat.monthlyTotal);
    }

    return { categories, monthlyGrand, yearlyGrand, yearlyByCategory };
}

module.exports = {
    GROWTH_PROFILES,
    CATEGORY_PROFILES,
    getGrowthProfile,
    computeExpenseMonthly,
    computeAllOpex,
};
