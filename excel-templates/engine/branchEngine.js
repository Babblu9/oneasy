/**
 * Branch Engine
 * =============
 * Drives branch expansion schedule over 72 months.
 * Source: Branch sheet (421 formulas)
 *
 * Provides:
 *   - Monthly branch count
 *   - Revenue growth factors per month
 *   - Subscription trajectories (Marketplace, Retail, Corporate)
 *
 * All functions are pure — no global state, no cell references.
 */

// ── Default Branch Expansion Schedule ───────────────────────────────
// Derived from Branch sheet rows 3-74.
// Each entry: { branches, growth, newBranches, subs }

const DEFAULT_BRANCH_SCHEDULE = [
    // Y1: Sep 2025 – Mar 2026 (months 0-6)
    { branches: 1, growth: 0, newBranches: 1, subs: { marketplace: 0, retail: 0, corporate: 0 } },
    { branches: 3, growth: 0.15, newBranches: 2, subs: { marketplace: 0, retail: 60, corporate: 0 } },
    { branches: 3, growth: 0.15, newBranches: 0, subs: { marketplace: 0, retail: 75, corporate: 0 } },
    { branches: 4, growth: 0.10, newBranches: 1, subs: { marketplace: 0, retail: 100, corporate: 0 } },
    { branches: 5, growth: 0.25, newBranches: 1, subs: { marketplace: 50, retail: 130, corporate: 1000 } },
    { branches: 8, growth: 0.30, newBranches: 3, subs: { marketplace: 80, retail: 180, corporate: 1000 } },
    { branches: 10, growth: 0.20, newBranches: 2, subs: { marketplace: 100, retail: 230, corporate: 1000 } },

    // Y2: Apr 2026 – Mar 2027 (months 7-18)
    { branches: 10, growth: 0.10, newBranches: 0, subs: { marketplace: 150, retail: 300, corporate: 1200 } },
    { branches: 10, growth: 0.10, newBranches: 0, subs: { marketplace: 200, retail: 340, corporate: 1200 } },
    { branches: 10, growth: 0.18, newBranches: 0, subs: { marketplace: 250, retail: 385, corporate: 1200 } },
    { branches: 10, growth: 0.17, newBranches: 0, subs: { marketplace: 300, retail: 420, corporate: 1500 } },
    { branches: 10, growth: 0, newBranches: 0, subs: { marketplace: 350, retail: 450, corporate: 1500 } },
    { branches: 10, growth: 0, newBranches: 0, subs: { marketplace: 400, retail: 500, corporate: 1500 } },
    { branches: 10, growth: 0, newBranches: 0, subs: { marketplace: 450, retail: 500, corporate: 1800 } },
    { branches: 10, growth: 0.09, newBranches: 0, subs: { marketplace: 500, retail: 500, corporate: 1800 } },
    { branches: 10, growth: 0, newBranches: 0, subs: { marketplace: 500, retail: 500, corporate: 1800 } },
    { branches: 10, growth: 0, newBranches: 0, subs: { marketplace: 500, retail: 500, corporate: 2100 } },
    { branches: 10, growth: 0, newBranches: 0, subs: { marketplace: 500, retail: 500, corporate: 2100 } },
    { branches: 10, growth: 0, newBranches: 0, subs: { marketplace: 500, retail: 500, corporate: 2100 } },
];

// Months 19-71 (Y3-Y6): branches stay at 10, growth is 0, subs at 500/500/2100 baseline
// with cumulative subs growing per Branch sheet data.

/**
 * Build the full 72-month branch schedule.
 * @param {Array} schedule - Custom schedule entries (optional, defaults to built-in)
 * @param {number} totalMonths - Projection horizon (default 72)
 * @returns {Array<Object>} Full month-by-month schedule
 */
function buildBranchSchedule(schedule = DEFAULT_BRANCH_SCHEDULE, totalMonths = 72) {
    const full = [];

    for (let m = 0; m < totalMonths; m++) {
        if (m < schedule.length) {
            full.push({ month: m, ...schedule[m] });
        } else {
            // Steady state: last known values, zero growth
            const last = schedule[schedule.length - 1];
            full.push({
                month: m,
                branches: last.branches,
                growth: 0,
                newBranches: 0,
                subs: { ...last.subs },
            });
        }
    }

    return full;
}

/**
 * Get branch count for a specific month.
 * @param {Array} schedule - Full branch schedule from buildBranchSchedule
 * @param {number} monthIndex - 0-based month
 * @returns {number}
 */
function getBranches(schedule, monthIndex) {
    if (monthIndex < schedule.length) return schedule[monthIndex].branches;
    return schedule[schedule.length - 1].branches;
}

/**
 * Get growth factor for a specific month (1 + growth rate).
 * @param {Array} schedule - Full branch schedule
 * @param {number} monthIndex - 0-based month
 * @returns {number}
 */
function getGrowthFactor(schedule, monthIndex) {
    if (monthIndex < schedule.length) return 1 + schedule[monthIndex].growth;
    return 1;
}

/**
 * Get subscription counts for a specific month.
 * @param {Array} schedule - Full branch schedule
 * @param {number} monthIndex - 0-based month
 * @returns {{ marketplace: number, retail: number, corporate: number }}
 */
function getSubscriptions(schedule, monthIndex) {
    if (monthIndex < schedule.length) return schedule[monthIndex].subs;
    return schedule[schedule.length - 1].subs;
}

/**
 * Determine the financial year index (0-based) for a given month.
 * Y1 = months 0-6 (7 months for partial first year)
 * Y2 = months 7-18 (12 months)
 * Y3 = months 19-30, etc.
 * @param {number} monthIndex
 * @returns {number} Year index (0 = Y1, 1 = Y2, ...)
 */
function getYearIndex(monthIndex) {
    if (monthIndex < 7) return 0;  // Y1: 7 months
    return Math.floor((monthIndex - 7) / 12) + 1;  // Y2+ = 12 months each
}

module.exports = {
    DEFAULT_BRANCH_SCHEDULE,
    buildBranchSchedule,
    getBranches,
    getGrowthFactor,
    getSubscriptions,
    getYearIndex,
};
