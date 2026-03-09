/**
 * Financial Engine Orchestrator
 * ============================
 * The central entry point that ties all modular engines together.
 *
 * Takes a single structured input object and returns a complete
 * 72-month financial projection including monthly, yearly, and
 * summary data.
 */

const branchEngine = require('./branchEngine');
const revenueEngine = require('./revenueEngine');
const opexEngine = require('./opexEngine');
const depreciationEngine = require('./depreciationEngine');
const pnlEngine = require('./pnlEngine');

/**
 * Run the full financial projection.
 *
 * @param {Object} input - Structured financial data
 * @param {Array} input.branchSchedule - Custom branch expansion schedule
 * @param {Array} input.services - Array of service objects (revenue streams)
 * @param {Array} input.expenses - Array of expense objects (opex categories)
 * @param {Array} input.assets - Array of fixed asset objects (capex)
 * @param {Object} [input.growthConfig] - Revenue growth rate overrides
 * @param {number} [input.taxRate] - Corporate tax rate factor
 * @param {number} [totalMonths=72] - Projection horizon
 * @returns {Object} { monthly, yearly, summary }
 */
function runProjection(input, totalMonths = 72) {
    const {
        branchSchedule: rawBranchSchedule,
        services,
        expenses,
        assets = [],
        growthConfig,
        taxRate
    } = input;

    // 1. Build Branch Schedule
    const branchSchedule = branchEngine.buildBranchSchedule(rawBranchSchedule, totalMonths);

    // 2. Compute Revenue
    const revenueData = revenueEngine.computeAllRevenue(
        services,
        branchSchedule,
        growthConfig,
        totalMonths
    );

    // 3. Compute OPEX
    const opexData = opexEngine.computeAllOpex(
        expenses,
        branchSchedule,
        totalMonths
    );

    // 4. Compute Depreciation (Fixed Assets)
    const totalYears = Math.ceil(totalMonths / 12);
    const depreciationData = depreciationEngine.computeAllDepreciation(assets, totalYears);

    // 5. Compute Yearly P&L
    const yearlyPnL = pnlEngine.computeYearlyPnL({
        yearlyRevenue: revenueData.yearlyGrand,
        yearlyOpex: opexData.yearlyGrand,
        yearlyDepreciation: depreciationData.yearlyTotal,
        yearlyRevenueByStream: revenueData.yearlyByStream,
        yearlyOpexByCategory: opexData.yearlyByCategory,
        taxRate
    });

    // 6. Compute Monthly P&L (Consolidated)
    const monthlyPnL = pnlEngine.computeMonthlyPnL({
        monthlyRevenue: revenueData.monthlyGrand,
        monthlyOpex: opexData.monthlyGrand,
        taxRate
    });

    // 7. Compute Summary Stats
    const summary = pnlEngine.computeSummary(yearlyPnL);

    return {
        monthly: {
            revenue: revenueData.monthlyGrand,
            opex: opexData.monthlyGrand,
            pnl: monthlyPnL,
            byStream: revenueData.streams,
            byCategory: opexData.categories
        },
        yearly: {
            pnl: yearlyPnL
        },
        summary
    };
}

module.exports = {
    runProjection,
    ...branchEngine,
    ...revenueEngine,
    ...opexEngine,
    ...depreciationEngine,
    ...pnlEngine
};
