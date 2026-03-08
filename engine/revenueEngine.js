/**
 * Revenue Engine
 * ==============
 * Converts B.I Sales - P1 (40,925 formulas) into pure computation logic.
 *
 * Excel architecture (decoded):
 *   Row 2: K2 = VLOOKUP(month, Branch, branches)  → branch count per month
 *          L2 = Branch!F4                          → growth factor
 *   Row 10+: K10 = BaseUnits × Branches × GrowthFactor  → monthly quantity
 *            L10 = Price × PriceGrowthFactor             → adjusted price
 *            M10 = K10 × L10                             → monthly revenue
 *   Summary: Row 68 = SUM(Clinical), Row 78 = SUM(Pharma), etc.
 *            Row 207 = Grand Total all streams
 *
 * This engine replaces all of that with:
 *   computeServiceMonthly()  → one service, all months
 *   computeAllRevenue()      → all services, aggregated
 */

const { getYearIndex } = require('./branchEngine');

// ── Growth Rate Configuration ────────────────────────────────────────
// From A.I Revenue Streams - P1: all services share these rates

const DEFAULT_REVENUE_GROWTH = {
    // Monthly compounding within each year
    monthlyRates: {
        0: 0,      // Y1: growth via branch expansion only
        1: 0.25,   // Y2: 25% monthly compound
        2: 0,      // Y3+: yearly step-ups applied at year boundary
        3: 0,
        4: 0,
        5: 0,
    },
    // Year-over-year step-ups (applied once at year boundary)
    yearlyStepUps: {
        // Applied at the START of each year to the base units
        2: 0.20,   // Y2→Y3: +20%
        3: 0.10,   // Y3→Y4: +10%
        4: 0.10,   // Y4→Y5: +10%
        5: 0.15,   // Y5→Y6: +15%
    },
};

/**
 * Compute 72-month projection for a single service.
 *
 * @param {Object} service
 * @param {number} service.baseUnits - Units sold per branch per month
 * @param {number} service.price - Sale price per unit (₹)
 * @param {boolean} service.active - Whether this service generates revenue
 * @param {Object} [service.customGrowth] - Override growth rates (optional)
 * @param {Array} branchSchedule - Full branch schedule from branchEngine
 * @param {Object} [growthConfig] - Growth rate configuration
 * @param {number} [totalMonths=72]
 * @returns {Array<{ month, quantity, price, revenue }>}
 */
function computeServiceMonthly(
    service,
    branchSchedule,
    growthConfig = DEFAULT_REVENUE_GROWTH,
    totalMonths = 72
) {
    if (!service.active || service.baseUnits <= 0 || service.price <= 0) {
        return Array.from({ length: totalMonths }, (_, m) => ({
            month: m, quantity: 0, price: 0, revenue: 0,
        }));
    }

    const results = [];
    let currentUnits = service.baseUnits;
    const currentPrice = service.price; // Price stays constant (no price inflation in model)

    for (let m = 0; m < totalMonths; m++) {
        const yearIdx = getYearIndex(m);
        const branches = branchSchedule[m]?.branches ?? branchSchedule[branchSchedule.length - 1].branches;

        // ── Apply yearly step-up at year boundary ──
        if (m > 0) {
            const prevYearIdx = getYearIndex(m - 1);
            if (yearIdx !== prevYearIdx) {
                const stepUp = growthConfig.yearlyStepUps[yearIdx] ?? 0;
                currentUnits = currentUnits * (1 + stepUp);
            }
        }

        // ── Apply monthly compounding ──
        if (m > 0) {
            const monthlyRate = growthConfig.monthlyRates[yearIdx] ?? 0;
            currentUnits = currentUnits * (1 + monthlyRate);
        }

        // ── Compute revenue ──
        const quantity = currentUnits * branches;
        const revenue = quantity * currentPrice;

        results.push({
            month: m,
            quantity: Math.round(quantity * 100) / 100,
            price: currentPrice,
            revenue: Math.round(revenue * 100) / 100,
        });
    }

    return results;
}

/**
 * Compute revenue for ALL services, grouped by stream.
 *
 * @param {Array<Object>} services - All services with:
 *   { name, streamName, subStreamName, baseUnits, price, active }
 * @param {Array} branchSchedule - From branchEngine.buildBranchSchedule()
 * @param {Object} [growthConfig] - Growth rates
 * @param {number} [totalMonths=72]
 * @returns {{
 *   streams: Object,      // Per-stream monthly totals
 *   monthlyGrand: number[],  // Grand total per month
 *   yearlyGrand: number[],   // Grand total per year
 *   yearlyByStream: Object   // Per-stream yearly totals
 * }}
 */
function computeAllRevenue(
    services,
    branchSchedule,
    growthConfig = DEFAULT_REVENUE_GROWTH,
    totalMonths = 72
) {
    const streams = {};

    // ── Compute per-service, aggregate by stream ──
    for (const service of services) {
        const monthly = computeServiceMonthly(service, branchSchedule, growthConfig, totalMonths);

        if (!streams[service.streamName]) {
            streams[service.streamName] = {
                name: service.streamName,
                services: [],
                monthlyTotal: new Array(totalMonths).fill(0),
            };
        }

        streams[service.streamName].services.push({
            name: service.name,
            subStream: service.subStreamName,
            monthly,
        });

        for (let m = 0; m < totalMonths; m++) {
            streams[service.streamName].monthlyTotal[m] += monthly[m].revenue;
        }
    }

    // ── Grand total per month (Row 207 equivalent) ──
    const monthlyGrand = new Array(totalMonths).fill(0);
    for (const stream of Object.values(streams)) {
        for (let m = 0; m < totalMonths; m++) {
            monthlyGrand[m] += stream.monthlyTotal[m];
        }
    }

    // ── Yearly aggregation ──
    const yearlyGrand = aggregateYearly(monthlyGrand);
    const yearlyByStream = {};
    for (const [name, stream] of Object.entries(streams)) {
        yearlyByStream[name] = aggregateYearly(stream.monthlyTotal);
    }

    return { streams, monthlyGrand, yearlyGrand, yearlyByStream };
}

/**
 * Aggregate monthly values into yearly totals.
 * Y1 = months 0-6 (partial), Y2 = months 7-18, Y3 = 19-30, etc.
 *
 * @param {number[]} monthly - Array of monthly values
 * @returns {number[]} Array of yearly totals
 */
function aggregateYearly(monthly) {
    const years = [];
    // Y1: months 0-6 (7 months)
    years.push(monthly.slice(0, 7).reduce((a, b) => a + b, 0));
    // Y2-Y6: 12 months each
    for (let y = 1; y <= 5; y++) {
        const start = 7 + (y - 1) * 12;
        const end = start + 12;
        years.push(monthly.slice(start, end).reduce((a, b) => a + b, 0));
    }
    return years.map(v => Math.round(v * 100) / 100);
}

module.exports = {
    DEFAULT_REVENUE_GROWTH,
    computeServiceMonthly,
    computeAllRevenue,
    aggregateYearly,
};
