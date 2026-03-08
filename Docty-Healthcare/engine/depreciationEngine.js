/**
 * Depreciation Engine
 * ===================
 * Converts FA Schedule sheet (129 formulas) into pure computation logic.
 *
 * Excel architecture:
 *   FA Schedule tracks fixed assets with acquisition cost, useful life,
 *   and computes annual depreciation using straight-line method.
 *   Accumulated depreciation and net book value tracked per year.
 *
 * Used by P&L (row 38): PBT = EBITDA - Depreciation
 * Used by Balance Sheet (row 23-24): Gross Block, Accumulated Dep, Net Block
 */

/**
 * Compute annual depreciation for a single asset.
 *
 * @param {Object} asset
 * @param {string} asset.name - Asset description
 * @param {number} asset.cost - Acquisition cost (₹)
 * @param {number} asset.usefulLife - Useful life in years
 * @param {number} asset.acquisitionYear - Year index (0-based) when acquired
 * @param {number} [totalYears=6]
 * @returns {{
 *   annualDepreciation: number,
 *   schedule: Array<{ year, depreciation, accumulated, netBookValue }>
 * }}
 */
function computeAssetDepreciation(asset, totalYears = 6) {
    const annualDepreciation = asset.usefulLife > 0
        ? asset.cost / asset.usefulLife
        : 0;

    const schedule = [];
    let accumulated = 0;

    for (let y = 0; y < totalYears; y++) {
        let depreciation = 0;

        if (y >= asset.acquisitionYear && accumulated < asset.cost) {
            depreciation = Math.min(annualDepreciation, asset.cost - accumulated);
        }

        accumulated += depreciation;
        const netBookValue = asset.cost - accumulated;

        schedule.push({
            year: y,
            depreciation: Math.round(depreciation * 100) / 100,
            accumulated: Math.round(accumulated * 100) / 100,
            netBookValue: Math.round(netBookValue * 100) / 100,
        });
    }

    return { annualDepreciation: Math.round(annualDepreciation * 100) / 100, schedule };
}

/**
 * Compute total depreciation across all fixed assets.
 *
 * @param {Array<Object>} assets - All fixed assets with:
 *   { name, cost, usefulLife, acquisitionYear }
 * @param {number} [totalYears=6]
 * @returns {{
 *   assets: Array,           // Per-asset depreciation schedules
 *   yearlyTotal: number[],   // Total depreciation per year
 *   yearlyAccumulated: number[],  // Cumulative depreciation
 *   yearlyNetBlock: number[]      // Net book value per year
 * }}
 */
function computeAllDepreciation(assets, totalYears = 6) {
    const assetResults = [];
    const yearlyTotal = new Array(totalYears).fill(0);
    const totalCost = assets.reduce((sum, a) => sum + a.cost, 0);

    for (const asset of assets) {
        const result = computeAssetDepreciation(asset, totalYears);
        assetResults.push({ name: asset.name, ...result });

        for (let y = 0; y < totalYears; y++) {
            yearlyTotal[y] += result.schedule[y].depreciation;
        }
    }

    // ── Cumulative and net block ──
    const yearlyAccumulated = [];
    const yearlyNetBlock = [];
    let runningAccumulated = 0;

    for (let y = 0; y < totalYears; y++) {
        runningAccumulated += yearlyTotal[y];
        yearlyAccumulated.push(Math.round(runningAccumulated * 100) / 100);
        yearlyNetBlock.push(Math.round((totalCost - runningAccumulated) * 100) / 100);
    }

    return {
        totalCost,
        assets: assetResults,
        yearlyTotal: yearlyTotal.map(v => Math.round(v * 100) / 100),
        yearlyAccumulated,
        yearlyNetBlock,
    };
}

module.exports = {
    computeAssetDepreciation,
    computeAllDepreciation,
};
