/**
 * Merges extracted financial data into the current knowledge graph
 */
export function mergeFinancialExtraction(current, extracted) {
    if (extracted.should_reset) {
        return {
            basics: {},
            revP1: [],
            opexP1: [],
            totalProjectCost: {},
            loan1: { amount: 0, duration: 0, rate: 0, startDate: "" },
            loan2: { amount: 0, duration: 0, rate: 0, startDate: "" },
            fixedAssets: [],
        };
    }

    const result = { ...current };

    if (extracted.basics) {
        result.basics = { ...result.basics, ...extracted.basics };
    }

    if (extracted.revenue_streams) {
        // Simple merge: append or replace based on header
        extracted.revenue_streams.forEach(stream => {
            const existingIdx = result.revP1.findIndex(s => s.header === stream.header);
            if (existingIdx > -1) {
                result.revP1[existingIdx] = { ...result.revP1[existingIdx], ...stream };
            } else {
                result.revP1.push({ id: String(result.revP1.length + 1), ...stream });
            }
        });
    }

    if (extracted.opex_streams) {
        extracted.opex_streams.forEach(stream => {
            const existingIdx = result.opexP1.findIndex(s => s.header === stream.header);
            if (existingIdx > -1) {
                result.opexP1[existingIdx] = { ...result.opexP1[existingIdx], ...stream };
            } else {
                result.opexP1.push({ id: String(result.opexP1.length + 1), ...stream });
            }
        });
    }

    if (extracted.funding) {
        if (extracted.funding.promoterContrib !== null) result.totalProjectCost.promoterContrib = extracted.funding.promoterContrib;
        if (extracted.funding.termLoan !== null) result.totalProjectCost.termLoan = extracted.funding.termLoan;
        if (extracted.funding.wcLoan !== null) result.totalProjectCost.wcLoan = extracted.funding.wcLoan;
        if (extracted.funding.loan1) result.loan1 = { ...result.loan1, ...extracted.funding.loan1 };
    }

    return result;
}

/**
 * Calculates current completion percentage
 */
export function getFinancialCompletionPercentage(kg) {
    let required = 5;
    let filled = 0;
    
    if (kg.basics?.legalName) filled++;
    if (kg.basics?.description) filled++;
    if (kg.revP1 && kg.revP1.length > 0) filled++;
    if (kg.opexP1 && kg.opexP1.length > 0) filled++;
    if (kg.totalProjectCost?.termLoan) filled++;
    
    return Math.round((filled / required) * 100);
}
