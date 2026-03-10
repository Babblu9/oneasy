/**
 * engine/financialEngine.js
 * ==========================
 * The core deterministic financial model engine.
 *
 * generateFinancialModel(assumptions) is the single public API.
 * It returns a fully-calculated FinancialModel object that:
 *   - Powers the UI tables (revenueStreams, opex, projections)
 *   - Powers the P&L summary
 *   - Produces Excel injection actions (backwards-compatible)
 *
 * Rules:
 *   - No AI calls happen here — purely deterministic math
 *   - No filesystem access — pure functions only
 *   - All inputs validated through Zod schemas before reaching here
 */

import { AssumptionsInputSchema, parseMoneyLike } from './schemas.js';
import { getIndustryTemplate, classifyIndustry, getGrowthRates } from './industryTemplates.js';

// ─── Constants ────────────────────────────────────────────────────────────────

const DEFAULT_YEARS = 5;
const FISCAL_YEAR_LABELS = ['2026-27', '2027-28', '2028-29', '2029-30', '2030-31', '2031-32', '2032-33'];

// ─── Internal: Revenue Calculation ───────────────────────────────────────────

/**
 * Calculates annual revenue for a stream across projection years.
 * Growth compound: each year applies (1 + growthRateY[n]) to the previous year.
 *
 * @param {number} monthlyBase  - base monthly revenue
 * @param {number[]} yearlyGrowthRates  - array of growth rates [y1, y2, ...]
 * @returns {number[]} annual revenue per year
 */
function projectRevenue(monthlyBase, yearlyGrowthRates) {
    const annual = [];
    let current = monthlyBase * 12;
    for (let i = 0; i < yearlyGrowthRates.length; i++) {
        if (i === 0) {
            annual.push(Math.round(current));
        } else {
            current = current * (1 + yearlyGrowthRates[i - 1]);
            annual.push(Math.round(current));
        }
    }
    return annual;
}

/**
 * Calculates annual OPEX for an item across projection years.
 * OPEX typically grows at a slower rate than revenue (operational leverage).
 */
function projectOpex(monthlyCost, yearlyGrowthRates) {
    const annual = [];
    let current = monthlyCost * 12;
    // OPEX grows at 60% of revenue growth rate (operational leverage assumption)
    for (let i = 0; i < yearlyGrowthRates.length; i++) {
        if (i === 0) {
            annual.push(Math.round(current));
        } else {
            const opexGrowth = yearlyGrowthRates[i - 1] * 0.6;
            current = current * (1 + opexGrowth);
            annual.push(Math.round(current));
        }
    }
    return annual;
}

// ─── Internal: Loan Schedule ──────────────────────────────────────────────────

/**
 * Calculates a loan EMI schedule.
 * @returns {{ emi, totalInterest, yearlyInterest: number[], yearlyPrincipal: number[] }}
 */
function calcLoanSchedule(loan, projectionYears) {
    const { amount, interestRatePA, tenureMonths, startDate } = loan;
    if (!amount || !interestRatePA || !tenureMonths) {
        return { emi: 0, totalInterest: 0, yearlyInterest: Array(projectionYears).fill(0), yearlyPrincipal: Array(projectionYears).fill(0) };
    }

    const monthly = interestRatePA / 12;
    const emi = amount * monthly * Math.pow(1 + monthly, tenureMonths)
        / (Math.pow(1 + monthly, tenureMonths) - 1);

    let bal = amount;
    const yearlyInterest = Array(projectionYears).fill(0);
    const yearlyPrincipal = Array(projectionYears).fill(0);

    const sd = startDate ? new Date(startDate) : new Date('2026-04-01');

    for (let i = 0; i < tenureMonths; i++) {
        const date = new Date(sd.getFullYear(), sd.getMonth() + i, 1);
        const fyYear = date.getMonth() >= 3
            ? date.getFullYear() - 2026
            : date.getFullYear() - 2027;
        const interest = bal * monthly;
        const principal = emi - interest;
        bal = Math.max(0, bal - principal);
        if (fyYear >= 0 && fyYear < projectionYears) {
            yearlyInterest[fyYear] += interest;
            yearlyPrincipal[fyYear] += principal;
        }
    }

    return {
        emi: Math.round(emi),
        totalInterest: Math.round(emi * tenureMonths - amount),
        yearlyInterest: yearlyInterest.map(Math.round),
        yearlyPrincipal: yearlyPrincipal.map(Math.round),
    };
}

// ─── Internal: Normalize Input to Streams ────────────────────────────────────

function normalizeRevenueStream(item, idx) {
    const stream = String(item.stream || item.substream || item.name || item.label || 'Revenue').trim();
    const name = String(item.name || item.substream || item.label || stream).trim();
    const price = parseMoneyLike(item.price);
    const quantity = Math.max(0, Number(item.quantity) || 1);
    const monthlyRevenue = item.monthlyValue
        ? parseMoneyLike(item.monthlyValue)
        : Math.round(price * quantity);

    return { stream, name, price, quantity, monthlyRevenue };
}

function normalizeOpexItem(item, idx) {
    const category = String(item.category || 'Operating Expense').trim();
    const name = String(item.name || item.label || 'Expense').trim();
    const monthlyCost = parseMoneyLike(item.monthlyCost || item.cost || item.value);
    return { category, name, monthlyCost };
}

function templateStreamToNormalized(t) {
    return { stream: t.stream, name: t.name, price: t.price, quantity: t.quantity, monthlyRevenue: Math.round(t.price * t.quantity) };
}

function templateOpexToNormalized(t) {
    return { category: t.category, name: t.name, monthlyCost: t.monthlyCost };
}

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * generateFinancialModel(rawAssumptions)
 * =======================================
 * The single entry point. Call this from API routes or the UI.
 *
 * @param {object} rawAssumptions - unchecked input (from AI or form)
 * @returns {FinancialModel} - fully calculated model
 */
export function generateFinancialModel(rawAssumptions) {
    // ── 1. Validate & normalise inputs ────────────────────────────────────────
    const parsed = AssumptionsInputSchema.safeParse(rawAssumptions || {});
    const assumptions = parsed.success ? parsed.data : AssumptionsInputSchema.parse({});

    // ── 2. Classify industry ──────────────────────────────────────────────────
    let industryId = (assumptions.industry || '').toLowerCase().trim();
    if (!industryId || industryId === 'default') {
        const classified = classifyIndustry(assumptions.businessDescription || industryId);
        industryId = classified.industryId;
    }
    const template = getIndustryTemplate(industryId);
    const growthRates = getGrowthRates(template.growthProfile);

    // ── 3. Merge user input with template defaults ────────────────────────────
    const years = assumptions.projectionYears || DEFAULT_YEARS;
    const yearlyGrowthRates = Array.from({ length: years }, (_, i) => {
        const gKey = `y${i + 1}`;
        return growthRates[gKey] ?? 0.10;
    });

    const rawRevStreams = assumptions.revenueStreams?.length
        ? assumptions.revenueStreams.map(normalizeRevenueStream)
        : template.revenueStreams.map(templateStreamToNormalized);

    const rawOpexItems = assumptions.opex?.length
        ? assumptions.opex.map(normalizeOpexItem)
        : template.opex.map(templateOpexToNormalized);

    // ── 4. Calculate projections ──────────────────────────────────────────────
    const revenueStreams = rawRevStreams.map((r, i) => ({
        id: `rev-${i + 1}`,
        stream: r.stream,
        name: r.name,
        price: r.price,
        quantity: r.quantity,
        monthlyRevenue: r.monthlyRevenue,
        annualRevenue: projectRevenue(r.monthlyRevenue, yearlyGrowthRates),
    }));

    const opexItems = rawOpexItems.map((o, i) => ({
        id: `opex-${i + 1}`,
        category: o.category,
        name: o.name,
        monthlyCost: o.monthlyCost,
        annualCost: projectOpex(o.monthlyCost, yearlyGrowthRates),
    }));

    // ── 5. P&L projections ────────────────────────────────────────────────────
    const loan1 = assumptions.funding?.termLoan;
    const loan2 = assumptions.funding?.wcLoan;
    const loanSchedule1 = loan1?.amount ? calcLoanSchedule(loan1, years) : null;
    const loanSchedule2 = loan2?.amount ? calcLoanSchedule(loan2, years) : null;

    const projections = Array.from({ length: years }, (_, yi) => {
        const revenue = revenueStreams.reduce((s, r) => s + r.annualRevenue[yi], 0);
        const opex = opexItems.reduce((s, o) => s + o.annualCost[yi], 0);
        const ebitda = revenue - opex;
        const ebitdaMargin = revenue > 0 ? ebitda / revenue : 0;
        const interest = (loanSchedule1?.yearlyInterest[yi] || 0) + (loanSchedule2?.yearlyInterest[yi] || 0);
        const ebit = ebitda - interest;
        const taxRate = 0.25;
        const tax = ebit > 0 ? ebit * taxRate : 0;
        const netProfit = ebit - tax;

        return {
            year: FISCAL_YEAR_LABELS[yi] || `Year ${yi + 1}`,
            revenue: Math.round(revenue),
            opex: Math.round(opex),
            ebitda: Math.round(ebitda),
            ebitdaMargin: Math.round(ebitdaMargin * 10000) / 10000, // 4dp
            interest: Math.round(interest),
            ebit: Math.round(ebit),
            tax: Math.round(tax),
            netProfit: Math.round(netProfit),
        };
    });

    // ── 6. Summary ────────────────────────────────────────────────────────────
    const monthlyRevenue = revenueStreams.reduce((s, r) => s + r.monthlyRevenue, 0);
    const monthlyOpex = opexItems.reduce((s, o) => s + o.monthlyCost, 0);

    const summary = {
        monthlyRevenue,
        monthlyOpex,
        monthlyNet: monthlyRevenue - monthlyOpex,
        year1Revenue: projections[0]?.revenue || 0,
        year1Opex: projections[0]?.opex || 0,
        year1Ebitda: projections[0]?.ebitda || 0,
        year5Revenue: projections[4]?.revenue || 0,
        cagr5y: projections.length >= 5 && projections[0]?.revenue > 0
            ? Math.pow(projections[4].revenue / projections[0].revenue, 1 / 4) - 1
            : yearlyGrowthRates[0],
        growthRate: yearlyGrowthRates[0],
    };

    // ── 7. Excel injection actions (backwards-compatible) ─────────────────────
    const actions = [
        {
            type: 'setStreams',
            revenue: revenueStreams.map(r => ({ label: r.name, value: r.price })),
            opex: opexItems.map(o => ({ label: o.name, value: o.monthlyCost })),
        },
        ...revenueStreams.map(r => ({
            type: 'addRevenueStream',
            streamName: r.stream,
            subName: r.stream,
            productName: r.name,
            price: r.price,
            units: r.quantity,
        })),
        ...opexItems.map(o => ({
            type: 'addOpex',
            category: o.category,
            subCategory: o.name,
            price: o.monthlyCost,
            units: 1,
        })),
    ];

    // Funding actions
    const funding = assumptions.funding || {};
    if (funding.termLoan?.amount || funding.wcLoan?.amount) {
        actions.push({
            type: 'setFunding',
            totalProjectCost: funding.totalProjectCost || 0,
            promoterContribution: funding.promoterContribution || 0,
            loanAmount: funding.termLoan?.amount || 0,
            interestRate: funding.termLoan?.interestRatePA || 0.09,
            loanTenureMonths: funding.termLoan?.tenureMonths || 60,
            wcLoanAmount: funding.wcLoan?.amount || 0,
        });
    }

    if (assumptions.businessInfo?.legalName) {
        actions.push({
            type: 'setBusinessInfo',
            ...assumptions.businessInfo,
        });
    }

    actions.push({
        type: 'setAssumption',
        key: 'revenueGrowthRate',
        value: yearlyGrowthRates[0],
    });

    // ── 8. Return full model ──────────────────────────────────────────────────
    return {
        industry: industryId,
        industryName: template.name,
        industryIcon: template.icon,
        kpis: template.kpis,
        growthProfile: template.growthProfile,
        revenueStreams,
        opex: opexItems,
        projections,
        summary,
        loanSchedule: loanSchedule1 || loanSchedule2
            ? { term: loanSchedule1, wc: loanSchedule2 }
            : null,
        assumptions: {
            years,
            yearlyGrowthRates,
            branches: assumptions.branches || 1,
            startDate: assumptions.startDate || '2026-04-01',
            businessInfo: assumptions.businessInfo || {},
        },
        actions,
    };
}

/**
 * Convenience: validate raw AI JSON and run the engine.
 * Returns { success, model, error }
 */
export function safeGenerateFinancialModel(rawInput) {
    try {
        const model = generateFinancialModel(rawInput);
        return { success: true, model };
    } catch (err) {
        return { success: false, error: err.message, model: null };
    }
}
