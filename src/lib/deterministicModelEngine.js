import industryCatalog from '@/data/industry_streams.json';
import { parseMoneyLike } from '@/lib/modelSchemas';

function mapIndustry(input) {
    const key = String(input || '').toLowerCase();
    if (key.includes('edtech')) return 'edtech';
    if (key.includes('saas') || key.includes('software')) return 'saas';
    if (key.includes('health')) return 'healthcare';
    return 'default';
}

function normalizeRevenueStream(item) {
    const name = String(item.substream || item.name || item.label || item.stream || 'Revenue').trim();
    const stream = String(item.stream || 'Revenue').trim();
    const priceRaw = item.price != null ? item.price : item.value;
    const qtyRaw = item.quantity != null ? item.quantity : 1;
    const price = parseMoneyLike(priceRaw);
    const quantity = Math.max(1, parseMoneyLike(qtyRaw) || Number(qtyRaw) || 1);
    return { stream, name, price, quantity, monthly: price * quantity };
}

function normalizeOpex(item) {
    const name = String(item.name || item.label || 'Expense').trim();
    const valueRaw = item.value != null ? item.value : item.price;
    const value = parseMoneyLike(valueRaw);
    return { name, value };
}

export function generateDeterministicModel(input) {
    const industry = mapIndustry(input.industry || input.templateId);
    const defaults = industryCatalog[industry] || industryCatalog.default;

    const revenueInput = Array.isArray(input.revenue_streams) && input.revenue_streams.length > 0
        ? input.revenue_streams
        : defaults.revenue;
    const opexInput = Array.isArray(input.opex) && input.opex.length > 0
        ? input.opex
        : defaults.opex;

    const revenue = revenueInput.map(normalizeRevenueStream);
    const opex = opexInput.map(normalizeOpex);

    const monthlyRevenue = revenue.reduce((s, r) => s + r.monthly, 0);
    const monthlyOpex = opex.reduce((s, o) => s + o.value, 0);
    const monthlyNet = monthlyRevenue - monthlyOpex;

    const growthRate = (() => {
        const raw = input.monthlyGrowthRate;
        if (raw == null) return 0.1;
        const n = typeof raw === 'string' ? Number(raw.replace('%', '')) : Number(raw);
        if (!Number.isFinite(n)) return 0.1;
        return n > 1 ? n / 100 : n;
    })();

    const funding = {
        total: parseMoneyLike(input?.funding?.total),
        founder: parseMoneyLike(input?.funding?.founder_investment),
        loan: parseMoneyLike(input?.funding?.bank_loan),
        loanInterest: Number(input?.funding?.loan_interest || 0),
        loanTenure: Number(input?.funding?.loan_tenure || 0),
    };

    const actions = [
        {
            type: 'setStreams',
            revenue: revenue.map(r => ({ label: r.name, value: r.price })),
            opex: opex.map(o => ({ label: o.name, value: o.value })),
        },
        ...revenue.map(r => ({
            type: 'addRevenueStream',
            streamName: r.stream,
            subName: r.stream,
            productName: r.name,
            price: r.price,
            units: r.quantity,
        })),
        ...opex.map(o => ({
            type: 'addOpex',
            category: 'Operating Expense',
            subCategory: o.name,
            price: o.value,
            units: 1,
        })),
    ];

    if (funding.total || funding.founder || funding.loan) {
        actions.push({
            type: 'setFunding',
            loanAmount: funding.loan,
            interestRate: funding.loanInterest,
            loanTenureMonths: funding.loanTenure ? funding.loanTenure * 12 : 0,
            equityFromPromoters: funding.founder || funding.total,
            moratoriumMonths: 0,
        });
    }

    actions.push({ type: 'setAssumption', key: 'revenueGrowthRate', value: growthRate });
    if (input.launchDate) actions.push({ type: 'setAssumption', key: 'launchDate', value: input.launchDate });

    return {
        industry,
        monthlyRevenue,
        monthlyOpex,
        monthlyNet,
        growthRate,
        funding,
        revenue,
        opex,
        actions,
    };
}
