import { z } from 'zod';

const streamSchema = z.object({
    name: z.string().optional(),
    stream: z.string().optional(),
    substream: z.string().optional(),
    label: z.string().optional(),
    price: z.union([z.number(), z.string()]).optional(),
    quantity: z.union([z.number(), z.string()]).optional(),
    value: z.union([z.number(), z.string()]).optional(),
});

const fundingSchema = z.object({
    total: z.union([z.number(), z.string()]).optional(),
    founder_investment: z.union([z.number(), z.string()]).optional(),
    bank_loan: z.union([z.number(), z.string()]).optional(),
    loan_interest: z.union([z.number(), z.string()]).optional(),
    loan_tenure: z.union([z.number(), z.string()]).optional(),
});

export const generateModelRequestSchema = z.object({
    industry: z.string().optional(),
    templateId: z.string().optional(),
    pricing_model: z.string().optional(),
    launchDate: z.string().optional(),
    monthlyGrowthRate: z.union([z.number(), z.string()]).optional(),
    revenue_streams: z.array(streamSchema).optional(),
    opex: z.array(streamSchema).optional(),
    funding: fundingSchema.optional(),
});

export function parseMoneyLike(input) {
    if (input == null || input === '') return 0;
    if (typeof input === 'number') return Number.isFinite(input) ? input : 0;

    let s = String(input).trim().toLowerCase();
    if (!s) return 0;
    s = s.replace(/[,$₹€£\s]/g, '');

    let multiplier = 1;
    if (s.endsWith('k')) {
        multiplier = 1_000;
        s = s.slice(0, -1);
    } else if (s.endsWith('m')) {
        multiplier = 1_000_000;
        s = s.slice(0, -1);
    } else if (s.endsWith('b')) {
        multiplier = 1_000_000_000;
        s = s.slice(0, -1);
    } else if (s.endsWith('l') || s.endsWith('lac') || s.endsWith('lakh')) {
        multiplier = 100_000;
        s = s.replace(/(l|lac|lakh)$/, '');
    } else if (s.endsWith('cr') || s.endsWith('crore')) {
        multiplier = 10_000_000;
        s = s.replace(/(cr|crore)$/, '');
    }

    const n = Number(s);
    if (!Number.isFinite(n)) return 0;
    return Math.round(n * multiplier);
}
