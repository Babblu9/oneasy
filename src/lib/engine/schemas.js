/**
 * engine/schemas.js
 * =================
 * Strict Zod schemas for the entire financial model pipeline.
 *
 * Input schemas  → validate what comes in from AI or API callers
 * Output schemas → document what generateFinancialModel() returns
 *
 * Keep this the single source of truth for all data shapes.
 */

import { z } from 'zod';

// ─── Helpers ─────────────────────────────────────────────────────────────────

/** Accept "₹1.2Cr", "50k", 50000, "50,000" — returns a plain number */
export function parseMoneyLike(input) {
    if (input == null || input === '') return 0;
    if (typeof input === 'number') return Number.isFinite(input) ? input : 0;

    let s = String(input).trim().toLowerCase();
    if (!s) return 0;
    s = s.replace(/[,$₹€£\s]/g, '');

    let multiplier = 1;
    if (s.endsWith('k')) { multiplier = 1_000; s = s.slice(0, -1); }
    else if (s.endsWith('m')) { multiplier = 1_000_000; s = s.slice(0, -1); }
    else if (s.endsWith('b')) { multiplier = 1_000_000_000; s = s.slice(0, -1); }
    else if (/l(ac|akh)?$/.test(s)) { multiplier = 100_000; s = s.replace(/l(ac|akh)?$/, ''); }
    else if (/cr(ore)?$/.test(s)) { multiplier = 10_000_000; s = s.replace(/cr(ore)?$/, ''); }

    const n = Number(s);
    return Number.isFinite(n) ? Math.round(n * multiplier) : 0;
}

/** Zod coercion helper: accepts number | money-string → outputs number */
const money = z
    .union([z.number(), z.string()])
    .transform(v => parseMoneyLike(v))
    .default(0);

/** Zod coercion helper: accepts number | percent-string (12% or 0.12) → outputs decimal (0.12) */
const pct = z
    .union([z.number(), z.string()])
    .transform(v => {
        const n = typeof v === 'string' ? Number(v.replace('%', '')) : Number(v);
        if (!Number.isFinite(n)) return 0.1;
        return n > 1 ? n / 100 : n;
    })
    .default(0.1);

// ─── Input Schemas ────────────────────────────────────────────────────────────

export const RevenueStreamInputSchema = z.object({
    /** Human-readable stream name, e.g. "Doctor Consultations" */
    stream: z.string().optional(),
    /** Sub-stream / product name, e.g. "Platform Access Fee" */
    name: z.string().optional(),
    /** Alias fields the AI might use */
    substream: z.string().optional(),
    label: z.string().optional(),
    /** Price per unit per month */
    price: money,
    /** Units per month */
    quantity: z.union([z.number(), z.string()]).transform(v => Math.max(0, Number(v) || 0)).default(1),
    /** Override monthly value directly (price × qty calculated if 0) */
    monthlyValue: money,
});

export const OpexItemInputSchema = z.object({
    name: z.string().optional(),
    label: z.string().optional(),
    category: z.string().optional(),
    /** Monthly cost */
    value: money,
    cost: money,
    monthlyCost: money,
});

export const LoanInputSchema = z.object({
    amount: money,
    interestRatePA: pct,
    tenureMonths: z.number().int().min(0).default(60),
    startDate: z.string().optional(),
});

export const AssumptionsInputSchema = z.object({
    /** Industry key: "healthcare" | "edtech" | "saas" | "ecommerce" | "pharma" | "consulting" | "manufacturing" */
    industry: z.string().optional().default('consulting'),
    /** Free-form business description (used for auto-classification if industry not given) */
    businessDescription: z.string().optional(),
    /** Monthly growth rate to apply to revenue streams year-over-year */
    monthlyGrowthRate: pct,
    /** Number of projection years */
    projectionYears: z.number().int().min(1).max(10).default(5),
    /** Fiscal year start (YYYY-MM-DD) */
    startDate: z.string().optional().default('2026-04-01'),
    /** Revenue streams — if empty, industry template defaults are used */
    revenueStreams: z.array(RevenueStreamInputSchema).optional().default([]),
    /** OPEX items — if empty, industry template defaults are used */
    opex: z.array(OpexItemInputSchema).optional().default([]),
    /** Business info */
    businessInfo: z.object({
        legalName: z.string().optional(),
        tradeName: z.string().optional(),
        address: z.string().optional(),
        email: z.string().optional(),
        phone: z.string().optional(),
    }).optional().default({}),
    /** Funding */
    funding: z.object({
        totalProjectCost: money,
        promoterContribution: money,
        termLoan: LoanInputSchema.optional(),
        wcLoan: LoanInputSchema.optional(),
    }).optional().default({}),
    /** Number of branches / locations */
    branches: z.number().int().min(1).default(1),
});

// ─── Output Schemas (document the shape of generateFinancialModel()) ──────────

export const RevenueStreamSchema = z.object({
    id: z.string(),
    stream: z.string(),
    name: z.string(),
    price: z.number(),
    quantity: z.number(),
    monthlyRevenue: z.number(),
    /** Revenue per year [y1, y2, y3, ...] */
    annualRevenue: z.array(z.number()),
});

export const OpexItemSchema = z.object({
    id: z.string(),
    category: z.string(),
    name: z.string(),
    monthlyCost: z.number(),
    /** Cost per year [y1, y2, y3, ...] */
    annualCost: z.array(z.number()),
});

export const YearlyProjectionSchema = z.object({
    year: z.string(),       // e.g. "2026-27"
    revenue: z.number(),
    opex: z.number(),
    ebitda: z.number(),
    ebitdaMargin: z.number(),
    netProfit: z.number(),
});

export const FinancialModelSchema = z.object({
    industry: z.string(),
    revenueStreams: z.array(RevenueStreamSchema),
    opex: z.array(OpexItemSchema),
    projections: z.array(YearlyProjectionSchema),
    summary: z.object({
        monthlyRevenue: z.number(),
        monthlyOpex: z.number(),
        monthlyNet: z.number(),
        year1Revenue: z.number(),
        year1Opex: z.number(),
        year1Ebitda: z.number(),
        growthRate: z.number(),
    }),
    /** Raw actions for Excel injection (backwards-compatible) */
    actions: z.array(z.any()),
});
