/**
 * POST /api/generate-model
 * ========================
 * Deterministic financial model generation endpoint.
 *
 * Input (AssumptionsInputSchema):
 *   { industry, revenueStreams, opex, growthRate, funding, businessInfo, ... }
 *
 * Output (FinancialModel):
 *   { revenueStreams, opex, projections, summary, actions, ... }
 *
 * The AI should call this endpoint after collecting user inputs.
 * It does NOT call the AI — it runs the deterministic engine.
 */

import { NextResponse } from 'next/server';
import { generateFinancialModel } from '@/lib/engine';
import { AssumptionsInputSchema } from '@/lib/engine';

export const maxDuration = 30;

export async function POST(req) {
    try {
        const body = await req.json();

        // Validate input with strict Zod schema
        const parsed = AssumptionsInputSchema.safeParse(body || {});
        if (!parsed.success) {
            return NextResponse.json({
                error: 'Invalid input',
                details: parsed.error.flatten(),
            }, { status: 400 });
        }

        // Run the deterministic engine
        const model = generateFinancialModel(parsed.data);

        return NextResponse.json({
            success: true,
            industry: model.industry,
            industryName: model.industryName,
            icon: model.industryIcon,
            kpis: model.kpis,
            growthProfile: model.growthProfile,
            revenueStreams: model.revenueStreams,
            opex: model.opex,
            projections: model.projections,
            summary: model.summary,
            assumptions: model.assumptions,
            loanSchedule: model.loanSchedule,
            // Backwards-compatible: Excel injection actions
            actions: model.actions,
        });

    } catch (error) {
        console.error('[generate-model] Error:', error);
        return NextResponse.json({
            error: error.message || 'Failed to generate financial model',
        }, { status: 500 });
    }
}

/**
 * GET /api/generate-model?industry=saas
 * Returns the default template for a given industry.
 * Useful for previewing what the engine will produce.
 */
export async function GET(req) {
    try {
        const { searchParams } = new URL(req.url);
        const industry = searchParams.get('industry') || 'consulting';

        const model = generateFinancialModel({ industry });

        return NextResponse.json({
            success: true,
            industry: model.industry,
            industryName: model.industryName,
            icon: model.industryIcon,
            kpis: model.kpis,
            revenueStreams: model.revenueStreams,
            opex: model.opex,
            projections: model.projections,
            summary: model.summary,
        });
    } catch (error) {
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}
