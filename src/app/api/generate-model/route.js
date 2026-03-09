import { NextResponse } from 'next/server';
import { generateModelRequestSchema } from '@/lib/modelSchemas';
import { generateDeterministicModel } from '@/lib/deterministicModelEngine';

export async function POST(req) {
    try {
        const body = await req.json();
        const parsed = generateModelRequestSchema.safeParse(body || {});

        if (!parsed.success) {
            return NextResponse.json({
                error: 'Invalid input',
                details: parsed.error.flatten(),
            }, { status: 400 });
        }

        const result = generateDeterministicModel(parsed.data);

        return NextResponse.json({
            success: true,
            model: {
                industry: result.industry,
                monthlyRevenue: result.monthlyRevenue,
                monthlyOpex: result.monthlyOpex,
                monthlyNet: result.monthlyNet,
                growthRate: result.growthRate,
                funding: result.funding,
            },
            actions: result.actions,
        });
    } catch (error) {
        return NextResponse.json({ error: error.message || 'Failed to generate deterministic model' }, { status: 500 });
    }
}
