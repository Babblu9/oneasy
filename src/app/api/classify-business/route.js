import { NextResponse } from 'next/server';
import { streamText } from 'ai';
import { createOpenAI } from '@ai-sdk/openai';
import { classifyByKeywords, getAllTemplates } from '@/lib/templateRegistry';

const gateway = createOpenAI({
    apiKey: process.env.AI_GATEWAY_API_KEY || '',
    baseURL: 'https://ai-gateway.vercel.sh/v1',
});

/**
 * POST /api/classify-business
 * Body: { description: "I run an EdTech platform selling online courses" }
 * Returns: { templateId, template, confidence, method }
 *
 * Strategy:
 * 1. Fast keyword match first (local, no AI call)
 * 2. If confidence is low, use AI for classification
 */
export async function POST(req) {
    try {
        const { description } = await req.json();

        if (!description || String(description).trim().length < 3) {
            return NextResponse.json({ error: 'description is required' }, { status: 400 });
        }

        // Step 1: Fast keyword classification
        const keywordResult = classifyByKeywords(description);

        if (keywordResult.score >= 2) {
            // High confidence keyword match
            return NextResponse.json({
                templateId: keywordResult.templateId,
                template: keywordResult.template,
                confidence: Math.min(0.95, 0.6 + keywordResult.score * 0.1),
                method: 'keyword'
            });
        }

        // Step 2: AI classification for ambiguous cases
        const templates = getAllTemplates();
        const templateList = templates.map(t => `- ${t.id}: ${t.name} (${t.description})`).join('\n');

        const result = await streamText({
            model: gateway(process.env.AI_GATEWAY_MODEL || 'anthropic/claude-sonnet-4-5'),
            system: `You are a business classifier. Given a business description, return ONLY the template ID that best matches it. Return ONLY the ID string, nothing else.

Available templates:
${templateList}

Rules:
- Return ONLY one of these IDs: healthcare, edtech, saas, ecommerce, pharma, consulting, manufacturing
- If unsure, return "consulting" as the default
- DO NOT return any explanation or extra text`,
            messages: [{ role: 'user', content: `Classify this business: "${description}"` }],
            maxTokens: 20,
            temperature: 0
        });

        let templateId = '';
        for await (const chunk of result.textStream) {
            templateId += chunk;
        }
        templateId = templateId.trim().toLowerCase().replace(/[^a-z]/g, '');

        // Validate the returned ID
        const validIds = templates.map(t => t.id);
        if (!validIds.includes(templateId)) {
            templateId = keywordResult.templateId; // Fall back to keyword result
        }

        const matched = templates.find(t => t.id === templateId);

        return NextResponse.json({
            templateId,
            template: matched,
            confidence: 0.85,
            method: 'ai'
        });

    } catch (error) {
        console.error('classify-business error:', error);
        return NextResponse.json({ error: 'Classification failed', details: error.message }, { status: 500 });
    }
}

/** GET /api/classify-business — returns all templates (for UI template picker) */
export async function GET() {
    return NextResponse.json({ templates: getAllTemplates() });
}
