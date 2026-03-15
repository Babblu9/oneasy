function inferBusinessType(kg) {
    const text = JSON.stringify(kg?.basics || {}).toLowerCase();
    if (text.includes("edtech") || text.includes("course") || text.includes("education")) return "edtech";
    if (text.includes("saas") || text.includes("software")) return "saas";
    if (text.includes("health")) return "healthcare";
    if (text.includes("consult")) return "consulting";
    if (text.includes("ecom") || text.includes("retail")) return "ecommerce";
    return "business";
}

function nextMissingStep(kg) {
    const basics = kg?.basics || {};
    const rev = Array.isArray(kg?.revP1) ? kg.revP1 : [];
    const opex = Array.isArray(kg?.opexP1) ? kg.opexP1 : [];
    const streamHeaders = rev.filter((s) => String(s?.header || "").trim()).length;
    const substreams = rev.flatMap((s) => s?.items || []).filter((i) => String(i?.sub || "").trim()).length;
    const priced = rev.flatMap((s) => s?.items || []).filter((i) => Number(i?.price) > 0).length;
    const volumes = rev.flatMap((s) => s?.items || []).filter((i) => Number(i?.qty) > 0).length;
    const opexItems = opex.flatMap((s) => s?.items || []).filter((i) => String(i?.sub || "").trim()).length;

    if (!String(basics.legalName || basics.tradeName || "").trim()) return "company_name";
    if (!String(basics.description || "").trim()) return "business_type";
    if (!streamHeaders) return "main_streams";
    if (!substreams) return "substreams";
    if (!String(basics.burningDesire || "").trim()) return "target_customers";
    if (!priced) return "pricing";
    if (!volumes) return "volumes";
    if (!opexItems) return "opex";
    if (!String(basics.startDateP1 || "").trim()) return "launch_funding";
    return "review";
}

function summarizeKnowledgeGraph(kg) {
    const basics = kg?.basics || {};
    const rev = Array.isArray(kg?.revP1) ? kg.revP1 : [];
    const opex = Array.isArray(kg?.opexP1) ? kg.opexP1 : [];

    return {
        basics: {
            tradeName: basics.tradeName || basics.legalName || '',
            description: basics.description || '',
            customers: basics.burningDesire || '',
            launchDate: basics.startDateP1 || '',
            pitchDeck: basics.pitchDeck || '',
        },
        revenue: rev
            .filter((s) => String(s?.header || '').trim())
            .slice(0, 5)
            .map((s) => ({
                header: s.header,
                items: (s.items || [])
                    .filter((i) => String(i?.sub || '').trim())
                    .slice(0, 5)
                    .map((i) => ({ sub: i.sub, qty: i.qty || 0, price: i.price || 0 }))
            })),
        opex: opex
            .filter((s) => String(s?.header || '').trim())
            .slice(0, 5)
            .map((s) => ({
                header: s.header,
                items: (s.items || [])
                    .filter((i) => String(i?.sub || '').trim())
                    .slice(0, 7)
                    .map((i) => ({ sub: i.sub, qty: i.qty || 0, cost: i.cost || 0 }))
            })),
        nextMissingStep: nextMissingStep(kg),
    };
}

export function getFinancialExtractorPrompt(kg, messageCount = 0) {
    return `You are a Financial Data Extractor. Extract only facts the user explicitly stated.

CURRENT KNOWLEDGE GRAPH:
${JSON.stringify(summarizeKnowledgeGraph(kg))}

Rules:
- Do not invent values.
- If the user shares a website URL, pitch deck, or business materials, capture those in basics fields if possible.
- Do not estimate revenue, opex, funding, growth, or staffing.
- Greetings like "hi" or "hello" contain no financial meaning.
- If the user only greets, return null for all fields and keep stage as discovery.
- Suggested stage should follow the next missing data area.

Extract:
1. Basics: legalName, tradeName, description, burningDesire, startDateP1, promoters, location, pitchDeck, website.
2. Revenue streams and items with qty, price, growth if explicitly given.
3. OPEX streams and items with qty, cost, growth if explicitly given.
4. Funding only if explicitly given.

Return JSON matching the schema exactly.`;
}

export function getFinancialConsultantPrompt(kg, stage, messageCount = 0) {
    const businessType = inferBusinessType(kg);
    const nextStep = nextMissingStep(kg);

    return `You are Fina, an expert financial strategist guiding a user through building a professional financial model.

CURRENT STAGE: ${stage}
MEANINGFUL USER MESSAGE COUNT: ${messageCount}
BUSINESS TYPE: ${businessType}
NEXT MISSING STEP: ${nextStep}
CURRENT KNOWLEDGE GRAPH:
${JSON.stringify(summarizeKnowledgeGraph(kg))}

Behavior rules:
- Be conversational, professional, and precise.
- Acknowledge what was just captured if it's the first time you're seeing it.
- Ask only one (or two related) questions at a time to keep the momentum.
- If the user says "okay" or "no changes", move to the next logical step immediately.
- If the user asks for suggestions, give 3-4 industry-specific ideas for ${businessType}.
- Never skip ahead to funding or review if basics or revenue structure are entirely missing.
- Keep responses under 80 words.

Discovery Flow (If stage is discovery, revenue_setup, etc.):
1. Validate the current data in the Knowledge Graph.
2. If NEXT MISSING STEP is not 'review', ask for that specific missing info (Name, Type, Streams, Pricing, etc.).
3. If you have enough info to suggest a starter model (streams/opex), tell the user you've drafted it based on their business type and ask for refinement (pricing/volumes).

Review Flow (If stage is review):
1. Congratulate the user on completing the basic structure.
2. Briefly summarize the key highlights (e.g., "We've captured your ${businessType} with revenue from Retainers and a launch next month").
3. Ask if they want to refine specific numbers, look at the P&L, or prepare for export.

Never emit [DATA] tags. Use markdown for bolding key terms.`;
}
