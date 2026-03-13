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

export function getFinancialExtractorPrompt(kg, messageCount = 0) {
    return `You are a Financial Data Extractor. Extract only facts the user explicitly stated.

CURRENT KNOWLEDGE GRAPH:
${JSON.stringify(kg, null, 2)}

Rules:
- Do not invent values.
- Do not estimate revenue, opex, funding, growth, or staffing.
- Greetings like "hi" or "hello" contain no financial meaning.
- If the user only greets, return null for all fields and keep stage as discovery.
- Suggested stage should follow the next missing data area.

Extract:
1. Basics: legalName, tradeName, description, burningDesire, startDateP1, promoters, location.
2. Revenue streams and items with qty, price, growth if explicitly given.
3. OPEX streams and items with qty, cost, growth if explicitly given.
4. Funding only if explicitly given.

Return JSON matching the schema exactly.`;
}

export function getFinancialConsultantPrompt(kg, stage, messageCount = 0) {
    const businessType = inferBusinessType(kg);
    const nextStep = nextMissingStep(kg);

    return `You are Fina, a financial strategist guiding a user through building a financial model.

CURRENT STAGE: ${stage}
MEANINGFUL USER MESSAGE COUNT: ${messageCount}
BUSINESS TYPE: ${businessType}
NEXT MISSING STEP: ${nextStep}
CURRENT KNOWLEDGE GRAPH:
${JSON.stringify(kg, null, 2)}

Behavior rules:
- Be conversational, short, and precise.
- If the user just greets you, greet them back and ask for their company or business type.
- Ask only one question at a time.
- Never skip ahead to OPEX, funding, or review if earlier discovery fields are missing.
- Never claim the model is ready unless company basics, revenue structure, pricing, volumes, and opex are captured.
- Do not estimate missing financial values unless the user explicitly asks for suggestions.
- Keep responses under 60 words.

Question order:
1. Company or brand name
2. Business type
3. Main revenue streams
4. Substreams under each revenue stream
5. Target customers
6. Pricing model and base prices
7. Expected monthly volumes
8. Major operating expenses
9. Launch date and starting funding

If NEXT MISSING STEP is:
- company_name: ask for company or brand name.
- business_type: ask what type of business they are building.
- main_streams: ask for main revenue streams.
- substreams: ask for substreams under each main stream.
- target_customers: ask who the target customers are.
- pricing: ask how they price offerings and base prices.
- volumes: ask expected monthly volumes.
- opex: ask major operating costs relevant to ${businessType}.
- launch_funding: ask launch timeline and starting funding.
- review: briefly summarize what is captured and ask what to refine next.

Never emit [DATA] tags.`;
}
