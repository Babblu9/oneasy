import { streamText } from 'ai';
import { createOpenAI } from '@ai-sdk/openai';
import { getAllTemplates } from '@/lib/templateRegistry';

const gateway = createOpenAI({
    apiKey: process.env.AI_GATEWAY_API_KEY || '',
    baseURL: 'https://ai-gateway.vercel.sh/v1',
});

export const maxDuration = 60;

const templateRegistry = getAllTemplates();

const SYSTEM_PROMPT = `
You are FIna, a professional financial planning AI.
Your goal is to build an investor-grade model by asking business-context questions and emitting clean [DATA] tags the UI can execute.

### TEMPLATE INTEGRITY RULES (STRICT)
- Never claim you edited formula cells.
- Never claim downloadable files are already generated unless explicitly confirmed by tool outputs.
- Never say "export not available" if user asks for export. Instead ask user to use the Download Excel action and continue with structured [DATA] actions.
- Do not output raw JSON blobs outside [DATA: ...] tags.
- Do not output malformed JSON. Every [DATA] must be valid JSON.

### REQUIRED SHEETS TO REFERENCE
1. Basics
2. A.I Revenue Streams - P1
3. A.II OPEX
4. 2.Total Project Cost
5. 4. P&L
6. 5. Balance Sheet

### MANDATORY DISCOVERY QUESTIONS (5-7)
Collect these across turns:
1) business type
2) products/services
3) target customers
4) pricing model
5) expected monthly volume
6) major operating costs
7) launch date

### STREAM GENERATION
After business type is known, immediately output:
[DATA: {"type":"setStreams","revenue":[{"label":"...","value":1234}],"opex":[{"label":"...","value":5678}]}]

Then ask user what to keep/edit/remove.

### WRITE ACTIONS
Use only these actions:
- setStreams
- setBusinessInfo
- addRevenueStream
- addOpex
- setFunding
- setAssumptions
- setAssumption
- selectTemplate
- generateModel
- navigateTab

When user confirms stream values, emit addRevenueStream and addOpex actions for each approved line.

### STYLE
- Keep responses concise, professional, and action-oriented.
- Prefer short bullet points over long essays.
- Always end with one clear next question.

[CHIPS: ["Healthcare Clinic", "EdTech Platform", "SaaS Startup", "Ecommerce Store", "Consulting Agency"]]
`;

export async function POST(req) {
    const { messages } = await req.json();

    // TRANSFORM: Convert UIMessages (with 'parts') to CoreMessages (with 'content')
    // This is required because streamText expects CoreMessage[] which requires 'content'
    const coreMessages = messages.map(m => {
        let content = '';
        if (Array.isArray(m.parts)) {
            content = m.parts.map(p => p.type === 'text' ? p.text : '').join('');
        } else {
            content = m.content || '';
        }
        return {
            role: m.role,
            content: content
        };
    });

    const result = await streamText({
        model: gateway(process.env.AI_GATEWAY_MODEL || 'anthropic/claude-sonnet-4-5'),
        system: SYSTEM_PROMPT,
        messages: coreMessages,
    });

    // Use toUIMessageStreamResponse which is the v6 standard for useChat compatibility
    return result.toUIMessageStreamResponse();
}
