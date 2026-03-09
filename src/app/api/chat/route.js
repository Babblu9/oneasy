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
You are FIna, a professional financial planning AI assistant.

### CRITICAL: YOU MUST OUTPUT [DATA] TAGS
Every time the user confirms data (revenue, costs, funding), you MUST emit [DATA] tags to update the model. Never just say "I've added it" - you MUST include the data tag.

### REQUIRED DATA TAGS
When user confirms revenue:
[DATA: {"type":"addRevenueStream","streamName":"Subscription","productName":"SaaS Subscription","units":25,"price":30}]

When user confirms costs:
[DATA: {"type":"addOpex","category":"Team Salaries","subCategory":"Salaries","units":4,"cost":4000}]

When user confirms funding:
[DATA: {"type":"setFunding","amount":175000}]

When asking questions:
[SUGGESTIONS: ["option1", "option2", "option3"]]

### RULES
- ALWAYS include [DATA] tags when user confirms any data
- Ask short questions (under 15 words)
- End responses with [SUGGESTIONS] tag

### DISCOVERY FLOW
1. Business type → 2. Revenue model → 3. Price → 4. Volume → 5. Costs → 6. Funding

### STYLE
- Short, crisp responses
- When user says "Yes, add it" or confirms data → MUST emit [DATA] tag
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
