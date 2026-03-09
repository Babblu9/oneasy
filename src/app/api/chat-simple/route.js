import { NextResponse } from 'next/server';
import { generateText } from 'ai';
import { createOpenAI } from '@ai-sdk/openai';

const gateway = createOpenAI({
    apiKey: process.env.AI_GATEWAY_API_KEY || '',
    baseURL: 'https://ai-gateway.vercel.sh/v1',
});

export const maxDuration = 60;

export async function POST(req) {
    try {
        const body = await req.json();
        const  = String(body?.system || '').trim();
        const prompt = String(body?.prompt || '').trim();
        const messages = Array.isArray(body?.messages) ? body.messages : [];

        if (!prompt && messages.length === 0) {
            return NextResponse.json({ error: 'prompt or messages is required' }, { status: 400 });
        }

        const history = messages
            .filter(m => m && (m.role === 'user' || m.role === 'assistant'))
            .slice(-12)
            .map(m => `${m.role.toUpperCase()}: ${String(m.text || '').trim()}`)
            .join('\n');

        const finalPrompt = [
            history ? `Conversation:\n${history}` : '',
            prompt || '',
        ].filter(Boolean).join('\n\n');

        const result = await generateText({
            model: gateway(process.env.AI_GATEWAY_MODEL || 'anthropic/claude-sonnet-4-5'),
            system: system || undefined,
            prompt: finalPrompt,
            temperature: 0.2,
        });

        return NextResponse.json({
            success: true,
            text: result?.text || '',
        });
    } catch (error) {
        console.error('chat-simple error:', error);
        return NextResponse.json({ error: 'Chat request failed.' }, { status: 500 });
    }
}
