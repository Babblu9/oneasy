import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { streamText, generateObject } from "ai";
import { createOpenAI } from '@ai-sdk/openai';
import { financialExtractionSchema } from "@/lib/ai/financialExtractionSchema";
import { mergeFinancialExtraction, getFinancialCompletionPercentage } from "@/lib/ai/financialAgent";
import { getFinancialExtractorPrompt, getFinancialConsultantPrompt } from "@/lib/ai/financialPrompts";
import { classifyIndustry, getGrowthRates } from "@/lib/engine/industryTemplates";
import { buildHyperFormulaFromWorkbook, extractDashboardMetrics } from '@/lib/hyperFormulaEngine';

const gateway = createOpenAI({
    apiKey: process.env.AI_GATEWAY_API_KEY || '',
    baseURL: 'https://ai-gateway.vercel.sh/v1',
});

export const maxDuration = 60;

const WORK_EXCEL = '/tmp/active_working.xlsx';
const SOURCE_EXCEL = path.join(process.cwd(), 'excel-templates', 'active_working.xlsx');

function ensureWorkingFile() {
    if (!fs.existsSync(WORK_EXCEL)) {
        if (!fs.existsSync(SOURCE_EXCEL)) throw new Error(`Source template not found: ${SOURCE_EXCEL}`);
        fs.copyFileSync(SOURCE_EXCEL, WORK_EXCEL);
    }
}

async function getWorkbookMetrics() {
    try {
        ensureWorkingFile();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(WORK_EXCEL);
        const hfContext = buildHyperFormulaFromWorkbook(workbook);
        return extractDashboardMetrics(hfContext);
    } catch (error) {
        console.error('workbook metrics error:', error);
        return null;
    }
}

function isSheetReadIntent(text) {
    return /read it from the sheet|from the sheet|sheet itself|is that p\s*&\s*l correct|is that p&l correct|profit\s*&\s*loss correct|profit and loss correct|check p\s*&\s*l|check p&l/i.test(String(text || ''));
}

function buildWorkbookReply(metrics) {
    if (!metrics) {
        return "I could not read the active workbook summary right now. Retry once the sheet has finished recalculating.";
    }

    const currency = (value) => {
        const num = Number(value || 0);
        if (!num) return '₹0';
        if (Math.abs(num) >= 10000000) return `₹${(num / 10000000).toFixed(2)} Cr`;
        if (Math.abs(num) >= 100000) return `₹${(num / 100000).toFixed(2)} L`;
        return `₹${Math.round(num).toLocaleString('en-IN')}`;
    };

    return `I read the current workbook summary. Year 1 revenue is ${currency(metrics.year1Revenue)}, Year 2 revenue is ${currency(metrics.year2Revenue)}, EBITDA is ${currency(metrics.ebitda)}, and the model break-even is around month ${metrics.breakEvenMonth}. Based on the live sheet, the P&L is using the current assumptions already in the workbook.`;
}

function extractUrls(text) {
    return String(text || '').match(/https?:\/\/[^\s]+/g) || [];
}

function looksLikeShortName(text) {
    const clean = String(text || '').trim();
    if (!clean) return false;
    if (/https?:\/\//i.test(clean)) return false;
    return clean.split(/\s+/).length <= 4 && clean.length <= 40;
}

function hasMeaningfulHeader(value) {
    const v = String(value || '').trim().toLowerCase();
    if (!v) return false;
    return ![
        'revenue stream 1',
        'phase 2 revenue',
        'operating expenses 1',
        'capital expenditure 1'
    ].includes(v);
}

function hasMeaningfulSub(value) {
    const v = String(value || '').trim().toLowerCase();
    if (!v) return false;
    return !['service', 'general', 'sub service...', 'expense item', 'item...'].includes(v);
}


function buildTemplateRevenueGroups(template, growth) {
    const groups = new Map();
    for (const item of template?.revenueStreams || []) {
        const header = String(item.stream || '').trim();
        if (!header) continue;
        if (!groups.has(header)) groups.set(header, []);
        const items = groups.get(header);
        if (items.length >= 5) continue;
        items.push({
            sub: item.name,
            qty: item.quantity,
            price: item.price,
            gY1: growth.y1,
            gY2: growth.y2,
            gY3: growth.y3,
            gY4: growth.y4,
            gY5: growth.y5,
        });
    }
    return Array.from(groups.entries()).slice(0, 5).map(([header, items]) => ({ header, items }));
}

function buildTemplateOpexGroups(template, growth) {
    const groups = new Map();
    for (const item of template?.opex || []) {
        const header = String(item.category || '').trim();
        if (!header) continue;
        if (!groups.has(header)) groups.set(header, []);
        const items = groups.get(header);
        if (items.length >= 7) continue;
        items.push({
            sub: item.name,
            qty: 1,
            cost: item.monthlyCost,
            gY1: growth.y1,
            gY2: growth.y2,
            gY3: growth.y3,
            gY4: growth.y4,
            gY5: growth.y5,
        });
    }
    return Array.from(groups.entries()).slice(0, 5).map(([header, items]) => ({ header, items }));
}

function inferBusinessType(kg) {
    const text = JSON.stringify(kg?.basics || {}).toLowerCase();
    if (text.includes("chartered") || text.includes("ca firm") || text.includes("accounting")) return "ca_firm";
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
    const streamHeaders = rev.filter((s) => hasMeaningfulHeader(s?.header)).length;
    const substreams = rev.flatMap((s) => s?.items || []).filter((i) => hasMeaningfulSub(i?.sub)).length;
    const priced = rev.flatMap((s) => s?.items || []).filter((i) => Number(i?.price) > 0).length;
    const volumes = rev.flatMap((s) => s?.items || []).filter((i) => Number(i?.qty) > 0).length;
    const opexItems = opex.flatMap((s) => s?.items || []).filter((i) => hasMeaningfulSub(i?.sub)).length;

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

function buildDiscoveryReply(kg, lastUserMessage, effectiveStep = null) {
    const text = String(lastUserMessage || "").trim().toLowerCase();
    const greetings = ["hi", "hello", "hey", "hi fina", "hello fina", "hey fina"];
    if (greetings.includes(text)) {
        return "Hello! Tell me more about yourself and the business you are working on. You can also share your website URL, pitch deck, any supporting documents, and anything else you are comfortable sharing.";
    }

    if (effectiveStep === "company_name") {
        return "Tell me more about yourself and the business you are working on. You can also share your website URL, pitch deck, any supporting documents, and anything else you are comfortable sharing.";
    }
    if (effectiveStep === "business_type") {
        return "I have drafted likely revenue streams and price points from the market. Who are your target customers?";
    }
    if (effectiveStep === "target_customers") {
        return "I have drafted standard price points. What pricing model or pricing changes do you want to keep?";
    }
    if (effectiveStep === "pricing") {
        return "What monthly volumes do you expect in Year 1?";
    }
    if (effectiveStep === "volumes") {
        return "I have added typical operating costs. What cost items should we adjust or add?";
    }
    if (effectiveStep === "opex") {
        return "What is your planned launch date and starting funding?";
    }
    if (effectiveStep === "launch_funding") {
        return "I have drafted the base model from market benchmarks. What do you want to refine first: revenue, opex, project cost, or P&L?";
    }

    const step = nextMissingStep(kg);
    const businessType = inferBusinessType(kg);

    if (step === "company_name") return "Tell me more about yourself and the business you are working on. You can also share your website URL, pitch deck, any supporting documents, and anything else you are comfortable sharing.";
    if (step === "business_type") return "What type of business are you building? Tell me in one line.";
    if (step === "main_streams") return `What are the main revenue streams for your ${businessType.replace(/_/g, " ")}?`;
    if (step === "substreams") return "I have drafted likely revenue streams and price points from the market. Who are your target customers?";
    if (step === "target_customers") return "Who are your target customers?";
    if (step === "pricing") return "I have drafted standard price points. What pricing model or pricing changes do you want to keep?";
    if (step === "volumes") return "What monthly volumes do you expect in Year 1?";
    if (step === "opex") return "I have added typical operating costs. What cost items should we adjust or add?";
    if (step === "launch_funding") return "What is your planned launch date and starting funding?";
    return "I have drafted the base model from market benchmarks. What do you want to refine first: revenue, opex, project cost, or P&L?";
}

function inferAnsweredStepFromConversation(messages = []) {
    const lastAssistant = [...messages].reverse().find((m) => m.role === "assistant");
    const text = String(lastAssistant?.content || "").toLowerCase();
    if (!text) return null;
    if (text.includes("name of your company") || text.includes("company or brand")) return "company_name";
    if (text.includes("what type of business")) return "business_type";
    if (text.includes("main revenue streams")) return "main_streams";
    if (text.includes("substreams")) return "substreams";
    if (text.includes("target customers")) return "target_customers";
    if (text.includes("pricing model") || text.includes("base prices") || text.includes("pricing changes")) return "pricing";
    if (text.includes("monthly volumes")) return "volumes";
    if (text.includes("operating costs") || text.includes("cost items")) return "opex";
    if (text.includes("launch date") || text.includes("starting funding")) return "launch_funding";
    return null;
}

export async function POST(req) {
    try {
        const {
            messages,
            knowledgeGraph,
            stage,
        } = await req.json();

        const lastUserMessage = messages[messages.length - 1];
        if (!lastUserMessage || lastUserMessage.role !== "user") {
            return new Response("No user message found", { status: 400 });
        }

        const isMeaningfulUserMessage = (content) => {
            const text = String(content || "").trim().toLowerCase();
            if (!text) return false;
            const trivial = [
                "hi", "hello", "hey", "hi fina", "hello fina", "hey fina",
                "ok", "okay", "cool", "yes", "no", "thanks", "thank you"
            ];
            if (trivial.includes(text)) return false;
            return text.length > 3;
        };

        // ── Phase 1: Extractor (Worker) ────────────────────────────
        const contextMessages = messages.slice(-3); // Context window
        const userMessageCount = messages.filter(m => m.role === "user" && isMeaningfulUserMessage(m.content)).length;

        const currentStep = nextMissingStep(knowledgeGraph);
        const conversationalStep = inferAnsweredStepFromConversation(messages.slice(0, -1));
        const effectiveStep = conversationalStep || currentStep;

        let extractedData = null;
        try {
            const extraction = await generateObject({
                model: gateway(process.env.AI_GATEWAY_MODEL || 'gpt-4o-mini'),
                schema: financialExtractionSchema,
                prompt: `${getFinancialExtractorPrompt(knowledgeGraph, userMessageCount)}

RECENT CONVERSATION:
${contextMessages.map((m) => `${m.role}: ${m.content}`).join("\n")}

USER'S LATEST MESSAGE:
${lastUserMessage.content}

Return structured JSON. Only include changed or new fields.`,
            });
            extractedData = extraction.object;
            if (userMessageCount >= 5) {
                console.log("[AGENT] Tipping point extraction:", JSON.stringify(extractedData, null, 2));
            }
        } catch (e) {
            console.error("Extraction error:", e);
        }

        // Deterministically capture the user's answer for the currently-missing field.
        if (isMeaningfulUserMessage(lastUserMessage.content)) {
            const latestText = String(lastUserMessage.content || "").trim();
            if (effectiveStep === "company_name") {
                const urls = extractUrls(latestText);
                const shortName = looksLikeShortName(latestText) ? latestText : (extractedData?.basics?.tradeName || extractedData?.basics?.legalName || "");
                extractedData = {
                    ...extractedData,
                    basics: {
                        ...extractedData?.basics,
                        legalName: extractedData?.basics?.legalName || shortName,
                        tradeName: extractedData?.basics?.tradeName || shortName,
                        pitchDeck: extractedData?.basics?.pitchDeck || urls[0] || "",
                        description: extractedData?.basics?.description || (!looksLikeShortName(latestText) ? latestText : ""),
                    },
                    suggested_stage: extractedData?.suggested_stage || "discovery",
                };
            }
            if (effectiveStep === "business_type") {
                extractedData = {
                    ...extractedData,
                    basics: {
                        ...extractedData?.basics,
                        description: latestText,
                    },
                    suggested_stage: extractedData?.suggested_stage || "revenue_setup",
                };
            }
            if (effectiveStep === "target_customers") {
                extractedData = {
                    ...extractedData,
                    basics: {
                        ...extractedData?.basics,
                        burningDesire: latestText,
                    },
                    suggested_stage: extractedData?.suggested_stage || "pricing",
                };
            }
            if (effectiveStep === "launch_funding") {
                extractedData = {
                    ...extractedData,
                    basics: {
                        ...extractedData?.basics,
                        startDateP1: extractedData?.basics?.startDateP1 || latestText,
                    },
                    suggested_stage: extractedData?.suggested_stage || "review",
                };
            }
        }

        const hasRevenue = Array.isArray(knowledgeGraph?.revP1) && knowledgeGraph.revP1.some(
            (stream) => hasMeaningfulHeader(stream?.header) || (stream?.items || []).some((item) => hasMeaningfulSub(item?.sub))
        );
        const hasOpex = Array.isArray(knowledgeGraph?.opexP1) && knowledgeGraph.opexP1.some(
            (stream) => hasMeaningfulHeader(stream?.header) || (stream?.items || []).some((item) => hasMeaningfulSub(item?.sub))
        );

        const latestBusinessHint = [
            lastUserMessage.content,
            extractedData?.basics?.description,
            extractedData?.basics?.burningDesire,
        ].filter(Boolean).join(" ");

        if (latestBusinessHint && (!hasRevenue || !hasOpex)) {
            const { template, score } = classifyIndustry(latestBusinessHint);
            if (template && score > 0) {
                const growth = getGrowthRates(template.growthProfile);
                extractedData = {
                    ...extractedData,
                    basics: {
                        ...extractedData?.basics,
                        description: extractedData?.basics?.description || template.name,
                    },
                    revenue_streams: extractedData?.revenue_streams?.length ? extractedData.revenue_streams : buildTemplateRevenueGroups(template, growth),
                    opex_streams: extractedData?.opex_streams?.length ? extractedData.opex_streams : buildTemplateOpexGroups(template, growth),
                    suggested_stage: extractedData?.suggested_stage || "revenue_setup",
                };
            }
        }

        // Merge and Progress
        const updatedKG = extractedData
            ? mergeFinancialExtraction(knowledgeGraph, extractedData)
            : knowledgeGraph;

        let updatedStage = stage;
        if (extractedData?.suggested_stage) {
            const stages = ["discovery", "revenue_setup", "opex_setup", "funding_setup", "review", "model_ready"];
            const currentIdx = stages.indexOf(stage);
            const suggestedIdx = stages.indexOf(extractedData.suggested_stage);
            if (suggestedIdx > currentIdx) {
                updatedStage = extractedData.suggested_stage;
            }
        }

        // Automatic stage promotion if nextMissingStep is review
        if (updatedStage !== "model_ready" && nextMissingStep(updatedKG) === "review") {
            updatedStage = "review";
        }

        const deterministicReplyHint = buildDiscoveryReply(updatedKG, lastUserMessage.content, isMeaningfulUserMessage(lastUserMessage.content) ? effectiveStep : null);

        const makeDirectResponse = (replyText) => {
            const stream = new ReadableStream({
                start(controller) {
                    const encoder = new TextEncoder();
                    try {
                        if (extractedData) controller.enqueue(encoder.encode(`2:${JSON.stringify(extractedData)}\n`));
                        controller.enqueue(encoder.encode(`0:${JSON.stringify(replyText)}\n`));
                        controller.close();
                    } catch (error) {
                        controller.error(error);
                    }
                },
            });

            return new Response(stream, {
                headers: {
                    "Content-Type": "text/plain; charset=utf-8",
                    "X-Stage": updatedStage || "discovery",
                    "X-Completion": String(getFinancialCompletionPercentage(updatedKG) || 0),
                    "Access-Control-Expose-Headers": "X-Stage, X-Completion"
                },
            });
        };

        if (isSheetReadIntent(lastUserMessage.content)) {
            const workbookMetrics = await getWorkbookMetrics();
            return makeDirectResponse(buildWorkbookReply(workbookMetrics));
        }

        const shouldBypassConsultant = updatedStage !== 'review' && updatedStage !== 'model_ready';
        if (shouldBypassConsultant && deterministicReplyHint) {
            return makeDirectResponse(deterministicReplyHint);
        }

        // ── Phase 2: Consultant (Manager) ─────────────────────────
        const result = streamText({
            model: gateway(process.env.AI_GATEWAY_MODEL_QUALITY || 'gpt-4o-mini'),
            system: getFinancialConsultantPrompt(updatedKG, updatedStage, userMessageCount),
            messages: messages.map((m) => ({
                role: m.role,
                content: m.content,
            })),
        });

        // Create a manual stream to format as Vercel AI Data Stream Protocol
        const stream = new ReadableStream({
            async start(controller) {
                const encoder = new TextEncoder();
                try {
                    // Send extraction data first if available (Prefix 2: is our custom 'data' part)
                    if (extractedData) {
                        controller.enqueue(encoder.encode(`2:${JSON.stringify(extractedData)}\n`));
                    }

                    for await (const text of result.textStream) {
                        if (text) {
                            controller.enqueue(encoder.encode(`0:${JSON.stringify(text)}\n`));
                        }
                    }
                    controller.close();
                } catch (error) {
                    console.error("Stream processing loop error:", error);
                    controller.error(error);
                }
            },
        });

        // Headers for Client State Sync
        const customHeaders = {
            "X-Stage": updatedStage || "discovery",
            "X-Completion": String(getFinancialCompletionPercentage(updatedKG) || 0),
            "Access-Control-Expose-Headers": "X-Stage, X-Completion"
        };

        return new Response(stream, {
            headers: {
                "Content-Type": "text/plain; charset=utf-8",
                ...customHeaders,
            },
        });

    } catch (error) {
        console.error("Financial Chat API error:", error);
        return new Response(JSON.stringify({ error: error.message }), { status: 500 });
    }
}
