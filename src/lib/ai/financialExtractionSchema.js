import { z } from "zod";

export const financialExtractionSchema = z.object({
    basics: z.object({
        legalName: z.string().nullable(),
        tradeName: z.string().nullable(),
        description: z.string().nullable(),
        burningDesire: z.string().nullable(),
        startDateP1: z.string().nullable(),
        promoters: z.string().nullable(),
        location: z.string().nullable(),
    }).nullable(),

    revenue_streams: z.array(z.object({
        header: z.string(),
        items: z.array(z.object({
            sub: z.string(),
            qty: z.number(),
            price: z.number(),
            gY1: z.number().nullable(),
            gY2: z.number().nullable(),
            gY3: z.number().nullable(),
            gY4: z.number().nullable(),
            gY5: z.number().nullable(),
        }))
    })).nullable(),

    opex_streams: z.array(z.object({
        header: z.string(),
        items: z.array(z.object({
            sub: z.string(),
            qty: z.number(),
            cost: z.number(),
            gY1: z.number().nullable(),
            gY2: z.number().nullable(),
            gY3: z.number().nullable(),
            gY4: z.number().nullable(),
            gY5: z.number().nullable(),
        }))
    })).nullable(),

    funding: z.object({
        promoterContrib: z.number().nullable(),
        termLoan: z.number().nullable(),
        wcLoan: z.number().nullable(),
        loan1: z.object({
            amount: z.number(),
            duration: z.number(),
            rate: z.number(),
            startDate: z.string(),
        }).nullable(),
    }).nullable(),

    suggested_stage: z.enum([
        "discovery", 
        "revenue_setup", 
        "opex_setup", 
        "funding_setup", 
        "review", 
        "model_ready"
    ]).nullable(),
    
    should_reset: z.boolean().nullable(),
});
