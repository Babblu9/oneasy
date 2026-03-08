/**
 * Model Catalog
 * =============
 * Parses model_inputs.json and produces structured catalogs for:
 *   1. AI system prompt (condensed product/OPEX listing)
 *   2. Excel cell map (auto-generated row mappings)
 *
 * Imported at build time by Next.js — runs server-side only.
 */

import modelInputs from '../../model_inputs.json';

// ── Revenue Product Catalog ─────────────────────────────────────────────────
// Flattens the nested revenue_streams into a flat array of products
// with their stream/sub-stream/product names, cell refs, prices, etc.

function buildRevenueCatalog() {
    const products = [];
    const { streams } = modelInputs.revenue_streams;

    for (const [streamName, subStreams] of Object.entries(streams)) {
        for (const [subStreamName, items] of Object.entries(subStreams)) {
            for (const item of items) {
                products.push({
                    stream: streamName,
                    subStream: subStreamName,
                    product: item.product,
                    id: item.id,
                    qtyFormula: item.qty_formula || '',
                    qtyNote: item.qty_note || '',
                    price: item.price_inr,
                    growthRates: item.growth_rates_pct || {},
                    cellRefs: item.cell_refs || {},
                    branchScaled: (item.qty_formula || '').includes('H7'),
                });
            }
        }
    }

    return products;
}

// ── OPEX Catalog ────────────────────────────────────────────────────────────
// Flattens opex.categories into a flat array of expense items

function buildOpexCatalog() {
    const items = [];
    const { categories } = modelInputs.opex;

    for (const [categoryName, subItems] of Object.entries(categories)) {
        for (const item of subItems) {
            // Parse cost: can be a number or a formula string like "formula: =295000*I8"
            let baseCost = 0;
            let costFormula = '';
            if (typeof item.cost_per_month_per_branch === 'number') {
                baseCost = item.cost_per_month_per_branch;
            } else if (typeof item.cost_per_month_per_branch === 'string') {
                costFormula = item.cost_per_month_per_branch;
                // Try to extract the numeric part before *I8
                const match = costFormula.match(/=(\d+)/);
                if (match) baseCost = parseInt(match[1], 10);
                // Handle division formulas like =130000/12 or =145000/12/5
                const divMatch = costFormula.match(/=(\d+)\/(\d+)(?:\/(\d+))?/);
                if (divMatch) {
                    baseCost = parseInt(divMatch[1], 10) / parseInt(divMatch[2], 10);
                    if (divMatch[3]) baseCost /= parseInt(divMatch[3], 10);
                }
            }

            items.push({
                category: categoryName,
                subService: item.sub_service,
                id: item.id,
                baseCost: Math.round(baseCost),
                costFormula,
                growthRates: item.growth_rates_pct || {},
                cellRefs: item.cell_refs || {},
            });
        }
    }

    return items;
}

// ── Branch Data ─────────────────────────────────────────────────────────────

function buildBranchData() {
    const br = modelInputs.branch_rollout;
    return {
        totalMonths: br.total_months,
        milestones: br.branch_milestones,
        keyInputs: br.key_inputs,
        branchesCell: modelInputs.revenue_streams.branches_cell, // H7
    };
}

// ── Condensed Prompt Text ───────────────────────────────────────────────────
// Produces a compact text representation for the AI system prompt

function buildPromptCatalog() {
    const rev = buildRevenueCatalog();
    const opex = buildOpexCatalog();

    // Group revenue by stream > sub-stream
    const revByStream = {};
    for (const p of rev) {
        if (!revByStream[p.stream]) revByStream[p.stream] = {};
        if (!revByStream[p.stream][p.subStream]) revByStream[p.stream][p.subStream] = [];
        revByStream[p.stream][p.subStream].push(p);
    }

    let revenueText = '';
    for (const [stream, subs] of Object.entries(revByStream)) {
        revenueText += `\n### ${stream}\n`;
        for (const [sub, products] of Object.entries(subs)) {
            revenueText += `**${sub}:**\n`;
            for (const p of products) {
                const priceStr = p.price != null ? `₹${p.price.toLocaleString('en-IN')}` : 'NO DEFAULT (ask user)';
                const qtyStr = p.qtyNote || 'no formula';
                const branchStr = p.branchScaled ? '×B' : 'flat';
                revenueText += `- ${p.product} | qty: ${qtyStr} | price: ${priceStr} | ${branchStr} | cell: qty=${p.cellRefs.qty || '?'}, price=${p.cellRefs.price || '?'}\n`;
            }
        }
    }

    // OPEX catalog
    let opexText = '';
    const opexByCategory = {};
    for (const item of opex) {
        if (!opexByCategory[item.category]) opexByCategory[item.category] = [];
        opexByCategory[item.category].push(item);
    }
    for (const [cat, items] of Object.entries(opexByCategory)) {
        opexText += `\n### ${cat}\n`;
        for (const item of items) {
            const costStr = item.baseCost > 0 ? `₹${item.baseCost.toLocaleString('en-IN')}/mo/branch` : 'formula-based';
            opexText += `- ${item.subService} | ${costStr} | cell: cost=${item.cellRefs.cost || '?'}\n`;
        }
    }

    return { revenueText, opexText, totalProducts: rev.length, totalOpexItems: opex.length };
}

// ── Exports ─────────────────────────────────────────────────────────────────

export const revenueCatalog = buildRevenueCatalog();
export const opexCatalog = buildOpexCatalog();
export const branchData = buildBranchData();
export const promptCatalog = buildPromptCatalog();

// Helper: find a revenue product by name (fuzzy match)
export function findRevenueProduct(productName) {
    const name = (productName || '').toLowerCase().trim();
    // Exact match
    let found = revenueCatalog.find(p => p.product.toLowerCase() === name);
    if (found) return found;
    // Partial match
    found = revenueCatalog.find(p =>
        name.includes(p.product.toLowerCase()) || p.product.toLowerCase().includes(name)
    );
    return found || null;
}

// Helper: find an OPEX item by sub-service name (fuzzy match)
export function findOpexItem(subService) {
    const name = (subService || '').toLowerCase().trim();
    let found = opexCatalog.find(p => p.subService.toLowerCase() === name);
    if (found) return found;
    found = opexCatalog.find(p =>
        name.includes(p.subService.toLowerCase()) || p.subService.toLowerCase().includes(name)
    );
    return found || null;
}

export default modelInputs;
