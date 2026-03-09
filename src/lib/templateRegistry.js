/**
 * Template Registry
 * =================
 * Central registry for all 7 financial model templates.
 * Provides:
 *  - Template metadata (name, description, file, icon)
 *  - Question sets per template
 *  - Assumption → Excel cell mappings
 *  - Keyword sets for classification
 */

import templateIndex from '../data/templates.json';

/** Get all templates */
export function getAllTemplates() {
    return templateIndex;
}

/** Get a single template by id (e.g. "edtech", "saas") */
export function getTemplate(id) {
    return templateIndex.find(t => t.id === id) || null;
}

/**
 * Classify a free-form business description into a template id.
 * Uses keyword scoring as a fast, local fallback (no AI call needed for obvious cases).
 * Returns { templateId, score, template }
 */
export function classifyByKeywords(description) {
    const lower = (description || '').toLowerCase();
    let best = { templateId: 'consulting', score: 0 };

    for (const template of templateIndex) {
        const score = template.keywords.reduce((acc, kw) => {
            return acc + (lower.includes(kw) ? 1 : 0);
        }, 0);
        if (score > best.score) {
            best = { templateId: template.id, score, template };
        }
    }

    return { ...best, template: getTemplate(best.templateId) };
}

/**
 * Get the assumption cell map for a template.
 * Returns { key: { sheet, cell } }
 */
export function getAssumptionCells(templateId) {
    const t = getTemplate(templateId);
    return t ? t.assumptionCells : {};
}

/**
 * Get the question set for a template.
 */
export function getQuestions(templateId) {
    const t = getTemplate(templateId);
    return t ? t.questions : [];
}

/**
 * Get the Excel file path for a template.
 */
export function getTemplatePath(templateId) {
    const path = require('path');
    const t = getTemplate(templateId);
    if (!t) return null;
    return path.join(process.cwd(), 'templates', t.file);
}

export default templateIndex;
