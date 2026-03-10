/**
 * engine/index.js
 * ================
 * Barrel export for the financial engine module.
 *
 * Usage:
 *   import { generateFinancialModel } from '@/lib/engine';
 *   import { INDUSTRY_TEMPLATES, getIndustryTemplate } from '@/lib/engine';
 *   import { AssumptionsInputSchema, parseMoneyLike } from '@/lib/engine';
 */

export { generateFinancialModel, safeGenerateFinancialModel } from './financialEngine.js';
export { INDUSTRY_TEMPLATES, getIndustryTemplate, classifyIndustry, getGrowthRates } from './industryTemplates.js';
export {
    parseMoneyLike,
    AssumptionsInputSchema,
    RevenueStreamInputSchema,
    OpexItemInputSchema,
    LoanInputSchema,
    FinancialModelSchema,
} from './schemas.js';
