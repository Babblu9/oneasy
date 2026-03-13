/**
 * engine/industryTemplates.js
 * ============================
 * Deterministic industry templates.
 * The AI selects an industry → the engine uses these defaults.
 * AI does NOT invent streams — it picks from here.
 *
 * Each template defines:
 *   - revenueStreams[]  : typical revenue lines with realistic Indian defaults
 *   - opex[]           : typical monthly operating expenses
 *   - growthProfile    : sane growth rate defaults per year
 *   - kpis             : human-readable KPI labels shown on the dashboard
 */

import { classifyByKeywords } from '@/lib/templateRegistry';

// ─── Growth Profiles ─────────────────────────────────────────────────────────

const GROWTH_PROFILES = {
    /** Conservative — services businesses, consulting, offline retail */
    conservative: { y1: 0.05, y2: 0.10, y3: 0.15, y4: 0.12, y5: 0.10 },
    /** Moderate — healthcare, education, ecommerce */
    moderate: { y1: 0.10, y2: 0.20, y3: 0.25, y4: 0.20, y5: 0.15 },
    /** Aggressive — SaaS, fintech, consumer apps */
    aggressive: { y1: 0.15, y2: 0.35, y3: 0.50, y4: 0.30, y5: 0.20 },
};

// ─── Industry Templates ───────────────────────────────────────────────────────

export const INDUSTRY_TEMPLATES = {

    healthcare: {
        id: 'healthcare',
        name: 'Healthcare Clinic',
        icon: '🏥',
        growthProfile: 'moderate',
        keywords: ['clinic', 'hospital', 'dental', 'doctor', 'pharmacy', 'diagnostic', 'physio', 'medical'],
        revenueStreams: [
            { stream: 'Doctor Consultations', name: 'OPD Consultation Fee', price: 500, quantity: 400 },
            { stream: 'Diagnostics', name: 'Lab Tests & Imaging', price: 1200, quantity: 200 },
            { stream: 'Pharmacy', name: 'Pharmacy Margin', price: 300, quantity: 350 },
            { stream: 'Telemedicine', name: 'Online Consultation', price: 300, quantity: 100 },
            { stream: 'Health Packages', name: 'Annual Health Plan', price: 5000, quantity: 20 },
        ],
        opex: [
            { category: 'Staff', name: 'Doctor Salaries', monthlyCost: 400000 },
            { category: 'Staff', name: 'Nursing & Para-Medical Staff', monthlyCost: 180000 },
            { category: 'Staff', name: 'Admin & Front Desk', monthlyCost: 80000 },
            { category: 'Facilities', name: 'Clinic Rent', monthlyCost: 120000 },
            { category: 'Supplies', name: 'Medical Consumables', monthlyCost: 95000 },
            { category: 'Operations', name: 'Software & Administration', monthlyCost: 35000 },
            { category: 'Marketing', name: 'Patient Acquisition', monthlyCost: 50000 },
        ],
        kpis: ['Patients/Month', 'Revenue/Patient', 'Occupancy Rate', 'EBITDA Margin'],
    },

    edtech: {
        id: 'edtech',
        name: 'EdTech / Education',
        icon: '🎓',
        growthProfile: 'aggressive',
        keywords: ['edtech', 'education', 'courses', 'learning', 'training', 'lms', 'tutoring', 'certification', 'e-learning'],
        revenueStreams: [
            { stream: 'Course Sales', name: 'Recorded Courses', price: 3000, quantity: 200 },
            { stream: 'Subscriptions', name: 'Monthly Membership', price: 499, quantity: 400 },
            { stream: 'Corporate Training', name: 'B2B Training Contracts', price: 80000, quantity: 3 },
            { stream: 'Certifications', name: 'Exam & Certification Fees', price: 1500, quantity: 100 },
            { stream: 'Live Workshops', name: 'Cohort/Bootcamp Fees', price: 8000, quantity: 30 },
        ],
        opex: [
            { category: 'Content', name: 'Faculty & Instructor Payouts', monthlyCost: 250000 },
            { category: 'Content', name: 'Content Production', monthlyCost: 80000 },
            { category: 'Technology', name: 'Platform Hosting & CDN', monthlyCost: 60000 },
            { category: 'Marketing', name: 'Performance Marketing', monthlyCost: 150000 },
            { category: 'Staff', name: 'Student Support Team', monthlyCost: 85000 },
            { category: 'Operations', name: 'Fees & Compliance', monthlyCost: 25000 },
        ],
        kpis: ['Students/Month', 'Revenue/Student', 'Course Completion Rate', 'CAC'],
    },

    saas: {
        id: 'saas',
        name: 'SaaS / Software',
        icon: '💻',
        growthProfile: 'aggressive',
        keywords: ['saas', 'software', 'subscription', 'platform', 'app', 'api', 'b2b software', 'tech product'],
        revenueStreams: [
            { stream: 'Subscriptions', name: 'Starter Plan MRR', price: 1499, quantity: 300 },
            { stream: 'Subscriptions', name: 'Pro Plan MRR', price: 4999, quantity: 80 },
            { stream: 'Enterprise', name: 'Enterprise Annual Contract', price: 120000, quantity: 5 },
            { stream: 'Usage', name: 'API Overage / Usage Fees', price: 0.05, quantity: 200000 },
            { stream: 'Services', name: 'Onboarding & Implementation', price: 25000, quantity: 8 },
        ],
        opex: [
            { category: 'Engineering', name: 'Engineering Team Salaries', monthlyCost: 500000 },
            { category: 'Infrastructure', name: 'Cloud Hosting & DevOps', monthlyCost: 120000 },
            { category: 'Sales', name: 'Sales & BD Team', monthlyCost: 200000 },
            { category: 'Customer Success', name: 'Customer Success Team', monthlyCost: 120000 },
            { category: 'Marketing', name: 'Product Marketing & Ads', monthlyCost: 150000 },
            { category: 'Operations', name: 'SaaS Tools & Subscriptions', monthlyCost: 30000 },
        ],
        kpis: ['MRR', 'ARR', 'Churn Rate', 'LTV/CAC', 'Net Revenue Retention'],
    },

    ecommerce: {
        id: 'ecommerce',
        name: 'E-commerce / D2C',
        icon: '🛒',
        growthProfile: 'moderate',
        keywords: ['ecommerce', 'shop', 'store', 'marketplace', 'd2c', 'retail', 'dropshipping', 'product sales'],
        revenueStreams: [
            { stream: 'Direct Sales', name: 'Website / App Orders', price: 1500, quantity: 600 },
            { stream: 'Marketplace', name: 'Amazon / Flipkart Sales', price: 1200, quantity: 400 },
            { stream: 'Wholesale', name: 'Distributor Orders', price: 800, quantity: 200 },
            { stream: 'Subscriptions', name: 'Subscription Box', price: 999, quantity: 80 },
            { stream: 'Upsells', name: 'Bundles & Add-ons', price: 500, quantity: 300 },
        ],
        opex: [
            { category: 'Inventory', name: 'Inventory / COGS', monthlyCost: 600000 },
            { category: 'Fulfillment', name: 'Warehousing & Logistics', monthlyCost: 180000 },
            { category: 'Marketing', name: 'Paid Ads & Influencers', monthlyCost: 200000 },
            { category: 'Staff', name: 'Operations Team', monthlyCost: 120000 },
            { category: 'Technology', name: 'Platform & Tools', monthlyCost: 40000 },
            { category: 'Returns', name: 'Returns & Refunds Buffer', monthlyCost: 60000 },
        ],
        kpis: ['AOV', 'Orders/Month', 'Return Rate', 'ROAS', 'Gross Margin'],
    },

    pharma: {
        id: 'pharma',
        name: 'Pharma / Pharmacy',
        icon: '💊',
        growthProfile: 'conservative',
        keywords: ['pharma', 'pharmacy', 'medicine', 'drug', 'chemist', 'medstore', 'otc'],
        revenueStreams: [
            { stream: 'Prescription', name: 'Rx Drug Sales', price: 350, quantity: 800 },
            { stream: 'OTC', name: 'OTC Medicines & Wellness', price: 250, quantity: 600 },
            { stream: 'Hospital Supply', name: 'Hospital Bulk Supply', price: 500000, quantity: 1 },
            { stream: 'Generic', name: 'Generic Medicines', price: 120, quantity: 500 },
        ],
        opex: [
            { category: 'Staff', name: 'Pharmacist & Staff Salaries', monthlyCost: 200000 },
            { category: 'Facilities', name: 'Store Rent & Utilities', monthlyCost: 120000 },
            { category: 'Inventory', name: 'Inventory Carry Cost', monthlyCost: 250000 },
            { category: 'Compliance', name: 'Regulatory & Licensing', monthlyCost: 30000 },
            { category: 'Operations', name: 'Cold Chain & Storage', monthlyCost: 50000 },
        ],
        kpis: ['Daily Sales', 'Inventory Turnover', 'Gross Margin %', 'Repeat Customer Rate'],
    },

    consulting: {
        id: 'consulting',
        name: 'Consulting / Services',
        icon: '📋',
        growthProfile: 'conservative',
        keywords: ['consulting', 'agency', 'services', 'advisory', 'freelance', 'it services', 'professional services'],
        revenueStreams: [
            { stream: 'Retainer', name: 'Monthly Retainer Clients', price: 150000, quantity: 6 },
            { stream: 'Project', name: 'Fixed-Price Projects', price: 250000, quantity: 3 },
            { stream: 'Advisory', name: 'Advisory Calls / CXO Coaching', price: 25000, quantity: 10 },
            { stream: 'Workshops', name: 'Training Workshops', price: 75000, quantity: 2 },
        ],
        opex: [
            { category: 'Staff', name: 'Consultant Salaries', monthlyCost: 350000 },
            { category: 'Business Development', name: 'Sales & BD', monthlyCost: 80000 },
            { category: 'Operations', name: 'Office & Travel', monthlyCost: 70000 },
            { category: 'Technology', name: 'Software & Tools', monthlyCost: 30000 },
            { category: 'Marketing', name: 'Brand & Thought Leadership', monthlyCost: 40000 },
        ],
        kpis: ['Active Clients', 'Revenue/Consultant', 'Utilization Rate', 'Project Margin'],
    },

    ca_firm: {
        id: 'ca_firm',
        name: 'CA Firm / Accounting Practice',
        icon: '🧾',
        growthProfile: 'conservative',
        keywords: ['ca firm', 'chartered accountant', 'chartered accountants', 'accounting firm', 'tax consultant', 'gst filing', 'audit firm', 'bookkeeping', 'roc compliance'],
        revenueStreams: [
            { stream: 'Audit & Assurance', name: 'Statutory Audit Engagements', price: 60000, quantity: 8 },
            { stream: 'Taxation', name: 'Income Tax Filing & Advisory', price: 12000, quantity: 35 },
            { stream: 'GST & Compliance', name: 'GST Returns and Compliance', price: 8000, quantity: 40 },
            { stream: 'Accounting Services', name: 'Bookkeeping & MIS', price: 15000, quantity: 18 },
            { stream: 'Advisory', name: 'Virtual CFO / Business Advisory', price: 50000, quantity: 6 },
        ],
        opex: [
            { category: 'Staff', name: 'Partner / CA Salaries', monthlyCost: 300000 },
            { category: 'Staff', name: 'Article Assistants & Accountants', monthlyCost: 180000 },
            { category: 'Operations', name: 'Office Rent & Utilities', monthlyCost: 90000 },
            { category: 'Technology', name: 'Accounting, Tax, and Compliance Software', monthlyCost: 35000 },
            { category: 'Business Development', name: 'Client Acquisition & Networking', monthlyCost: 40000 },
            { category: 'Admin', name: 'Admin & Filing Costs', monthlyCost: 25000 },
        ],
        kpis: ['Active Clients', 'Revenue/Client', 'Monthly Filings', 'Advisory Margin'],
    },

    manufacturing: {
        id: 'manufacturing',
        name: 'Manufacturing',
        icon: '🏭',
        growthProfile: 'conservative',
        keywords: ['manufacturing', 'factory', 'production', 'fmcg', 'industrial', 'assembly', 'supply chain'],
        revenueStreams: [
            { stream: 'Direct Sales', name: 'Retail Product Sales', price: 550, quantity: 5000 },
            { stream: 'Wholesale', name: 'Distributor / Bulk Orders', price: 400, quantity: 3000 },
            { stream: 'OEM', name: 'OEM Contracts', price: 300000, quantity: 2 },
            { stream: 'Export', name: 'Export Orders', price: 450, quantity: 2000 },
        ],
        opex: [
            { category: 'Labour', name: 'Factory Labour & Workers', monthlyCost: 500000 },
            { category: 'Materials', name: 'Raw Materials & Components', monthlyCost: 800000 },
            { category: 'Utilities', name: 'Power & Utilities', monthlyCost: 220000 },
            { category: 'Facilities', name: 'Plant Rent & Maintenance', monthlyCost: 180000 },
            { category: 'Logistics', name: 'Distribution & Freight', monthlyCost: 150000 },
            { category: 'Compliance', name: 'Quality & Regulatory', monthlyCost: 60000 },
        ],
        kpis: ['Units Produced', 'Revenue/Unit', 'Capacity Utilization', 'Gross Margin'],
    },
};

// ─── Selector / Classifier ────────────────────────────────────────────────────

const INDUSTRY_KEYS = Object.keys(INDUSTRY_TEMPLATES);

/**
 * Get a template by id. Falls back to consulting if unknown.
 * @param {string} id
 * @returns {object} template
 */
export function getIndustryTemplate(id) {
    const key = (id || 'consulting').toLowerCase().trim();
    return INDUSTRY_TEMPLATES[key] || INDUSTRY_TEMPLATES.consulting;
}

/**
 * Classify a free-text description into an industry template.
 * Uses keyword scoring — no AI call needed for obvious cases.
 * @param {string} text
 * @returns {{ industryId: string, template: object, score: number }}
 */
export function classifyIndustry(text) {
    const lower = (text || '').toLowerCase();
    let best = { industryId: 'consulting', score: 0 };

    for (const [id, tmpl] of Object.entries(INDUSTRY_TEMPLATES)) {
        const score = tmpl.keywords.reduce((acc, kw) => acc + (lower.includes(kw) ? 1 : 0), 0);
        if (score > best.score) best = { industryId: id, score };
    }

    // Also try the keyword-based classifier from templateRegistry
    try {
        const reg = classifyByKeywords(text);
        if (reg?.score > best.score) best = { industryId: reg.templateId, score: reg.score };
    } catch { /* ignore */ }

    return { industryId: best.industryId, template: getIndustryTemplate(best.industryId), score: best.score };
}

/**
 * Get the growth rates for an industry profile.
 * @param {string} profile - 'conservative' | 'moderate' | 'aggressive'
 * @returns {object}
 */
export function getGrowthRates(profile) {
    return GROWTH_PROFILES[profile] || GROWTH_PROFILES.moderate;
}

export default INDUSTRY_TEMPLATES;
