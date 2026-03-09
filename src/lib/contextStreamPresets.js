export const STREAM_PRESETS = {
    healthcare: {
        revenue: [
            { label: 'Doctor Consultations', value: 1800 },
            { label: 'Diagnostics / Lab Tests', value: 2200 },
            { label: 'Pharmacy Sales', value: 1400 },
            { label: 'Health Packages', value: 3500 },
            { label: 'Teleconsultation', value: 1200 }
        ],
        opex: [
            { label: 'Doctor Salaries', value: 350000 },
            { label: 'Nursing Staff', value: 180000 },
            { label: 'Clinic Rent', value: 120000 },
            { label: 'Medical Consumables', value: 95000 },
            { label: 'Front Desk/Admin', value: 80000 }
        ]
    },
    edtech: {
        revenue: [
            { label: 'Course Sales', value: 2200 },
            { label: 'Subscription Plans', value: 1200 },
            { label: 'Corporate Training', value: 60000 },
            { label: 'Certifications', value: 3500 },
            { label: 'Live Workshops', value: 2500 }
        ],
        opex: [
            { label: 'Faculty Payouts', value: 250000 },
            { label: 'Platform Hosting', value: 90000 },
            { label: 'Performance Marketing', value: 140000 },
            { label: 'Content Production', value: 110000 },
            { label: 'Student Support Team', value: 85000 }
        ]
    },
    saas: {
        revenue: [
            { label: 'Starter Plan MRR', value: 1500 },
            { label: 'Pro Plan MRR', value: 4000 },
            { label: 'Enterprise Contracts', value: 150000 },
            { label: 'Implementation Fees', value: 50000 },
            { label: 'Add-on Modules', value: 2500 }
        ],
        opex: [
            { label: 'Engineering Team', value: 500000 },
            { label: 'Cloud Infrastructure', value: 180000 },
            { label: 'Sales Team', value: 220000 },
            { label: 'Customer Success', value: 140000 },
            { label: 'Product Marketing', value: 120000 }
        ]
    },
    ecommerce: {
        revenue: [
            { label: 'Direct Product Sales', value: 1800 },
            { label: 'Marketplace Sales', value: 1500 },
            { label: 'Subscriptions', value: 800 },
            { label: 'Bundles / Upsells', value: 2200 },
            { label: 'Wholesale Orders', value: 50000 }
        ],
        opex: [
            { label: 'Fulfillment & Logistics', value: 210000 },
            { label: 'Inventory Purchase', value: 350000 },
            { label: 'Ad Spend', value: 240000 },
            { label: 'Warehouse Rent', value: 140000 },
            { label: 'Returns & Refunds', value: 90000 }
        ]
    },
    pharma: {
        revenue: [
            { label: 'Prescription Sales', value: 1900 },
            { label: 'OTC Sales', value: 1100 },
            { label: 'Hospital Supply', value: 85000 },
            { label: 'Health Products', value: 1400 },
            { label: 'Diagnostic Referrals', value: 1800 }
        ],
        opex: [
            { label: 'Pharmacist Salaries', value: 180000 },
            { label: 'Store Rent', value: 100000 },
            { label: 'Inventory Carry Cost', value: 250000 },
            { label: 'Cold Storage & Utilities', value: 70000 },
            { label: 'Regulatory Compliance', value: 45000 }
        ]
    },
    consulting: {
        revenue: [
            { label: 'Retainer Clients', value: 120000 },
            { label: 'Project Consulting', value: 180000 },
            { label: 'Workshops', value: 50000 },
            { label: 'Advisory Calls', value: 25000 },
            { label: 'Audit Services', value: 80000 }
        ],
        opex: [
            { label: 'Consultant Salaries', value: 320000 },
            { label: 'Travel & Client Meetings', value: 70000 },
            { label: 'Office Rent', value: 90000 },
            { label: 'Software Licenses', value: 45000 },
            { label: 'Business Development', value: 85000 }
        ]
    },
    manufacturing: {
        revenue: [
            { label: 'Core Product Sales', value: 550 },
            { label: 'Bulk Distributor Orders', value: 450000 },
            { label: 'OEM Contracts', value: 320000 },
            { label: 'After-Sales Service', value: 60000 },
            { label: 'Spare Parts', value: 300 }
        ],
        opex: [
            { label: 'Factory Labor', value: 550000 },
            { label: 'Raw Materials', value: 800000 },
            { label: 'Power & Utilities', value: 220000 },
            { label: 'Maintenance', value: 120000 },
            { label: 'Plant Rent', value: 180000 }
        ]
    }
};

const DEFAULT_TEMPLATE_ID = 'consulting';

export function getPresetStreams(templateId) {
    return STREAM_PRESETS[templateId] || STREAM_PRESETS[DEFAULT_TEMPLATE_ID];
}

export function normalizeStreamItems(items, type) {
    if (!Array.isArray(items)) return [];

    const pickNumericValue = (item) => {
        const candidates = [
            item?.value,
            item?.price,
            item?.amount,
            item?.monthly,
            item?.monthly_value,
            item?.monthlyValue,
            item?.monthly_revenue,
            item?.monthlyRevenue,
            item?.cost
        ];
        for (const c of candidates) {
            const n = Number(c);
            if (Number.isFinite(n) && n > 0) return n;
        }

        const y1 = Number(item?.y1);
        if (Number.isFinite(y1) && y1 > 0) return Math.round(y1 / 12);

        const qty = Number(item?.quantity ?? item?.qty ?? item?.units);
        const price = Number(item?.price_per_unit ?? item?.unit_price ?? item?.price);
        if (Number.isFinite(qty) && Number.isFinite(price) && qty > 0 && price > 0) {
            return qty * price;
        }

        return 0;
    };

    return items
        .map((item, index) => {
            const label = String(item?.label || item?.name || item?.substream || item?.stream || '').trim();
            if (!label) return null;
            return {
                id: item?.id || `${type}-${label.toLowerCase().replace(/[^a-z0-9]+/g, '-') || index}`,
                label,
                value: Math.max(0, pickNumericValue(item))
            };
        })
        .filter(Boolean);
}
