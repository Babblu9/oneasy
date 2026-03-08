'use client';
import React from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

export default function BasicsSheet() {
    const { businessInfo, setBusinessInfo, assumptions, setAssumptions, funding } = useFinancial();

    const rows = [
        // ── Business Identity ──
        { id: 'section-identity', metric: '━━ BUSINESS IDENTITY', value: '', type: 'header', editable: false },
        { id: 'legalName', metric: 'Legal Business Name', value: businessInfo.legalName || '—', type: 'text', editable: true, section: 'info' },
        { id: 'tradeName', metric: 'Trading / Brand Name', value: businessInfo.tradeName || '—', type: 'text', editable: true, section: 'info' },
        { id: 'companyType', metric: 'Company Type', value: businessInfo.companyType || '—', type: 'text', editable: true, section: 'info' },
        { id: 'address', metric: 'Registered Address', value: businessInfo.address || '—', type: 'text', editable: true, section: 'info' },
        { id: 'email', metric: 'Official Email', value: businessInfo.email || '—', type: 'text', editable: true, section: 'info' },
        { id: 'phone', metric: 'Official Phone', value: businessInfo.phone || '—', type: 'text', editable: true, section: 'info' },
        { id: 'promoters', metric: 'Founders / Promoters', value: Array.isArray(businessInfo.promoters) && businessInfo.promoters.length > 0 ? businessInfo.promoters.join(', ') : '—', type: 'text', editable: true, section: 'info' },
        { id: 'startDate', metric: 'Phase 1 Start Date', value: businessInfo.startDate || '—', type: 'text', editable: true, section: 'info' },
        { id: 'description', metric: 'Business Description', value: businessInfo.description || '—', type: 'text', editable: true, section: 'info' },

        // ── Share Capital ──
        { id: 'section-capital', metric: '━━ SHARE CAPITAL', value: '', type: 'header', editable: false },
        { id: 'equityShares', metric: 'Number of Equity Shares', value: businessInfo.equityShares || 0, type: 'number', editable: true, section: 'info' },
        { id: 'faceValue', metric: 'Face Value per Share (₹)', value: businessInfo.faceValue || 0, type: 'currency', editable: true, section: 'info' },
        { id: 'paidUpCapital', metric: 'Paid-Up Capital (₹)', value: businessInfo.paidUpCapital || (businessInfo.equityShares * businessInfo.faceValue) || 0, type: 'currency', editable: false },

        // ── Financial Assumptions ──
        { id: 'section-assumptions', metric: '━━ FINANCIAL ASSUMPTIONS', value: '', type: 'header', editable: false },
        { id: 'initialInvestment', metric: 'Initial Seed Investment (₹)', value: assumptions.initialInvestment, type: 'currency', editable: true, section: 'assumptions' },
        { id: 'taxRate', metric: 'Corporate Tax Rate', value: assumptions.taxRate, type: 'percent', editable: true, section: 'assumptions' },
        { id: 'inflationRate', metric: 'Inflation Rate (YoY)', value: assumptions.inflationRate, type: 'percent', editable: true, section: 'assumptions' },

        // ── Funding Structure ──
        { id: 'section-funding', metric: '━━ FUNDING STRUCTURE', value: '', type: 'header', editable: false },
        { id: 'equityFromPromoters', metric: 'Promoter Equity (₹)', value: funding.equityFromPromoters || 0, type: 'currency', editable: false },
        { id: 'loanAmount', metric: 'Loan Amount (₹)', value: funding.loanAmount || 0, type: 'currency', editable: false },
        { id: 'interestRate', metric: 'Interest Rate (%)', value: funding.interestRate || 0, type: 'percent', editable: false },
        { id: 'loanTenureMonths', metric: 'Loan Tenure (Months)', value: funding.loanTenureMonths || 0, type: 'number', editable: false },
        { id: 'moratoriumMonths', metric: 'Moratorium Period (Months)', value: funding.moratoriumMonths || 0, type: 'number', editable: false },
    ];

    const handleEdit = (rowIndex, colKey, val) => {
        if (colKey !== 'value') return;
        const row = rows[rowIndex];
        if (!row || !row.editable) return;

        if (row.section === 'assumptions') {
            const numVal = Number(val);
            if (!isNaN(numVal)) {
                setAssumptions(prev => ({ ...prev, [row.id]: numVal }));
            }
        } else if (row.section === 'info') {
            if (row.type === 'number' || row.type === 'currency') {
                setBusinessInfo({ [row.id]: Number(val) || 0 });
            } else {
                setBusinessInfo({ [row.id]: String(val).substring(0, 500) });
            }
        }
    };

    return (
        <DataGrid
            headers={[
                { key: 'metric', label: 'Field', editable: false, sticky: true },
                { key: 'value', label: 'Value', editable: true }
            ]}
            rows={rows}
            onCellEdit={handleEdit}
        />
    );
}
