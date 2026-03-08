'use client';
import React, { useMemo } from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

export default function RatiosSheet() {
    const { projectionOutputs } = useFinancial();

    const { headers, rows } = useMemo(() => {
        if (!projectionOutputs || !projectionOutputs.yearly.pnl) return { headers: [], rows: [] };

        const { pnl } = projectionOutputs.yearly;
        const numYears = pnl.length;

        const hdrs = [{ key: 'lineItem', label: 'Financial Ratios & Margins', sticky: true, editable: false }];
        for (let i = 0; i < numYears; i++) {
            hdrs.push({ key: `yr${i}`, label: `Year ${i + 1}`, type: 'percent', editable: false });
        }

        const makeRow = (id, label, propName, isCagr = false) => {
            const row = { id, lineItem: label, isReadOnly: true };
            for (let i = 0; i < numYears; i++) {
                // If it's a growth metric, it's already in whole numbers from the engine (e.g. 15 for 15%).
                // Our DataGrid format expects decimals for 'percent' type (0.15 for 15%), so we divide by 100.
                row[`yr${i}`] = pnl[i][propName] / 100;
            }
            return row;
        };

        const rws = [
            { id: 'h1', lineItem: 'Profitability Ratios', isReadOnly: true, isSummary: true },
            makeRow('ebitdaM', 'EBITDA Margin', 'ebitdaMargin'),
            makeRow('patM', 'Net Profit Margin', 'patMargin'),
            { id: 'sp1', lineItem: '', isReadOnly: true },
            { id: 'h2', lineItem: 'Growth Metrics (YoY)', isReadOnly: true, isSummary: true },
            makeRow('revG', 'Revenue Growth', 'revenueGrowth'),
            makeRow('opexG', 'OPEX Growth', 'opexGrowth'),
        ];

        return { headers: hdrs, rows: rws };
    }, [projectionOutputs]);

    if (!projectionOutputs) return <div className="p-8 text-center text-slate-500">Calculating...</div>;

    return (
        <div className="space-y-4">
            <p className="text-sm text-slate-500">Key performance indicators, profitability margins, and YoY growth trajectories.</p>
            <DataGrid headers={headers} rows={rows} onCellEdit={() => { }} />
        </div>
    );
}
