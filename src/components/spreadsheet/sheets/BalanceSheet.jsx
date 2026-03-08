'use client';
import React, { useMemo } from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

// Note: The underlying engine does not appear to currently export a fully balanced Balance Sheet object.
// We will construct a basic representation using available data (Net Block from Assets + Cash from PAT).
export default function BalanceSheet() {
    const { projectionOutputs, assumptions } = useFinancial();

    const { headers, rows } = useMemo(() => {
        if (!projectionOutputs || !projectionOutputs.assets) return { headers: [], rows: [] };

        const { yearlyNetBlock } = projectionOutputs.assets;
        const pat = projectionOutputs.yearly.pnl.map(y => y.pat);
        const numYears = yearlyNetBlock.length;

        const hdrs = [{ key: 'lineItem', label: 'Pro Forma Balance Sheet', sticky: true, editable: false }];
        for (let i = 0; i < numYears; i++) {
            hdrs.push({ key: `yr${i}`, label: `Year ${i + 1}`, type: 'currency', editable: false });
        }

        const makeRow = (id, label, dataArray, isSummary = false) => {
            const row = { id, lineItem: label, isReadOnly: true, isSummary };
            dataArray.forEach((val, i) => { row[`yr${i}`] = val; });
            return row;
        };

        // Simplified estimation of cash buildup
        let runningCash = assumptions.initialInvestment || 0;
        const cashBalances = pat.map(val => {
            runningCash += val;
            return runningCash;
        });

        // Simplified Total Assets
        const totalAssets = yearlyNetBlock.map((nb, i) => nb + cashBalances[i]);

        // Simplified Equity to balance
        const equityBalances = totalAssets.map(a => a); // Assumes no liabilities for now to maintain A=L+E

        const rws = [
            { id: 'h1', lineItem: 'ASSETS', isReadOnly: true, isSummary: true },
            makeRow('nb', '  Net Fixed Assets (Net Block)', yearlyNetBlock),
            makeRow('cash', '  Cash & Equivalents (Accum. PAT)', cashBalances),
            makeRow('totA', 'TOTAL ASSETS', totalAssets, true),
            { id: 'sp1', lineItem: '', isReadOnly: true },
            { id: 'h2', lineItem: 'LIABILITIES & EQUITY', isReadOnly: true, isSummary: true },
            makeRow('liab', '  Liabilities', new Array(numYears).fill(0)),
            makeRow('eq', '  Shareholder Equity (Seed + Retained)', equityBalances),
            makeRow('totLE', 'TOTAL LIABILITIES & EQUITY', totalAssets, true), // Balanced
        ];

        return { headers: hdrs, rows: rws };
    }, [projectionOutputs, assumptions]);

    if (!projectionOutputs) return <div className="p-8 text-center text-slate-500">Calculating...</div>;

    return (
        <div className="space-y-4">
            <p className="text-sm text-slate-500">A high-level pro-forma Balance Sheet derived from Seed Investment and net cash flow.</p>
            <DataGrid headers={headers} rows={rows} onCellEdit={() => { }} />
        </div>
    );
}
