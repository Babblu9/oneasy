'use client';
import React, { useMemo } from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

export default function PnlSheet() {
    const { projectionOutputs } = useFinancial();

    const { headers, rows } = useMemo(() => {
        if (!projectionOutputs || !projectionOutputs.yearly) {
            return { headers: [], rows: [] };
        }

        const yearlyPnl = projectionOutputs.yearly.pnl || [];
        const numYears = yearlyPnl.length;

        // Build headers: Line Item + Year 1 to Year N
        const hdrs = [{ key: 'lineItem', label: 'Consolidated P&L', sticky: true, editable: false }];
        for (let i = 0; i < numYears; i++) {
            hdrs.push({ key: `yr${i}`, label: yearlyPnl[i].fy || `Year ${i + 1}`, type: 'currency', editable: false });
        }

        // Helper to format a row from an array of values
        const makeRow = (id, label, dataArray, isSummary = false) => {
            const row = { id, lineItem: label, isReadOnly: true, isSummary };
            dataArray.forEach((val, i) => {
                row[`yr${i}`] = val;
            });
            return row;
        };

        // Extract arrays for each property
        const getArray = (propName) => yearlyPnl.map(y => y[propName] || 0);

        const rws = [
            makeRow('rev', 'Gross Revenue Receipts', getArray('revenue'), true),
            makeRow('opex', 'Operating Expenses (OPEX)', getArray('opex')),
            makeRow('ebitda', 'EBITDA', getArray('ebitda'), true),
            makeRow('ebitdaMargin', 'EBITDA Margin', getArray('ebitdaMargin'), true),
            makeRow('depr', 'Depreciation & Amortization', getArray('depreciation')),
            makeRow('pbt', 'Profit Before Tax (PBT)', getArray('pbt'), true),
            makeRow('tax', 'Taxes', getArray('tax')),
            makeRow('netIncome', 'Net Income (PAT)', getArray('pat'), true),
            makeRow('patMargin', 'Net Profit Margin', getArray('patMargin'), true)
        ];

        // Format margins explicitly as percentages (already multiplied by 100 in engine)
        rws.forEach(r => {
            if (r.id === 'ebitdaMargin' || r.id === 'patMargin') {
                for (let i = 0; i < numYears; i++) {
                    const val = r[`yr${i}`];
                    r[`yr${i}`] = val !== undefined ? `${Number(val).toFixed(1)}%` : '0.0%';
                }
            }
        });

        // Add blank spacer row
        rws.splice(1, 0, { id: 'sp1', lineItem: '', isReadOnly: true });
        rws.splice(4, 0, { id: 'sp2', lineItem: '', isReadOnly: true });
        rws.splice(7, 0, { id: 'sp3', lineItem: '', isReadOnly: true });

        return { headers: hdrs, rows: rws };
    }, [projectionOutputs]);

    if (!projectionOutputs) {
        return <div className="p-8 text-center text-slate-500">Calculating Engine Outputs...</div>;
    }

    return (
        <DataGrid
            headers={headers}
            rows={rows}
            onCellEdit={() => { }}
        />
    );
}
