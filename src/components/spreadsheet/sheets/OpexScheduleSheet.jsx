'use client';
import React, { useMemo } from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

export default function OpexScheduleSheet() {
    const { projectionOutputs } = useFinancial();

    const { headers, rows } = useMemo(() => {
        if (!projectionOutputs || !projectionOutputs.yearly.byCategory) return { headers: [], rows: [] };

        const catData = projectionOutputs.yearly.byCategory;
        const opexTotal = projectionOutputs.yearly.pnl.map(y => y.opex);
        // By default pnl array is the length of years
        const numYears = opexTotal.length;

        const hdrs = [{ key: 'category', label: 'OPEX Category', sticky: true, editable: false }];
        for (let i = 0; i < numYears; i++) {
            hdrs.push({ key: `yr${i}`, label: `Year ${i + 1}`, type: 'currency', editable: false });
        }

        const rws = [];

        Object.keys(catData).forEach(cat => {
            const arr = catData[cat];
            const row = { id: cat, category: cat, isReadOnly: true };
            for (let i = 0; i < numYears; i++) {
                row[`yr${i}`] = arr[i] || 0;
            }
            rws.push(row);
        });

        // Add Total Row
        const totalRow = { id: 'total', category: 'Total Operating Expenses', isReadOnly: true, isSummary: true };
        for (let i = 0; i < numYears; i++) {
            totalRow[`yr${i}`] = opexTotal[i] || 0;
        }

        rws.push({ id: 'sp1', category: '', isReadOnly: true });
        rws.push(totalRow);

        return { headers: hdrs, rows: rws };
    }, [projectionOutputs]);

    if (!projectionOutputs) return <div className="p-8 text-center text-slate-500">Calculating...</div>;

    return (
        <div className="space-y-4">
            <p className="text-sm text-slate-500">Yearly Operating Expense (OPEX) projections including inflation (1%) and ramp-up factors.</p>
            <DataGrid headers={headers} rows={rows} onCellEdit={() => { }} />
        </div>
    );
}
