'use client';
import React, { useMemo } from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

export default function FaScheduleSheet() {
    const { projectionOutputs } = useFinancial();

    const { headers, rows } = useMemo(() => {
        if (!projectionOutputs || !projectionOutputs.assets) return { headers: [], rows: [] };

        const { details, yearlyTotal, yearlyAccumulated, yearlyNetBlock } = projectionOutputs.assets;
        const numYears = yearlyTotal.length;

        const hdrs = [{ key: 'lineItem', label: 'Fixed Asset Schedule', sticky: true, editable: false }];
        for (let i = 0; i < numYears; i++) {
            hdrs.push({ key: `yr${i}`, label: `Year ${i + 1}`, type: 'currency', editable: false });
        }

        const makeRow = (id, label, dataArray, isSummary = false) => {
            const row = { id, lineItem: label, isReadOnly: true, isSummary };
            dataArray.forEach((val, i) => {
                row[`yr${i}`] = val;
            });
            return row;
        };

        const rws = [];

        // Detail Rows
        details.forEach(asset => {
            const row = { id: asset.name, lineItem: `${asset.name} (Depreciation)`, isReadOnly: true };
            for (let i = 0; i < numYears; i++) {
                row[`yr${i}`] = asset.schedule[i]?.depreciation || 0;
            }
            rws.push(row);
        });

        rws.push({ id: 'sp1', lineItem: '', isReadOnly: true });

        // Summary Rows
        rws.push(makeRow('totalDepr', 'Total Annual Depreciation', yearlyTotal, true));
        rws.push(makeRow('accumDepr', 'Accumulated Depreciation', yearlyAccumulated));
        rws.push(makeRow('netBlock', 'Closing Net Block (Book Value)', yearlyNetBlock, true));

        return { headers: hdrs, rows: rws };

    }, [projectionOutputs]);

    if (!projectionOutputs) return <div className="p-8 text-center text-slate-500">Calculating...</div>;

    return (
        <div className="space-y-4">
            <p className="text-sm text-slate-500">Straight-line depreciation of Capital Expenditures over their useful life.</p>
            <DataGrid headers={headers} rows={rows} onCellEdit={() => { }} />
        </div>
    );
}
