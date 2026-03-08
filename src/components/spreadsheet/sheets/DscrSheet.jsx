'use client';
import React, { useMemo } from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

// Debt Service Coverage Ratio
export default function DscrSheet() {
    const { projectionOutputs } = useFinancial();

    const { headers, rows } = useMemo(() => {
        if (!projectionOutputs || !projectionOutputs.yearly.pnl) return { headers: [], rows: [] };

        const { pnl } = projectionOutputs.yearly;
        const numYears = pnl.length;

        const hdrs = [{ key: 'lineItem', label: 'Debt Service Coverage', sticky: true, editable: false }];
        for (let i = 0; i < numYears; i++) {
            hdrs.push({ key: `yr${i}`, label: `Year ${i + 1}`, type: 'number', editable: false });
        }

        const rws = [];

        // DSCR = Net Operating Income / Debt Service
        // For now, assume 0 debt until a Loan component is added to Context. DSCR = infinite or N/A.
        const row = { id: 'dscr', lineItem: 'DSCR (Times)', isReadOnly: true };
        for (let i = 0; i < numYears; i++) {
            row[`yr${i}`] = 'No Debt'; // Placeholder
        }

        rws.push(row);

        return { headers: hdrs, rows: rws };
    }, [projectionOutputs]);

    if (!projectionOutputs) return <div className="p-8 text-center text-slate-500">Calculating...</div>;

    return (
        <div className="space-y-4">
            <p className="text-sm text-slate-500">Debt Service Coverage Ratio. Note: Debt schedule inputs are not yet active in the current model.</p>
            <DataGrid headers={headers} rows={rows} onCellEdit={() => { }} />
        </div>
    );
}
