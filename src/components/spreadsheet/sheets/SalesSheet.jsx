'use client';
import React, { useMemo, useState } from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

export default function SalesSheet() {
    const { projectionOutputs } = useFinancial();
    const [viewMode, setViewMode] = useState('monthly'); // 'monthly' or 'yearly'

    const { headers, rows } = useMemo(() => {
        if (!projectionOutputs) return { headers: [], rows: [] };

        const isMonthly = viewMode === 'monthly';
        const data = isMonthly ? projectionOutputs.monthly.byStream : projectionOutputs.monthly.byStream; // for now just show monthly
        const periods = isMonthly ? 72 : 6;

        const hdrs = [{ key: 'stream', label: 'Revenue Stream', sticky: true, editable: false }];
        // Limit rendering to 36 months to prevent massive memory usage, or we can render all 72
        const renderCols = isMonthly ? 36 : 6;

        for (let i = 0; i < renderCols; i++) {
            hdrs.push({ key: `col${i}`, label: isMonthly ? `M${i + 1}` : `Y${i + 1}`, type: 'currency', editable: false });
        }

        const rws = [];
        if (data) {
            Object.keys(data).forEach(streamName => {
                const streamData = data[streamName]; // Array of 72 months
                const row = { id: streamName, stream: streamName, isReadOnly: true };
                for (let i = 0; i < renderCols; i++) {
                    row[`col${i}`] = streamData[i] || 0;
                }
                rws.push(row);
            });
        }

        return { headers: hdrs, rows: rws };
    }, [projectionOutputs, viewMode]);

    if (!projectionOutputs) {
        return <div className="p-8 text-center text-slate-500">Calculating Engine Outputs...</div>;
    }

    return (
        <div className="space-y-4">
            <div className="flex justify-between items-center text-sm">
                <p className="text-slate-500">Showing Operating Sales (First 36 Months) derived from Revenue Inputs & Branch Rollout.</p>
                <div className="bg-slate-100 p-1 rounded-lg flex space-x-1">
                    <button onClick={() => setViewMode('monthly')} className={`px-3 py-1 rounded-md ${viewMode === 'monthly' ? 'bg-white shadow-sm font-semibold' : 'text-slate-500 hover:bg-slate-200'}`}>Monthly</button>
                    {/* Add Yearly button implementation later */}
                </div>
            </div>
            <DataGrid headers={headers} rows={rows} onCellEdit={() => { }} />
        </div>
    );
}
