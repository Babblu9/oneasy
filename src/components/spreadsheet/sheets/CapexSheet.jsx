'use client';
import React from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

export default function CapexSheet() {
    const { capex, updateCapexCell } = useFinancial();

    return (
        <DataGrid
            headers={[
                { key: 'name', label: 'Asset Name', editable: true, sticky: true },
                { key: 'baselineValue', label: 'Value ($)', editable: true, type: 'currency', inputType: 'number' },
                { key: 'branchMultiplier', label: 'Per Branch?', editable: true },
                { key: 'depreciationYears', label: 'Useful Life (Yrs)', editable: true, type: 'number', inputType: 'number' }
            ]}
            rows={capex.map(c => ({
                ...c,
                branchMultiplier: c.branchMultiplier ? 'Yes' : 'No'
            }))}
            onCellEdit={(r, c, v) => {
                let val = v;
                if (c === 'branchMultiplier') {
                    val = v.toLowerCase() === 'yes' || v.toLowerCase() === 'true';
                } else if (c === 'baselineValue' || c === 'depreciationYears') {
                    val = Number(v) || 0;
                }
                updateCapexCell(capex[r].id, c, val);
            }}
        />
    );
}
