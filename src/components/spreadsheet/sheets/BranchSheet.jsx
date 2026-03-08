'use client';
import React from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

export default function BranchSheet() {
    const { branches, updateBranchCell } = useFinancial();

    return (
        <DataGrid
            headers={[
                { key: 'name', label: 'Branch Name', editable: true },
                { key: 'startMonth', label: 'Start Month', editable: true, type: 'number' },
                { key: 'status', label: 'Status', editable: false }
            ]}
            rows={branches}
            onCellEdit={(rowIndex, colKey, val) => updateBranchCell(branches[rowIndex].id, colKey, val)}
        />
    );
}
