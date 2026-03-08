'use client';
import React from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

export default function OpexSheet() {
    const { opex, updateOpexCell } = useFinancial();

    return (
        <DataGrid
            headers={[
                { key: 'category', label: 'Category', editable: true, sticky: true },
                { key: 'subCategory', label: 'Sub-Category', editable: true },
                { key: 'units', label: 'Headcount/Units', editable: true, type: 'number', inputType: 'number' },
                { key: 'price', label: 'Monthly Cost', editable: true, type: 'currency', inputType: 'number' }
            ]}
            rows={opex}
            onCellEdit={(r, c, v) => updateOpexCell(opex[r].id, c, v)}
        />
    );
}
