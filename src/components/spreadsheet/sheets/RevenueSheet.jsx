'use client';
import React from 'react';
import { useFinancial } from '@/context/FinancialContext';
import DataGrid from '../DataGrid';

export default function RevenueSheet() {
    const { revenueStreams, updateProductCell } = useFinancial();

    const getRows = () => {
        let rows = [];
        revenueStreams.forEach(stream => {
            stream.subStreams.forEach(sub => {
                sub.products.forEach(prod => {
                    rows.push({
                        id: prod.id,
                        streamId: stream.id,
                        subId: sub.id,
                        prodId: prod.id,
                        streamName: stream.name,
                        subName: sub.name,
                        productName: prod.name,
                        units: prod.units,
                        price: prod.price,
                        total: prod.units * prod.price,
                        isReadOnly: false
                    });
                });
            });
        });
        return rows;
    };

    const handleEdit = (rowIndex, colKey, newValue) => {
        const rows = getRows();
        const row = rows[rowIndex];
        updateProductCell(row.streamId, row.subId, row.prodId, colKey, newValue);
    };

    return (
        <DataGrid
            headers={[
                { key: 'streamName', label: 'Stream', sticky: true },
                { key: 'productName', label: 'Product / Service', editable: true },
                { key: 'units', label: 'Units/Mo', editable: true, type: 'number', inputType: 'number' },
                { key: 'price', label: 'Unit Price', editable: true, type: 'currency', inputType: 'number' },
                { key: 'total', label: 'Monthly Gross', type: 'currency' }
            ]}
            rows={getRows()}
            onCellEdit={handleEdit}
        />
    );
}
