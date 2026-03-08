'use client';

import React, { useState } from 'react';
import { motion, AnimatePresence } from 'framer-motion';

export default function DataGrid({ headers, rows, onCellEdit, groupingKey = null }) {
    const [editingCell, setEditingCell] = useState(null);
    const [editValue, setEditValue] = useState('');

    const startEdit = (rowIndex, colIndex, currentValue, isEditable) => {
        if (!isEditable) return;
        setEditingCell({ r: rowIndex, c: colIndex });
        setEditValue(currentValue);
    };

    const commitEdit = (rowIndex, colKey) => {
        if (editingCell) {
            onCellEdit(rowIndex, colKey, editValue);
            setEditingCell(null);
        }
    };

    const handleKeyDown = (e, rowIndex, colKey) => {
        if (e.key === 'Enter') {
            commitEdit(rowIndex, colKey);
        } else if (e.key === 'Escape') {
            setEditingCell(null);
        }
    };

    const formatValue = (val, col) => {
        if (val === null || val === undefined) return '';
        if (col.type === 'currency' || col.prefix === '$') {
            return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(val);
        }
        if (col.type === 'percent') {
            return `${(val * 100).toFixed(1)}%`;
        }
        if (col.type === 'number') {
            return new Intl.NumberFormat('en-US').format(val);
        }
        return val;
    };

    return (
        <div className="overflow-x-auto rounded-xl border border-slate-200 bg-white shadow-sm w-full max-w-full">
            <table className="w-full text-left border-collapse text-sm whitespace-nowrap">
                <thead>
                    <tr className="bg-slate-100/50 border-b border-slate-200">
                        {headers.map((h, i) => (
                            <th
                                key={i}
                                className={`py-3 px-4 font-semibold text-slate-600 tracking-wide text-xs ${h.sticky ? 'sticky left-0 bg-slate-100 z-10 border-r border-slate-200' : ''}`}
                            >
                                {h.label}
                            </th>
                        ))}
                    </tr>
                </thead>
                <tbody className="align-top divide-y divide-slate-100">
                    <AnimatePresence>
                        {rows.length === 0 ? (
                            <tr>
                                <td colSpan={headers.length} className="text-center py-8 text-slate-400">
                                    No data available.
                                </td>
                            </tr>
                        ) : (
                            rows.map((row, rowIndex) => (
                                <motion.tr
                                    initial={{ opacity: 0 }}
                                    animate={{ opacity: 1 }}
                                    exit={{ opacity: 0 }}
                                    transition={{ duration: 0.1 }}
                                    key={row.id || rowIndex}
                                    className={`hover:bg-blue-50/30 transition-colors group ${row.isSummary ? 'bg-slate-50 font-semibold' : ''}`}
                                >
                                    {headers.map((col, colIndex) => {
                                        const isEditing = editingCell?.r === rowIndex && editingCell?.c === colIndex;
                                        const value = row[col.key];

                                        return (
                                            <td
                                                key={colIndex}
                                                className={`py-3 px-4 relative 
                                                ${col.editable && !row.isReadOnly ? 'cursor-pointer hover:bg-blue-50/50' : 'bg-transparent'} 
                                                ${col.sticky ? 'sticky left-0 z-10 border-r border-slate-100' : ''}
                                                ${col.sticky ? (row.isSummary ? 'bg-slate-50' : 'bg-white') : ''}
                                                `}
                                                onClick={() => startEdit(rowIndex, colIndex, value, col.editable && !row.isReadOnly)}
                                            >
                                                {isEditing ? (
                                                    <input
                                                        autoFocus
                                                        type={col.inputType || 'text'}
                                                        className="w-full bg-white border-primary outline-none rounded p-1 absolute inset-y-1 inset-x-1 shadow-md z-20"
                                                        value={editValue}
                                                        onChange={(e) => setEditValue(e.target.value)}
                                                        onBlur={() => commitEdit(rowIndex, col.key)}
                                                        onKeyDown={(e) => handleKeyDown(e, rowIndex, col.key)}
                                                    />
                                                ) : (
                                                    <span className={`
                                                        ${col.editable && !row.isReadOnly ? 'border-b border-dashed border-primary/30 text-primary group-hover:border-primary' : 'text-slate-700'}
                                                    `}>
                                                        {formatValue(value, col)}
                                                    </span>
                                                )}
                                            </td>
                                        )
                                    })}
                                </motion.tr>
                            ))
                        )}
                    </AnimatePresence>
                </tbody>
            </table>
        </div>
    );
}
