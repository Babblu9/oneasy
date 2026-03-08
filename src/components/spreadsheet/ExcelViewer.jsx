'use client';

import React, { useState, useEffect, useRef, useMemo } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { RefreshCw, AlertCircle, FileSpreadsheet, Maximize2, Info } from 'lucide-react';

// ── ARGB color helper ────────────────────────────────────────────────────────
function argbToHex(argb) {
    if (!argb || argb.length < 6) return null;
    // ARGB format: AARRGGBB
    const hex = argb.startsWith('FF') ? '#' + argb.slice(2) : '#' + argb.slice(2);
    return hex;
}

function isDark(hex) {
    if (!hex || hex.length < 7) return false;
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return (r * 0.299 + g * 0.587 + b * 0.114) < 128;
}

// ── Loading skeleton ─────────────────────────────────────────────────────────
function ExcelSkeleton() {
    return (
        <div className="animate-pulse p-2">
            {/* Header row */}
            <div className="flex gap-1 mb-2">
                {[60, 120, 100, 90, 80, 110, 75].map((w, i) => (
                    <div key={i} className="h-7 bg-slate-300 rounded" style={{ width: w }} />
                ))}
            </div>
            {/* Data rows */}
            {Array.from({ length: 12 }).map((_, row) => (
                <div key={row} className="flex gap-1 mb-1">
                    {[60, 120, 100, 90, 80, 110, 75].map((w, i) => (
                        <div key={i} className="h-6 bg-slate-100 rounded" style={{ width: w, opacity: 1 - row * 0.06 }} />
                    ))}
                </div>
            ))}
        </div>
    );
}

// ── Single cell ──────────────────────────────────────────────────────────────
function ExcelCell({ cell, isFlashing, colWidth }) {
    const bgHex = cell.bgColor ? argbToHex(cell.bgColor) : null;
    const fgHex = cell.color ? argbToHex(cell.color) : null;
    const textColor = fgHex || (bgHex && isDark(bgHex) ? '#ffffff' : '#1e293b');

    const style = {
        width: `${colWidth * 7.5}px`,
        minWidth: `${colWidth * 7.5}px`,
        maxWidth: `${colWidth * 7.5}px`,
        backgroundColor: bgHex || 'transparent',
        color: textColor,
        fontWeight: cell.bold ? '600' : '400',
        fontStyle: cell.italic ? 'italic' : 'normal',
        textAlign: cell.align === 'center' ? 'center' : cell.align === 'right' ? 'right' : 'left',
        borderBottom: cell.hasBorder ? '1px solid #cbd5e1' : '1px solid #f1f5f9',
        whiteSpace: cell.wrap ? 'normal' : 'nowrap',
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        fontSize: '12px',
        padding: '2px 6px',
        lineHeight: '1.4',
        height: '26px',
        position: 'relative',
    };

    if (isFlashing) {
        style.backgroundColor = '#bbf7d0';
        style.boxShadow = 'inset 0 0 0 2px #22c55e';
        style.transition = 'background-color 0.3s, box-shadow 0.3s';
    }

    return (
        <td style={style} title={cell.value}>
            {cell.value}
            {isFlashing && (
                <motion.span
                    initial={{ opacity: 1 }}
                    animate={{ opacity: 0 }}
                    transition={{ delay: 1.5, duration: 0.5 }}
                    style={{
                        position: 'absolute',
                        top: 1,
                        right: 2,
                        fontSize: '8px',
                        color: '#16a34a',
                        fontWeight: 'bold',
                        pointerEvents: 'none',
                    }}
                >
                    ✓
                </motion.span>
            )}
        </td>
    );
}

// ── Main ExcelViewer ─────────────────────────────────────────────────────────
export default function ExcelViewer({ sheet, flashCells = [], patchVersion = 0 }) {
    const [data, setData] = useState(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState(null);
    const [flashingCells, setFlashingCells] = useState(new Set());
    const tableRef = useRef(null);
    const prevSheetRef = useRef(null);
    const prevPatchRef = useRef(0);

    // Flash cells when patches arrive
    useEffect(() => {
        if (patchVersion > prevPatchRef.current && flashCells.length > 0) {
            prevPatchRef.current = patchVersion;
            const relevant = flashCells
                .filter(fc => fc.sheet === (data?.sheet || ''))
                .map(fc => fc.cell);
            if (relevant.length > 0) {
                setFlashingCells(new Set(relevant));
                setTimeout(() => setFlashingCells(new Set()), 2500);
            }
        }
    }, [patchVersion, flashCells, data?.sheet]);

    // Load sheet data when sheet or patchVersion changes
    useEffect(() => {
        const sheetChanged = sheet !== prevSheetRef.current;
        prevSheetRef.current = sheet;

        let cancelled = false;
        setLoading(true);
        if (sheetChanged) setData(null);
        setError(null);

        fetch(`/api/excel-workbook?sheet=${encodeURIComponent(sheet)}`)
            .then(r => r.json())
            .then(json => {
                if (cancelled) return;
                if (json.error) {
                    setError(json.error);
                    setData(null);
                } else {
                    setData(json);
                }
                setLoading(false);
            })
            .catch(err => {
                if (!cancelled) {
                    setError(err.message);
                    setLoading(false);
                }
            });

        return () => { cancelled = true; };
    }, [sheet, patchVersion]);

    // Determine if a cell address is in the flash set
    function cellIsFlashing(rowNum, colNum) {
        // Convert column number to letter
        const letter = String.fromCharCode(64 + colNum);
        const addr = `${letter}${rowNum}`;
        return flashingCells.has(addr);
    }

    const defaultColWidth = 12;

    return (
        <div className="flex flex-col h-full bg-white">
            {/* Sheet info bar */}
            <div className="flex items-center justify-between px-3 py-1.5 bg-slate-50 border-b border-slate-200 shrink-0">
                <div className="flex items-center gap-2 text-xs text-slate-500">
                    <FileSpreadsheet size={12} className="text-green-600" />
                    <span className="font-medium text-slate-700">Docty Healthcare - Business Plan.xlsx</span>
                    <span className="text-slate-400">›</span>
                    <span className="text-blue-600 font-medium">{data?.sheet || sheet}</span>
                    {data?.truncated && (
                        <span className="flex items-center gap-1 text-amber-600 text-[10px] font-medium ml-1">
                            <Info size={10} />
                            Showing first {data?.rows?.length} rows
                        </span>
                    )}
                </div>
                <div className="flex items-center gap-2">
                    {loading && <RefreshCw size={11} className="animate-spin text-blue-500" />}
                    {data && (
                        <span className="text-[10px] text-slate-400">
                            {data.totalRows?.toLocaleString()} rows × {data.totalCols} cols
                        </span>
                    )}
                </div>
            </div>

            {/* Table area */}
            <div className="flex-1 overflow-auto" ref={tableRef}>
                <AnimatePresence mode="wait">
                    {loading && !data ? (
                        <motion.div
                            key="skeleton"
                            initial={{ opacity: 0 }}
                            animate={{ opacity: 1 }}
                            exit={{ opacity: 0 }}
                        >
                            <ExcelSkeleton />
                        </motion.div>
                    ) : error ? (
                        <motion.div
                            key="error"
                            initial={{ opacity: 0 }}
                            animate={{ opacity: 1 }}
                            className="flex flex-col items-center justify-center h-64 gap-3 text-slate-500"
                        >
                            <AlertCircle size={32} className="text-amber-400" />
                            <p className="text-sm font-medium">Could not load sheet</p>
                            <p className="text-xs text-slate-400 max-w-xs text-center">{error}</p>
                        </motion.div>
                    ) : data ? (
                        <motion.div
                            key={`${sheet}-${patchVersion}`}
                            initial={{ opacity: 0 }}
                            animate={{ opacity: 1 }}
                            transition={{ duration: 0.2 }}
                        >
                            <table
                                className="border-collapse text-left select-text"
                                style={{ tableLayout: 'fixed', fontSize: '12px' }}
                            >
                                {/* Column widths */}
                                <colgroup>
                                    {/* Row number column */}
                                    <col style={{ width: '36px' }} />
                                    {/* Data columns */}
                                    {Array.from({ length: data.colWidths?.length || 20 }).map((_, i) => (
                                        <col key={i} style={{ width: `${(data.colWidths?.[i] || defaultColWidth) * 7.5}px` }} />
                                    ))}
                                </colgroup>
                                <thead className="sticky top-0 z-10">
                                    {/* Column letters header */}
                                    <tr>
                                        <th className="bg-slate-200 border border-slate-300 text-center text-[10px] text-slate-500 font-medium" style={{ width: 36, height: 22 }} />
                                        {Array.from({ length: data.colWidths?.length || 20 }).map((_, i) => (
                                            <th key={i}
                                                className="bg-slate-200 border border-slate-300 text-center text-[10px] text-slate-500 font-medium px-1"
                                                style={{ height: 22 }}
                                            >
                                                {String.fromCharCode(65 + i)}
                                            </th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {data.rows.map(({ row: rowNum, cells }) => (
                                        <tr key={rowNum} className="hover:bg-blue-50/30 transition-colors">
                                            {/* Row number */}
                                            <td className="bg-slate-100 border border-slate-200 text-center text-[10px] text-slate-400 font-medium sticky left-0 z-10"
                                                style={{ width: 36, height: 26, minWidth: 36, padding: '0 2px' }}>
                                                {rowNum}
                                            </td>
                                            {/* Fill missing cells */}
                                            {Array.from({ length: (data.colWidths?.length || 20) }).map((_, colIdx) => {
                                                const colNum = colIdx + 1;
                                                const cell = cells.find(c => c.col === colNum) || { col: colNum, value: '', bold: false };
                                                const isFlash = cellIsFlashing(rowNum, colNum);
                                                return (
                                                    <ExcelCell
                                                        key={colNum}
                                                        cell={cell}
                                                        isFlashing={isFlash}
                                                        colWidth={data.colWidths?.[colIdx] || defaultColWidth}
                                                    />
                                                );
                                            })}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </motion.div>
                    ) : null}
                </AnimatePresence>
            </div>
        </div>
    );
}
