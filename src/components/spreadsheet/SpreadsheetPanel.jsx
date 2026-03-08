'use client';

import React, { useCallback } from 'react';
import { useFinancial } from '@/context/FinancialContext';
import ExcelViewer from './ExcelViewer';
import { Download, RefreshCw } from 'lucide-react';
import { motion } from 'framer-motion';

const categories = {
    'Inputs': ['2. Basics', 'Branch', 'A.I Revenue Streams', 'A.II OPEX', 'A.III CAPEX'],
    'Operational Schedules': ['B.I Sales - P1', 'B.II OPEX', 'FA Schedule'],
    'Financial Statements': ['1. P&L', '5. Balance Sheet', '6. Ratios', 'DSCR'],
};

const tabToCategory = {};
Object.entries(categories).forEach(([cat, tabs]) => {
    tabs.forEach(tab => { tabToCategory[tab] = cat; });
});

export default function SpreadsheetPanel() {
    const {
        activeSpreadsheetTab,
        setActiveSpreadsheetTab,
        flashingTab,
        excelPatches,
        excelPatchVersion,
    } = useFinancial();

    const activeCategory = tabToCategory[activeSpreadsheetTab] || 'Inputs';

    const handleTabClick = useCallback((tab) => {
        setActiveSpreadsheetTab(tab);
    }, [setActiveSpreadsheetTab]);

    const handleCategoryClick = useCallback((cat) => {
        setActiveSpreadsheetTab(categories[cat][0]);
    }, [setActiveSpreadsheetTab]);

    const handleDownload = useCallback(async () => {
        try {
            const res = await fetch('/api/excel-fill');
            if (!res.ok) {
                alert('Could not prepare download. Please try again.');
                return;
            }
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Docty_Healthcare_Financial_Model.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (err) {
            console.error('Download failed:', err);
            alert('Download failed. Please try again.');
        }
    }, []);

    return (
        <div className="flex flex-col h-full bg-slate-50 relative">
            {/* ── Top bar: Category tabs + Download ── */}
            <div className="flex items-center justify-between px-4 pt-0 bg-slate-800 shrink-0">
                <div className="flex space-x-1 pt-2">
                    {Object.keys(categories).map(cat => (
                        <button
                            key={cat}
                            onClick={() => handleCategoryClick(cat)}
                            className={`px-4 py-2 text-xs font-semibold rounded-t-lg transition-colors ${activeCategory === cat
                                    ? 'bg-white text-slate-800'
                                    : 'text-slate-300 hover:bg-slate-700 rounded-t-lg'
                                }`}
                        >
                            {cat}
                        </button>
                    ))}
                </div>
                {/* Download button */}
                <button
                    onClick={handleDownload}
                    className="flex items-center gap-1.5 px-3 py-1.5 my-1.5 text-xs font-semibold bg-green-500 hover:bg-green-400 text-white rounded-lg transition-colors shadow-sm"
                    title="Download filled Excel workbook"
                >
                    <Download size={12} />
                    Download Excel
                </button>
            </div>

            {/* ── Sub-tabs ── */}
            <div className="flex flex-wrap gap-1.5 px-4 py-2.5 bg-white border-b border-slate-200 sticky top-0 z-20 shadow-sm shrink-0">
                {categories[activeCategory].map(tab => {
                    const isActive = activeSpreadsheetTab === tab;
                    const isFlashing = flashingTab === tab;
                    return (
                        <motion.button
                            key={tab}
                            onClick={() => handleTabClick(tab)}
                            whileTap={{ scale: 0.96 }}
                            className={`
                                px-3 py-1.5 text-xs font-semibold rounded-md border transition-all duration-300
                                ${isActive
                                    ? 'bg-blue-600 border-blue-600 text-white shadow-sm'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50 hover:border-slate-300'
                                }
                                ${isFlashing ? 'ring-2 ring-green-400 ring-offset-1 bg-green-50 border-green-400 text-green-700 scale-105' : ''}
                            `}
                        >
                            {isFlashing ? '✨ ' : ''}{tab}
                        </motion.button>
                    );
                })}
            </div>

            {/* ── Sheet header ── */}
            <div className="px-4 py-2 bg-white border-b border-slate-100 flex items-center justify-between shrink-0">
                <h2 className="font-bold text-sm text-slate-800">{activeSpreadsheetTab}</h2>
                <div className="flex items-center gap-2">
                    {flashingTab === activeSpreadsheetTab && (
                        <motion.span
                            initial={{ opacity: 0, scale: 0.9 }}
                            animate={{ opacity: 1, scale: 1 }}
                            className="text-xs font-semibold text-green-600 bg-green-50 border border-green-200 px-2 py-0.5 rounded-full flex items-center gap-1"
                        >
                            <span className="w-1.5 h-1.5 rounded-full bg-green-500 animate-ping inline-block" />
                            Updated by AI
                        </motion.span>
                    )}
                    {excelPatches.length > 0 && (
                        <span className="text-[10px] text-slate-400">
                            {excelPatches.length} field{excelPatches.length !== 1 ? 's' : ''} filled
                        </span>
                    )}
                </div>
            </div>

            {/* ── Main Excel viewer ── */}
            <div className="flex-1 overflow-hidden">
                <ExcelViewer
                    sheet={activeSpreadsheetTab}
                    flashCells={excelPatches}
                    patchVersion={excelPatchVersion}
                />
            </div>
        </div>
    );
}
