'use client';

import React, { createContext, useContext, useState, useCallback, useEffect, useRef } from 'react';
import { runProjection } from '../../engine/index.js';
import { dataActionToPatches } from '../lib/excelCellMap.js';

const FinancialContext = createContext();

export function FinancialProvider({ children }) {

    // ---- BUSINESS IDENTITY (Block A of CA Framework) ----
    const [businessInfo, setBusinessInfoState] = useState({
        legalName: '',
        tradeName: '',
        address: '',
        email: '',
        phone: '',
        promoters: [],
        startDate: '',
        description: '',
        companyType: '', // 'Pvt Ltd', 'LLP', 'Sole Prop', 'Partnership'
        equityShares: 0,
        faceValue: 0,
        paidUpCapital: 0,
    });

    // ---- BRANCH DATA ----
    const [branches, setBranches] = useState([]);
    const [branchCount, setBranchCountState] = useState(10); // Master branch count (H7)

    // ---- REVENUE STREAMS (Hierarchical: Stream → Sub-Stream → Product) ----
    const [revenueStreams, setRevenueStreams] = useState([]);

    // ---- OPEX ----
    const [opex, setOpex] = useState([]);

    // ---- CAPEX / Fixed Assets ----
    const [capex, setCapex] = useState([]);

    // ---- BASICS / ASSUMPTIONS ----
    const [assumptions, setAssumptions] = useState({
        taxRate: 0,
        inflationRate: 0,
        initialInvestment: 0
    });

    // ---- FUNDING / LOAN STRUCTURE (Block D of CA Framework) ----
    const [funding, setFundingState] = useState({
        loanAmount: 0,
        interestRate: 0,
        loanTenureMonths: 0,
        moratoriumMonths: 0,
        equityFromPromoters: 0,
        grantAmount: 0,
    });

    // ---- COLLECTION PROGRESS (derived from real data state) ----
    // This is computed, not stored — it updates automatically as data flows in
    const collectionProgress = {
        basics: !!(businessInfo.legalName || businessInfo.tradeName),
        branches: branches.length > 0,
        revenue: revenueStreams.some(s => s.subStreams?.some(sub => sub.products?.length > 0)),
        opex: opex.length > 0,
        capex: capex.length > 0,
        funding: !!(funding.loanAmount > 0 || funding.equityFromPromoters > 0),
    };

    // ---- ACTIVE SPREADSHEET TAB (controlled by AI) ----
    const [activeSpreadsheetTab, setActiveSpreadsheetTab] = useState('2. Basics');
    const [flashingTab, setFlashingTab] = useState(null);

    // ---- ENGINE OUTPUTS ----
    const [projectionOutputs, setProjectionOutputs] = useState(null);

    // ---- EXCEL PATCHES (real Excel cell updates) ----
    const [excelPatches, setExcelPatches] = useState([]);
    const [excelPatchVersion, setExcelPatchVersion] = useState(0);
    const patchQueueRef = useRef([]);
    const patchTimerRef = useRef(null);
    const didResetRef = useRef(false);

    // ---- AUTO-RESET Excel on first mount (gives a blank slate) ----
    useEffect(() => {
        if (didResetRef.current) return;
        didResetRef.current = true;
        fetch('/api/excel-reset', { method: 'POST' })
            .then(r => r.json())
            .then(d => { if (d.success) setExcelPatchVersion(v => v + 1); })
            .catch(err => console.warn('excel-reset failed (non-critical):', err));
    }, []);

    // Re-run engine whenever inputs change
    useEffect(() => {
        const engineInputs = {
            branchSchedule: branches,
            services: revenueStreams.flatMap(stream =>
                stream.subStreams.flatMap(sub =>
                    sub.products.map(p => ({
                        id: p.id,
                        name: p.name,
                        category: stream.name,
                        baseVolume: Number(p.units) || 0,
                        unitPrice: Number(p.price) || 0,
                        growthRate: Number(p.growthY1) || 0.1,
                    }))
                )
            ),
            expenses: opex.map(o => ({
                id: o.id,
                name: o.category + ' - ' + o.subCategory,
                type: 'fixed',
                amount: Number(o.price) * (Number(o.units) || 1)
            })),
            assets: capex.map(c => ({
                id: c.id,
                name: c.name,
                value: Number(c.baselineValue) || 0,
                usefulLife: Number(c.depreciationYears) || 5
            })),
            taxRate: assumptions.taxRate
        };
        // Don't run the engine until at least one branch exists
        if (!branches || branches.length === 0) return;
        try {
            const result = runProjection(engineInputs, 72);
            setProjectionOutputs(result);
        } catch (e) {
            console.error('Engine Calculation Failed:', e);
        }
    }, [branches, revenueStreams, opex, capex, assumptions]);


    // ---- FLASH ANIMATION HELPER ----
    const flashTab = useCallback((tabName) => {
        setFlashingTab(tabName);
        setTimeout(() => setFlashingTab(null), 1500);
    }, []);

    // Navigate spreadsheet tab programmatically (AI-driven)
    const navigateToTab = useCallback((tabName) => {
        setActiveSpreadsheetTab(tabName);
        flashTab(tabName);
    }, [flashTab]);

    // ---- PATCH EXCEL (real Excel write-back) --------------------------------
    // Debounced: batches patches within 300ms and sends one API call
    const patchExcel = useCallback((patches) => {
        if (!patches || patches.length === 0) return;

        const stamped = patches.map(p => ({ ...p, timestamp: Date.now() }));
        patchQueueRef.current = [...patchQueueRef.current, ...stamped];

        if (patchTimerRef.current) clearTimeout(patchTimerRef.current);
        patchTimerRef.current = setTimeout(async () => {
            const batch = patchQueueRef.current;
            patchQueueRef.current = [];
            if (batch.length === 0) return;

            // Update local flash state
            setExcelPatches(prev => [...prev, ...batch]);
            setExcelPatchVersion(v => v + 1);

            // Write to real Excel on server
            try {
                const res = await fetch('/api/excel-fill', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ patches: batch }),
                });
                if (!res.ok) {
                    const err = await res.json();
                    console.warn('Excel fill warning:', err);
                }
            } catch (err) {
                console.error('Excel fill error:', err);
            }
        }, 300);
    }, []);

    // ---- BUSINESS INFO ACTIONS ----
    const setBusinessInfo = useCallback((patch) => {
        setBusinessInfoState(prev => ({ ...prev, ...patch }));
        // Derive and send Excel patches
        patchExcel(dataActionToPatches({ type: 'setBusinessInfo', ...patch }));
    }, [patchExcel]);

    // ---- FUNDING ACTIONS ----
    const setFunding = useCallback((patch) => {
        setFundingState(prev => ({ ...prev, ...patch }));
        patchExcel(dataActionToPatches({ type: 'setFunding', ...patch }));
    }, [patchExcel]);

    // ---- ASSUMPTIONS (also patch Excel) ----
    const setAssumptionsWithPatch = useCallback((updater) => {
        setAssumptions(prev => {
            const next = typeof updater === 'function' ? updater(prev) : { ...prev, ...updater };
            patchExcel(dataActionToPatches({ type: 'setAssumptions', ...next }));
            return next;
        });
    }, [patchExcel]);

    // ---- PROGRESS ACTIONS (kept for AI markComplete DATA tags, no-op now since derived) ----
    const markComplete = useCallback((_block) => { /* derived from state now */ }, []);

    const getCompletionPercent = useCallback(() => {
        const blocks = Object.values(collectionProgress);
        const done = blocks.filter(Boolean).length;
        return Math.round((done / blocks.length) * 100);
    }, [collectionProgress]);

    // ---- REVENUE STREAM ACTIONS ----
    const updateProductCell = useCallback((streamId, subId, prodId, field, value) => {
        setRevenueStreams(prev => prev.map(stream => {
            if (stream.id !== streamId) return stream;
            return {
                ...stream,
                subStreams: stream.subStreams.map(sub => {
                    if (sub.id !== subId) return sub;
                    return {
                        ...sub,
                        products: sub.products.map(prod => {
                            if (prod.id !== prodId) return prod;
                            return { ...prod, [field]: value };
                        })
                    };
                })
            };
        }));
    }, []);

    const addProductViaChat = useCallback((streamName, subName, productName, units, price, growthY1 = 0.1, cellRef = null, growthRates = null) => {
        setRevenueStreams(prev => {
            let newStreams = [...prev];
            let stream = newStreams.find(s => s.name.toLowerCase() === streamName.toLowerCase());
            if (!stream) {
                stream = { id: `stream-${Date.now()}`, name: streamName, subStreams: [] };
                newStreams.push(stream);
            }
            let sub = stream.subStreams.find(s => s.name.toLowerCase() === subName.toLowerCase());
            if (!sub) {
                sub = { id: `sub-${Date.now()}`, name: subName, products: [] };
                stream.subStreams.push(sub);
            }
            const exists = sub.products.find(p => p.name.toLowerCase() === productName.toLowerCase());
            if (!exists) {
                sub.products.push({
                    id: `prod-${Date.now()}`,
                    name: productName,
                    units: Number(units) || 0,
                    price: Number(price) || 0,
                    growthY1: Number(growthY1) || 0.1,
                    growthY2: 0.08,
                    growthY3: 0.05,
                    cellRef: cellRef || null,
                    growthRates: growthRates || null,
                });
            }
            return newStreams;
        });
        // Also patch the real Excel — pass cellRef and growthRates for exact cell patching
        patchExcel(dataActionToPatches({ type: 'addRevenueStream', streamName, subName, productName, units, price, growthY1, cellRef, growthRates }));
        flashTab('A.I Revenue Streams');
    }, [flashTab, patchExcel]);

    // ---- OPEX ACTIONS ----
    const updateOpexCell = useCallback((opexId, field, value) => {
        setOpex(prev => prev.map(o => {
            if (o.id !== opexId) return o;
            return { ...o, [field]: value };
        }));
    }, []);

    const addOpexViaChat = useCallback((category, subCategory, units, price, cellRef = null, growthRates = null) => {
        setOpex(prev => {
            const exists = prev.find(o =>
                o.category.toLowerCase() === category.toLowerCase() &&
                o.subCategory.toLowerCase() === subCategory.toLowerCase()
            );
            if (exists) return prev;
            return [...prev, {
                id: `opex-${Date.now()}`,
                category: String(category).substring(0, 100),
                subCategory: String(subCategory).substring(0, 100),
                units: Number(units) || 0,
                price: Number(price) || 0,
                cellRef: cellRef || null,
                growthRates: growthRates || null,
            }];
        });
        patchExcel(dataActionToPatches({ type: 'addOpex', category, subCategory, units, price, cellRef, growthRates }));
        flashTab('A.II OPEX');
    }, [flashTab, patchExcel]);

    // ---- BRANCH COUNT ACTION ----
    const setBranchCount = useCallback((count) => {
        const safeCount = Math.max(1, Math.min(100, Number(count) || 1));
        setBranchCountState(safeCount);
        patchExcel(dataActionToPatches({ type: 'setBranchCount', count: safeCount }));
        flashTab('Branch');
    }, [flashTab, patchExcel]);

    // ---- BRANCH ACTIONS ----
    const updateBranchCell = useCallback((branchId, field, value) => {
        setBranches(prev => prev.map(b => {
            if (b.id !== branchId) return b;
            return { ...b, [field]: value };
        }));
    }, []);

    const setBranchesViaChat = useCallback((newBranches) => {
        setBranches(newBranches);
        // Flash the Branch tab
        flashTab('Branch');
        // Build patches: write each branch's start month into Branch sheet
        // Branch sheet col A = name, col B = month offset, rows from 4
        const patches = newBranches.flatMap((b, i) => [
            { sheet: 'Branch', cell: `A${4 + i}`, value: b.name },
            { sheet: 'Branch', cell: `B${4 + i}`, value: b.startMonth },
        ]);
        patchExcel(patches);
    }, [flashTab, patchExcel]);

    // ---- CAPEX ACTIONS ----
    const updateCapexCell = useCallback((capexId, field, value) => {
        setCapex(prev => prev.map(c => {
            if (c.id !== capexId) return c;
            return { ...c, [field]: value };
        }));
    }, []);

    const addCapexViaChat = useCallback((name, value, usefulLife) => {
        setCapex(prev => {
            const exists = prev.find(c => c.name.toLowerCase() === name.toLowerCase());
            if (exists) return prev;
            return [...prev, {
                id: `capex-${Date.now()}`,
                name: String(name).substring(0, 150),
                baselineValue: Number(value) || 0,
                branchMultiplier: true,
                depreciationYears: Number(usefulLife) || 5
            }];
        });
        patchExcel(dataActionToPatches({ type: 'addCapex', name, cost: value, usefulLife }));
        flashTab('A.III CAPEX');
    }, [flashTab, patchExcel]);

    return (
        <FinancialContext.Provider value={{
            // Business Identity
            businessInfo, setBusinessInfo,

            // Branches
            branches, setBranches, updateBranchCell, setBranchesViaChat,
            branchCount, setBranchCount,

            // Revenue
            revenueStreams, setRevenueStreams, updateProductCell, addProductViaChat,

            // OPEX
            opex, setOpex, updateOpexCell, addOpexViaChat,

            // CAPEX
            capex, setCapex, updateCapexCell, addCapexViaChat,

            // Assumptions
            assumptions, setAssumptions: setAssumptionsWithPatch,

            // Funding
            funding, setFunding,

            // Progress
            collectionProgress, markComplete, getCompletionPercent,

            // Spreadsheet navigation
            activeSpreadsheetTab, setActiveSpreadsheetTab, flashingTab, navigateToTab,

            // Engine outputs
            projectionOutputs,

            // Excel patch tracking
            excelPatches, excelPatchVersion, patchExcel,
        }}>
            {children}
        </FinancialContext.Provider>
    );
}

export const useFinancial = () => useContext(FinancialContext);
