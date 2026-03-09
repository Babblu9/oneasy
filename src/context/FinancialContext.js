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

    // ---- AI GENERATED STREAMS (Context-Aware Auto-fill) ----
    const [generatedStreams, setGeneratedStreams] = useState({ revenue: [], opex: [] });

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

    // ---- UNIVERSAL MULTI-TEMPLATE STATE ----
    const [templateId, setTemplateId] = useState('healthcare');
    const [universalAssumptions, setUniversalAssumptions] = useState({});
    const [allTemplates, setAllTemplates] = useState([]);

    // Get active template metadata
    const activeTemplate = allTemplates.find(t => t.id === templateId) || { name: 'Healthcare Clinic', icon: '🏥' };



    // ---- COLLECTION PROGRESS (Universal 5-Step Flow) ----
    const collectionProgress = {
        type: !!templateId && templateId !== 'loading',
        revenue: !!universalAssumptions.revenue_model,
        price: !!(universalAssumptions.avg_ticket || universalAssumptions.course_price || universalAssumptions.mrr_per_customer || universalAssumptions.avg_order_value || universalAssumptions.daily_sales || universalAssumptions.avg_project_value || universalAssumptions.price_per_unit),
        volume: !!(universalAssumptions.patients_per_month || universalAssumptions.students_per_month || universalAssumptions.customers_per_month || universalAssumptions.orders_per_month || universalAssumptions.daily_sales || universalAssumptions.clients_per_month || universalAssumptions.units_per_month),
        team: !!universalAssumptions.employees,
        funding: !!universalAssumptions.funding,
    };

    // ---- ACTIVE SPREADSHEET TAB (controlled by AI) ----
    const [activeSpreadsheetTab, setActiveSpreadsheetTab] = useState('Loading...');
    const [flashingTab, setFlashingTab] = useState(null);

    // ---- ENGINE OUTPUTS ----
    const [projectionOutputs, setProjectionOutputs] = useState(null);

    // ---- EXCEL PATCHES (real Excel cell updates) ----
    const [excelPatches, setExcelPatches] = useState([]);
    const [excelPatchVersion, setExcelPatchVersion] = useState(0);
    const [injectionReport, setInjectionReport] = useState(null);
    const patchQueueRef = useRef([]);
    const patchTimerRef = useRef(null);
    const didResetRef = useRef(false);

    // ---- LIVE EDIT RELOAD: listen for AI editCell events ----
    useEffect(() => {
        const handler = () => setExcelPatchVersion(v => v + 1);
        window.addEventListener('excel-cell-edited', handler);
        return () => window.removeEventListener('excel-cell-edited', handler);
    }, []);

    // Fetch available templates on load
    useEffect(() => {
        fetch('/api/select-template')
            .then(res => res.json())
            .then(data => {
                if (data.templates) setAllTemplates(data.templates);
            })
            .catch(err => console.error("Failed to fetch templates:", err));
    }, []);

    // ---- AUTO-RESET Excel on first mount (gives a blank slate) + Fetch INITIAL SHEETS ----
    useEffect(() => {
        if (didResetRef.current) return;
        didResetRef.current = true;

        const fetchSheets = () => {
            fetch('/api/excel-sheets')
                .then(res => res.json())
                .then(data => {
                    if (data.sheets && data.sheets.length > 0) {
                        if (data.sheets.includes('2. Basics')) {
                            setActiveSpreadsheetTab('2. Basics');
                        } else if (data.sheets.includes('Assumptions')) {
                            setActiveSpreadsheetTab('Assumptions');
                        } else {
                            setActiveSpreadsheetTab(data.sheets[0]);
                        }
                    }
                })
                .catch(err => {
                    console.error("Failed to fetch initial sheets:", err);
                    setActiveSpreadsheetTab('Error Loading Sheets');
                });
        };

        fetch('/api/excel-reset', { method: 'POST' })
            .then(r => r.json())
            .then(d => {
                if (d.success) {
                    setExcelPatchVersion(v => v + 1);
                    fetchSheets(); // Fetch sheets ONLY after reset is done
                }
            })
            .catch(err => {
                console.warn('excel-reset failed (non-critical):', err);
                fetchSheets(); // Fallback to fetch sheets anyway
            });
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
                const body = await res.json().catch(() => ({}));
                if (body && (body.report || body.errors)) {
                    setInjectionReport({
                        timestamp: Date.now(),
                        report: body.report || null,
                        errors: Array.isArray(body.errors) ? body.errors : [],
                        patchedCount: Number(body.patchedCount) || 0,
                    });
                }
                if (!res.ok) {
                    console.warn('Excel fill warning:', body);
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
            const safeUnits = Number(units) || 0;
            const safePrice = Number(price) || 0;
            if (!exists) {
                sub.products.push({
                    id: `prod-${Date.now()}`,
                    name: productName,
                    units: safeUnits,
                    price: safePrice,
                    growthY1: Number(growthY1) || 0.1,
                    growthY2: 0.08,
                    growthY3: 0.05,
                    cellRef: cellRef || null,
                    growthRates: growthRates || null,
                });
            } else {
                // Upsert behavior: if stream already exists with zero placeholders, overwrite with real values.
                if ((Number(exists.units) || 0) <= 0 && safeUnits > 0) exists.units = safeUnits;
                if ((Number(exists.price) || 0) <= 0 && safePrice > 0) exists.price = safePrice;
                if (!exists.cellRef && cellRef) exists.cellRef = cellRef;
                if (!exists.growthRates && growthRates) exists.growthRates = growthRates;
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
            const safeUnits = Number(units) || 0;
            const safePrice = Number(price) || 0;
            if (exists) {
                // Upsert behavior: replace zero placeholder values when better values arrive.
                return prev.map(o => {
                    if (!(o.category.toLowerCase() === category.toLowerCase() &&
                        o.subCategory.toLowerCase() === subCategory.toLowerCase())) return o;
                    return {
                        ...o,
                        units: (Number(o.units) || 0) <= 0 && safeUnits > 0 ? safeUnits : o.units,
                        price: (Number(o.price) || 0) <= 0 && safePrice > 0 ? safePrice : o.price,
                        cellRef: o.cellRef || cellRef || null,
                        growthRates: o.growthRates || growthRates || null,
                    };
                });
            }
            return [...prev, {
                id: `opex-${Date.now()}`,
                category: String(category).substring(0, 100),
                subCategory: String(subCategory).substring(0, 100),
                units: safeUnits,
                price: safePrice,
                cellRef: cellRef || null,
                growthRates: growthRates || null,
            }];
        });
        patchExcel(dataActionToPatches({ type: 'addOpex', category, subCategory, units, price, cellRef, growthRates }));
        flashTab('A.II OPEX');
    }, [flashTab, patchExcel]);

    const removeOpexByName = useCallback((subCategory) => {
        const target = String(subCategory || '').trim().toLowerCase();
        if (!target) return;
        setOpex(prev => prev.filter(o => String(o.subCategory || '').trim().toLowerCase() !== target));
    }, []);

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

    const removeRevenueProductByName = useCallback((productName) => {
        const target = String(productName || '').trim().toLowerCase();
        if (!target) return;

        setRevenueStreams(prev => prev
            .map(stream => ({
                ...stream,
                subStreams: stream.subStreams
                    .map(sub => ({
                        ...sub,
                        products: sub.products.filter(p => String(p.name || '').trim().toLowerCase() !== target)
                    }))
                    .filter(sub => sub.products.length > 0)
            }))
            .filter(stream => stream.subStreams.length > 0)
        );
    }, []);

    // ---- TEMPLATE SWITCH RESET ----
    // Prevent cross-template bleed by clearing in-memory model state
    // before applying assumptions/streams for the newly selected template.
    const clearSessionForTemplateSwitch = useCallback(() => {
        setBranches([]);
        setBranchCountState(10);
        setRevenueStreams([]);
        setOpex([]);
        setCapex([]);
        setGeneratedStreams({ revenue: [], opex: [] });
        setAssumptions({
            taxRate: 0,
            inflationRate: 0,
            initialInvestment: 0
        });
        setFundingState({
            loanAmount: 0,
            interestRate: 0,
            loanTenureMonths: 0,
            moratoriumMonths: 0,
            equityFromPromoters: 0,
            grantAmount: 0,
        });
        setInjectionReport(null);
        setExcelPatches([]);
    }, []);

    // ---- GENERATED STREAMS ACTIONS ----
    const setGeneratedStreamsViaChat = useCallback((payload) => {
        setGeneratedStreams(payload);
    }, []);

    const revenueMonthlyTotal = revenueStreams.reduce((sum, stream) => {
        const streamTotal = stream.subStreams.reduce((subSum, sub) => {
            const subTotal = sub.products.reduce((prodSum, prod) =>
                prodSum + ((Number(prod.units) || 0) * (Number(prod.price) || 0)), 0);
            return subSum + subTotal;
        }, 0);
        return sum + streamTotal;
    }, 0);

    const opexMonthlyTotal = opex.reduce((sum, item) =>
        sum + ((Number(item.units) || 0) * (Number(item.price) || 0)), 0);

    const netMonthlyTotal = revenueMonthlyTotal - opexMonthlyTotal;

    return (
        <FinancialContext.Provider value={{
            // Business Identity
            businessInfo, setBusinessInfo,

            // Branches
            branches, setBranches, updateBranchCell, setBranchesViaChat,
            branchCount, setBranchCount,

            // Revenue
            revenueStreams, setRevenueStreams, updateProductCell, addProductViaChat, removeRevenueProductByName,

            // OPEX
            opex, setOpex, updateOpexCell, addOpexViaChat, removeOpexByName,

            // CAPEX
            capex, setCapex, updateCapexCell, addCapexViaChat,

            // Assumptions
            assumptions, setAssumptions: setAssumptionsWithPatch,

            // Funding
            funding, setFunding,

            // Progress
            collectionProgress, markComplete, getCompletionPercent,

            // AI Generated Streams
            generatedStreams, setGeneratedStreams, setGeneratedStreamsViaChat,

            // Calculated totals
            revenueMonthlyTotal, opexMonthlyTotal, netMonthlyTotal,

            // Spreadsheet navigation
            activeSpreadsheetTab, setActiveSpreadsheetTab, flashingTab, navigateToTab,

            // Multi-Template
            templateId, setTemplateId,
            universalAssumptions, setUniversalAssumptions,
            allTemplates, activeTemplate,
            clearSessionForTemplateSwitch,

            // Excel patch tracking
            excelPatches, excelPatchVersion, setExcelPatchVersion, patchExcel,
            injectionReport, setInjectionReport,
        }}>
            {children}
        </FinancialContext.Provider>
    );
}

export const useFinancial = () => useContext(FinancialContext);
