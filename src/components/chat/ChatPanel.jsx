'use client';

import React, { useEffect, useRef, useState, useCallback } from 'react';
import { useChat } from '@ai-sdk/react';
import ReactMarkdown from 'react-markdown';
import { User, Bot, Send, Activity, CheckCircle, Zap, ChevronRight, Download } from 'lucide-react';
import { useFinancial } from '@/context/FinancialContext';
import { motion, AnimatePresence } from 'framer-motion';

// ─── Extract text from any message format ──────────────────────────────────
function extractText(msg) {
    if (Array.isArray(msg.parts) && msg.parts.length > 0) {
        return msg.parts.filter(p => p.type === 'text').map(p => p.text).join('');
    }
    if (typeof msg.content === 'string') return msg.content;
    return '';
}

// ─── Strip DATA/CHIPS tags from display text ───────────────────────────────
function cleanForDisplay(raw) {
    return raw
        .replace(/\[DATA:[^\]]*\]/g, '')
        .replace(/\[CHIPS:[^\]]*\]/g, '')
        .trim();
}

// ─── Progress Bar ──────────────────────────────────────────────────────────
function ProgressBar({ progress }) {
    const blocks = [
        { key: 'basics', label: 'Basics' },
        { key: 'branches', label: 'Branches' },
        { key: 'revenue', label: 'Revenue' },
        { key: 'opex', label: 'OPEX' },
        { key: 'capex', label: 'CAPEX' },
        { key: 'funding', label: 'Funding' },
    ];
    const done = blocks.filter(b => progress[b.key]).length;
    const pct = Math.round((done / blocks.length) * 100);

    return (
        <div className="px-4 py-2 bg-slate-50 border-b border-slate-100">
            <div className="flex items-center justify-between mb-1.5">
                <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wide">Model Completion</span>
                <span className="text-[11px] font-bold text-blue-600">{pct}%</span>
            </div>
            <div className="flex gap-1 mb-1.5">
                {blocks.map(b => (
                    <div key={b.key} title={b.label}
                        className={`flex-1 h-1.5 rounded-full transition-all duration-700 ${progress[b.key] ? 'bg-green-500' : 'bg-slate-200'}`}
                    />
                ))}
            </div>
            <div className="flex gap-1">
                {blocks.map(b => (
                    <div key={b.key} className="flex-1 text-center">
                        <span className={`text-[9px] font-medium ${progress[b.key] ? 'text-green-600' : 'text-slate-400'}`}>
                            {progress[b.key] ? '✓ ' : ''}{b.label}
                        </span>
                    </div>
                ))}
            </div>
        </div>
    );
}

// ─── Data Confirmation Card ─────────────────────────────────────────────────
function DataCard({ card }) {
    const icons = { setBusinessInfo: '🏢', setBranches: '🏥', setBranchCount: '🔢', addRevenueStream: '📈', addOpex: '💸', addCapex: '🔧', setAssumptions: '📊', setFunding: '💰', markComplete: '✅' };
    const labels = { setBusinessInfo: 'Business Info Saved', setBranches: 'Branches Updated', setBranchCount: 'Branch Count Set', addRevenueStream: 'Revenue Stream Added', addOpex: 'Expense Added', addCapex: 'Asset Added', setAssumptions: 'Assumptions Set', setFunding: 'Funding Saved', markComplete: 'Section Complete' };
    const fmt = (n) => Number(n || 0).toLocaleString('en-IN');

    const detail = (() => {
        switch (card.type) {
            case 'addRevenueStream': return `${card.productName} · ${card.units} units/mo · ₹${fmt(card.price)} → ₹${fmt((card.units || 0) * (card.price || 0))}/mo`;
            case 'addOpex': return `${card.category} — ${card.subCategory} · ₹${fmt(card.price)}/mo`;
            case 'addCapex': return `${card.name} · ₹${fmt(card.cost)} · ${card.usefulLife}yr life`;
            case 'setBranches': return `${Array.isArray(card.branches) ? card.branches.length : 1} branch(es) configured`;
            case 'setBranchCount': return `Master branch count set to ${card.count}`;
            case 'setBusinessInfo': return card.legalName || card.tradeName || 'Business details updated';
            case 'setFunding': return `Loan ₹${fmt(card.loanAmount)} · ${card.interestRate}% · ${card.loanTenureMonths}mo`;
            case 'setAssumptions': return `Investment ₹${fmt(card.initialInvestment)} · Tax ${((card.taxRate || 0.25) * 100).toFixed(0)}%`;
            case 'markComplete': return `${card.block} section fully collected`;
            default: return '';
        }
    })();

    return (
        <motion.div initial={{ opacity: 0, scale: 0.95, y: 4 }} animate={{ opacity: 1, scale: 1, y: 0 }}
            className="mt-2 flex items-start gap-2 bg-green-50 border border-green-200 rounded-xl px-3 py-2 text-xs shadow-sm">
            <span className="text-base leading-none mt-0.5">{icons[card.type] || '✓'}</span>
            <div className="flex-1 min-w-0">
                <p className="font-semibold text-green-800">{labels[card.type] || 'Data Saved'}</p>
                <p className="text-green-700 mt-0.5 leading-relaxed truncate">{detail}</p>
            </div>
            <CheckCircle size={14} className="text-green-500 shrink-0 mt-0.5" />
        </motion.div>
    );
}

// ─── Quick Reply Chips ──────────────────────────────────────────────────────
function QuickReplyChips({ chips, onSelect }) {
    if (!chips || chips.length === 0) return null;
    return (
        <motion.div initial={{ opacity: 0, y: 6 }} animate={{ opacity: 1, y: 0 }}
            className="flex flex-wrap gap-2 px-4 pb-2">
            {chips.map((chip, i) => (
                <button key={i} onClick={() => onSelect(chip)}
                    className="flex items-center gap-1 px-3 py-1.5 bg-blue-50 hover:bg-blue-100 border border-blue-200 hover:border-blue-400 text-blue-700 text-sm font-medium rounded-full transition-all">
                    <ChevronRight size={12} />{chip}
                </button>
            ))}
        </motion.div>
    );
}

// ─── Excel Download Button ──────────────────────────────────────────────────
function DownloadButton({ financialData }) {
    const [loading, setLoading] = useState(false);

    const handleDownload = async () => {
        setLoading(true);
        try {
            // Dynamically import ExcelJS to keep bundle lean
            const ExcelJS = (await import('exceljs')).default;
            const wb = new ExcelJS.Workbook();
            wb.creator = 'Docty AI';
            wb.created = new Date();

            const { businessInfo, branches, revenueStreams, opex, capex, assumptions, funding, projectionOutputs } = financialData;

            // ── Sheet 1: Basics ──────────────────────────────────────
            const ws1 = wb.addWorksheet('2. Basics');
            ws1.columns = [
                { header: 'Field', key: 'field', width: 32 },
                { header: 'Value', key: 'value', width: 45 },
            ];
            ws1.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            ws1.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };

            const basicsRows = [
                ['Legal Business Name', businessInfo.legalName],
                ['Trading / Brand Name', businessInfo.tradeName],
                ['Company Type', businessInfo.companyType],
                ['Registered Address', businessInfo.address],
                ['Official Email', businessInfo.email],
                ['Official Phone', businessInfo.phone],
                ['Founders / Promoters', Array.isArray(businessInfo.promoters) ? businessInfo.promoters.join(', ') : ''],
                ['Phase 1 Start Date', businessInfo.startDate],
                ['Business Description', businessInfo.description],
                ['Equity Shares', businessInfo.equityShares],
                ['Face Value per Share (₹)', businessInfo.faceValue],
                ['Paid-Up Capital (₹)', businessInfo.paidUpCapital || (businessInfo.equityShares * businessInfo.faceValue)],
                ['', ''],
                ['Initial Investment (₹)', assumptions.initialInvestment],
                ['Tax Rate', `${(assumptions.taxRate * 100).toFixed(1)}%`],
                ['Inflation Rate', `${(assumptions.inflationRate * 100).toFixed(1)}%`],
                ['', ''],
                ['Loan Amount (₹)', funding.loanAmount],
                ['Interest Rate', `${funding.interestRate}%`],
                ['Loan Tenure (Months)', funding.loanTenureMonths],
                ['Moratorium (Months)', funding.moratoriumMonths],
                ['Equity from Promoters (₹)', funding.equityFromPromoters],
            ];
            basicsRows.forEach(([f, v], i) => {
                const row = ws1.addRow({ field: f, value: v });
                if (i % 2 === 0) row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
            });

            // ── Sheet 2: Branches ──────────────────────────────────
            const ws2 = wb.addWorksheet('Branch');
            ws2.columns = [
                { header: '#', key: 'id', width: 8 },
                { header: 'Branch Name', key: 'name', width: 25 },
                { header: 'Start Month', key: 'startMonth', width: 15 },
                { header: 'Status', key: 'status', width: 15 },
            ];
            ws2.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            ws2.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
            branches.forEach(b => ws2.addRow({ id: b.id, name: b.name, startMonth: b.startMonth, status: b.status }));

            // ── Sheet 3: Revenue Streams ────────────────────────────
            const ws3 = wb.addWorksheet('A.I Revenue Streams');
            ws3.columns = [
                { header: 'Stream', key: 'stream', width: 20 },
                { header: 'Sub-Stream', key: 'sub', width: 20 },
                { header: 'Product / Service', key: 'product', width: 28 },
                { header: 'Units/Month', key: 'units', width: 15 },
                { header: 'Unit Price (₹)', key: 'price', width: 15 },
                { header: 'Monthly Revenue (₹)', key: 'monthly', width: 20 },
                { header: 'Growth Rate Y1', key: 'growth', width: 15 },
            ];
            ws3.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            ws3.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
            revenueStreams.forEach(stream =>
                stream.subStreams.forEach(sub =>
                    sub.products.forEach(p => {
                        ws3.addRow({
                            stream: stream.name, sub: sub.name, product: p.name,
                            units: p.units, price: p.price,
                            monthly: p.units * p.price,
                            growth: `${((p.growthY1 || 0.1) * 100).toFixed(0)}%`
                        });
                    })
                )
            );

            // ── Sheet 4: OPEX ──────────────────────────────────────
            const ws4 = wb.addWorksheet('A.II OPEX');
            ws4.columns = [
                { header: 'Category', key: 'cat', width: 22 },
                { header: 'Sub-Category', key: 'sub', width: 22 },
                { header: 'Units', key: 'units', width: 10 },
                { header: 'Monthly Cost (₹)', key: 'price', width: 18 },
                { header: 'Annual Cost (₹)', key: 'annual', width: 18 },
            ];
            ws4.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            ws4.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
            opex.forEach(o => ws4.addRow({ cat: o.category, sub: o.subCategory, units: o.units, price: o.price, annual: o.price * o.units * 12 }));

            // ── Sheet 5: CAPEX ──────────────────────────────────────
            const ws5 = wb.addWorksheet('A.III CAPEX');
            ws5.columns = [
                { header: 'Asset Name', key: 'name', width: 28 },
                { header: 'Value (₹)', key: 'value', width: 18 },
                { header: 'Useful Life (Yrs)', key: 'life', width: 18 },
                { header: 'Per Branch?', key: 'branch', width: 14 },
                { header: 'Annual Depreciation (₹)', key: 'dep', width: 22 },
            ];
            ws5.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            ws5.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
            capex.forEach(c => ws5.addRow({
                name: c.name, value: c.baselineValue, life: c.depreciationYears,
                branch: c.branchMultiplier ? 'Yes' : 'No',
                dep: Math.round(c.baselineValue / (c.depreciationYears || 5))
            }));

            // ── Sheet 6: P&L Summary ────────────────────────────────
            if (projectionOutputs?.yearly?.pnl) {
                const ws6 = wb.addWorksheet('1. P&L');
                const years = projectionOutputs.yearly.pnl;
                ws6.columns = [
                    { header: 'Metric', key: 'metric', width: 30 },
                    ...years.map((_, i) => ({ header: `Year ${i + 1}`, key: `y${i + 1}`, width: 16 }))
                ];
                ws6.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
                ws6.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };

                const pnlMetrics = [
                    { label: 'Total Revenue (₹)', key: 'revenue' },
                    { label: 'Total OPEX (₹)', key: 'opex' },
                    { label: 'EBITDA (₹)', key: 'ebitda' },
                    { label: 'Depreciation (₹)', key: 'depreciation' },
                    { label: 'EBIT (₹)', key: 'ebit' },
                    { label: 'Tax (₹)', key: 'tax' },
                    { label: 'Net Profit (₹)', key: 'netProfit' },
                ];
                pnlMetrics.forEach(m => {
                    const row = { metric: m.label };
                    years.forEach((y, i) => { row[`y${i + 1}`] = Math.round(y[m.key] || 0); });
                    ws6.addRow(row);
                });
            }

            // Generate buffer and trigger download
            const buffer = await wb.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${businessInfo.tradeName || businessInfo.legalName || 'Docty'}_Financial_Model.xlsx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (err) {
            console.error('Excel export error:', err);
            alert('Could not generate Excel. Please try again.');
        } finally {
            setLoading(false);
        }
    };

    return (
        <button
            onClick={handleDownload}
            disabled={loading}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-green-600 hover:bg-green-700 text-white rounded-lg transition-colors disabled:opacity-60 shadow-sm"
            title="Download Excel model"
        >
            {loading
                ? <Activity size={13} className="animate-spin" />
                : <Download size={13} />}
            {loading ? 'Generating...' : 'Download Excel'}
        </button>
    );
}

// ─── Main ChatPanel ─────────────────────────────────────────────────────────
export default function ChatPanel() {
    const financial = useFinancial();
    const {
        addProductViaChat, addOpexViaChat, setBranchesViaChat,
        addCapexViaChat, setAssumptions, setBusinessInfo,
        setFunding, markComplete, collectionProgress, navigateToTab,
        setBranchCount
    } = financial;

    const chatEndRef = useRef(null);
    const inputRef = useRef(null);
    const [localInput, setLocalInput] = useState('');
    const [quickReplies, setQuickReplies] = useState([]);
    const [dataCards, setDataCards] = useState({});  // { msgId: [card,...] }
    const [typingStatus, setTypingStatus] = useState('');

    // Track which message IDs we've already processed
    const processedIds = useRef(new Set());

    const addDataCard = useCallback((msgId, card) => {
        setDataCards(prev => ({
            ...prev,
            [msgId]: [...(prev[msgId] || []), card]
        }));
    }, []);

    // ── Robust DATA tag extractor ─────────────────────────────────────────
    // Handles multi-field JSON objects properly
    const extractDataTags = useCallback((text) => {
        const results = [];
        // Find [DATA: {…}] — grab everything between the first { and matching }
        const re = /\[DATA:\s*(\{[\s\S]*?\})\]/g;
        let m;
        while ((m = re.exec(text)) !== null) {
            try {
                const obj = JSON.parse(m[1]);
                results.push(obj);
            } catch {
                // Try repairing simple truncation issues
                try {
                    const repaired = m[1].replace(/,\s*$/, '') + '}';
                    results.push(JSON.parse(repaired));
                } catch { /* skip malformed */ }
            }
        }
        return results;
    }, []);

    // ── Dispatch a single parsed DATA object ──────────────────────────────
    const dispatch = useCallback((data, msgId) => {
        const allowed = ['setBusinessInfo', 'setBranches', 'setBranchCount', 'addRevenueStream', 'addOpex', 'addCapex', 'setAssumptions', 'setFunding', 'markComplete', 'navigateTab', 'showChips'];
        if (!allowed.includes(data.type)) return;

        switch (data.type) {
            case 'setBusinessInfo': {
                const patch = {};
                ['legalName', 'tradeName', 'address', 'email', 'phone', 'description', 'companyType', 'startDate']
                    .forEach(f => { if (data[f] != null) patch[f] = String(data[f]).slice(0, 500); });
                ['equityShares', 'faceValue', 'paidUpCapital']
                    .forEach(f => { if (data[f] != null) patch[f] = Number(data[f]) || 0; });
                if (Array.isArray(data.promoters)) patch.promoters = data.promoters.map(p => String(p).slice(0, 100));
                setBusinessInfo(patch);
                addDataCard(msgId, { type: 'setBusinessInfo', ...patch });
                break;
            }
            case 'setBranches': {
                if (Array.isArray(data.branches)) {
                    const safe = data.branches.map((b, i) => ({
                        id: Number(b.id) || i + 1,
                        name: String(b.name || `Branch ${i + 1}`).slice(0, 100),
                        startMonth: Math.max(1, Number(b.startMonth) || 1),
                        status: ['Active', 'Planned', 'Inactive'].includes(b.status) ? b.status : 'Active'
                    }));
                    setBranchesViaChat(safe);
                    addDataCard(msgId, { type: 'setBranches', branches: safe });
                }
                break;
            }
            case 'addRevenueStream': {
                const sn = String(data.streamName || 'Revenue').slice(0, 100);
                const sub = String(data.subName || 'General').slice(0, 100);
                const pn = String(data.productName || 'Service').slice(0, 100);
                const u = Math.max(0, Math.min(10000, Number(data.units) || 0));
                const p = Math.max(0, Math.min(10000000, Number(data.price) || 0));
                const g = Math.max(0, Math.min(2, Number(data.growthY1) || 0.1));
                const cellRef = (data.cellRef && typeof data.cellRef === 'object') ? data.cellRef : null;
                const growthRates = (data.growthRates && typeof data.growthRates === 'object') ? data.growthRates : null;
                addProductViaChat(sn, sub, pn, u, p, g, cellRef, growthRates);
                addDataCard(msgId, { type: 'addRevenueStream', productName: pn, units: u, price: p });
                break;
            }
            case 'addOpex': {
                const cat = String(data.category || 'Expense').slice(0, 100);
                const sub = String(data.subCategory || 'General').slice(0, 100);
                const price = Math.max(0, Math.min(100000000, Number(data.price) || 0));
                const units = Math.max(0, Number(data.units) || 1);
                const cellRef = (data.cellRef && typeof data.cellRef === 'object') ? data.cellRef : null;
                const growthRates = (data.growthRates && typeof data.growthRates === 'object') ? data.growthRates : null;
                addOpexViaChat(cat, sub, units, price, cellRef, growthRates);
                addDataCard(msgId, { type: 'addOpex', category: cat, subCategory: sub, price });
                break;
            }
            case 'addCapex': {
                const name = String(data.name || 'Asset').slice(0, 150);
                const cost = Math.max(0, Math.min(1e9, Number(data.cost) || 0));
                const life = Math.max(1, Math.min(50, Number(data.usefulLife) || 5));
                addCapexViaChat(name, cost, life);
                addDataCard(msgId, { type: 'addCapex', name, cost, usefulLife: life });
                break;
            }
            case 'setAssumptions': {
                const patch = {};
                if (data.initialInvestment != null) patch.initialInvestment = Math.max(0, Number(data.initialInvestment) || 0);
                if (data.taxRate != null) patch.taxRate = Math.max(0, Math.min(1, Number(data.taxRate) || 0.25));
                if (data.inflationRate != null) patch.inflationRate = Math.max(0, Math.min(1, Number(data.inflationRate) || 0.05));
                setAssumptions(prev => ({ ...prev, ...patch }));
                addDataCard(msgId, { type: 'setAssumptions', ...patch });
                break;
            }
            case 'setFunding': {
                const patch = {
                    loanAmount: Math.max(0, Number(data.loanAmount) || 0),
                    interestRate: Math.max(0, Math.min(100, Number(data.interestRate) || 0)),
                    loanTenureMonths: Math.max(0, Number(data.loanTenureMonths) || 0),
                    moratoriumMonths: Math.max(0, Number(data.moratoriumMonths) || 0),
                    equityFromPromoters: Math.max(0, Number(data.equityFromPromoters) || 0),
                };
                setFunding(patch);
                addDataCard(msgId, { type: 'setFunding', ...patch });
                break;
            }
            case 'markComplete': {
                const valid = ['basics', 'branches', 'revenue', 'opex', 'capex', 'funding'];
                if (valid.includes(data.block)) {
                    markComplete(data.block);
                    addDataCard(msgId, { type: 'markComplete', block: data.block });
                }
                break;
            }
            case 'navigateTab': {
                const validTabs = ['2. Basics', 'Branch', 'A.I Revenue Streams', 'A.II OPEX', 'A.III CAPEX', 'B.I Sales - P1', '1. P&L', '5. Balance Sheet'];
                if (validTabs.includes(data.tab)) navigateToTab(data.tab);
                break;
            }
            case 'setBranchCount': {
                const count = Math.max(1, Math.min(100, Number(data.count) || 1));
                setBranchCount(count);
                addDataCard(msgId, { type: 'setBranchCount', count });
                break;
            }
            case 'showChips': {
                if (Array.isArray(data.options)) {
                    setQuickReplies(data.options.slice(0, 6).map(o => String(o).slice(0, 60)));
                }
                break;
            }
        }
    }, [addProductViaChat, addOpexViaChat, setBranchesViaChat, addCapexViaChat, setAssumptions, setBusinessInfo, setFunding, markComplete, navigateToTab, addDataCard, setBranchCount]);

    // ── KEY FIX: Watch all messages via useEffect, process new ones ────────
    const { messages, sendMessage, isLoading } = useChat({
        api: '/api/chat',
        initialMessages: [
            {
                id: 'init-0',
                role: 'assistant',
                content: "👋 Hello! I'm **Docty**, your AI financial model assistant.\n\nI'll ask you questions the way a CA would — and fill your financial model live as we chat.\n\nWe'll cover business basics, revenue streams, expenses, capital investments, branches, and funding.\n\n**Let's start: What is the full legal name of your business?**"
            }
        ],
    });

    // Process DATA tags from every new assistant message
    useEffect(() => {
        messages.forEach(msg => {
            if (msg.role !== 'assistant') return;
            if (processedIds.current.has(msg.id)) return;
            // Only process complete messages (not while streaming)
            if (isLoading && msg === messages[messages.length - 1]) return;

            const text = extractText(msg);
            if (!text) return;

            processedIds.current.add(msg.id);
            const tags = extractDataTags(text);
            tags.forEach(tag => dispatch(tag, msg.id));

            // Also extract CHIPS
            const chipsMatch = text.match(/\[CHIPS:\s*(\[[^\]]+\])\]/);
            if (chipsMatch) {
                try {
                    const opts = JSON.parse(chipsMatch[1]);
                    if (Array.isArray(opts)) setQuickReplies(opts.slice(0, 6).map(o => String(o).slice(0, 60)));
                } catch { /* ignore */ }
            }
        });
    }, [messages, isLoading, extractDataTags, dispatch]);

    // Typing status cycle
    useEffect(() => {
        if (!isLoading) { setTypingStatus(''); return; }
        const statuses = ['Docty is thinking...', 'Analysing your business...', 'Computing projections...', 'Preparing your model...'];
        let i = 0;
        setTypingStatus(statuses[0]);
        const t = setInterval(() => { i = (i + 1) % statuses.length; setTypingStatus(statuses[i]); }, 2000);
        return () => clearInterval(t);
    }, [isLoading]);

    // Auto-scroll
    useEffect(() => { chatEndRef.current?.scrollIntoView({ behavior: 'smooth' }); }, [messages, isLoading]);

    const handleSend = useCallback((text) => {
        const trimmed = text.trim();
        if (!trimmed || isLoading) return;
        setQuickReplies([]);
        setLocalInput('');
        sendMessage({ role: 'user', content: trimmed });
    }, [isLoading, sendMessage]);

    return (
        <div className="flex flex-col h-full bg-white/80 backdrop-blur-xl rounded-r-2xl border-r border-slate-200 shadow-sm relative overflow-hidden">

            {/* Header */}
            <div className="p-3 border-b border-slate-100 bg-white/90 flex items-center gap-3 shrink-0">
                <div className="w-9 h-9 rounded-full bg-blue-600 flex items-center justify-center text-white shadow-md shrink-0">
                    <Zap size={16} />
                </div>
                <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2">
                        <h2 className="font-bold text-slate-800 text-sm">Docty AI</h2>
                        <span className="text-[10px] font-bold px-1.5 py-0.5 bg-green-100 text-green-700 rounded">LIVE</span>
                    </div>
                    <p className="text-[11px] text-slate-500 font-medium">Agentic Financial Model Builder</p>
                </div>
                <DownloadButton financialData={financial} />
            </div>

            {/* Progress Bar */}
            <ProgressBar progress={collectionProgress} />

            {/* Messages */}
            <div className="flex-1 overflow-y-auto p-4 space-y-4">
                {messages.map((msg) => {
                    const raw = extractText(msg);
                    const text = msg.role === 'assistant' ? cleanForDisplay(raw) : raw;
                    const cards = dataCards[msg.id] || [];

                    return (
                        <motion.div key={msg.id} initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                            className={`flex gap-3 max-w-[92%] ${msg.role === 'user' ? 'ml-auto flex-row-reverse' : ''}`}>
                            <div className={`w-8 h-8 rounded-full flex items-center justify-center shrink-0 border ${msg.role === 'user' ? 'bg-slate-800 text-white border-slate-700' : 'bg-blue-600 text-white border-blue-700'}`}>
                                {msg.role === 'user' ? <User size={14} /> : <Bot size={14} />}
                            </div>

                            <div className={`flex flex-col flex-1 min-w-0 ${msg.role === 'user' ? 'items-end' : 'items-start'}`}>
                                {text ? (
                                    <div className={`p-3 rounded-2xl text-[13.5px] leading-relaxed ${msg.role === 'user' ? 'bg-slate-800 text-white rounded-tr-sm font-medium' : 'bg-slate-100 text-slate-800 border border-slate-200 rounded-tl-sm'}`}>
                                        {msg.role === 'assistant' ? (
                                            <ReactMarkdown components={{
                                                h1: ({ children }) => <h1 className="text-base font-bold mt-2 mb-1">{children}</h1>,
                                                h2: ({ children }) => <h2 className="text-sm font-bold mt-2 mb-1">{children}</h2>,
                                                h3: ({ children }) => <h3 className="text-sm font-semibold mt-1.5 mb-0.5">{children}</h3>,
                                                p: ({ children }) => <p className="mb-1.5 last:mb-0">{children}</p>,
                                                strong: ({ children }) => <strong className="font-semibold text-slate-900">{children}</strong>,
                                                ul: ({ children }) => <ul className="list-disc pl-4 mb-2 space-y-0.5">{children}</ul>,
                                                ol: ({ children }) => <ol className="list-decimal pl-4 mb-2 space-y-0.5">{children}</ol>,
                                                li: ({ children }) => <li className="text-[13px]">{children}</li>,
                                                hr: () => <hr className="my-2 border-slate-300" />,
                                                code: ({ inline, children }) => inline
                                                    ? <code className="bg-slate-200 px-1 py-0.5 rounded text-xs font-mono">{children}</code>
                                                    : <pre className="bg-slate-800 text-green-300 p-3 rounded-lg text-xs font-mono my-2 overflow-x-auto"><code>{children}</code></pre>,
                                            }}>
                                                {text}
                                            </ReactMarkdown>
                                        ) : text}
                                    </div>
                                ) : null}

                                <AnimatePresence>
                                    {cards.map((card, i) => <DataCard key={i} card={card} />)}
                                </AnimatePresence>
                            </div>
                        </motion.div>
                    );
                })}

                {isLoading && (
                    <div className="flex gap-3 max-w-[85%]">
                        <div className="w-8 h-8 rounded-full flex items-center justify-center bg-blue-600 text-white border border-blue-700 shrink-0">
                            <Bot size={14} />
                        </div>
                        <div className="p-3 rounded-2xl bg-slate-100 border border-slate-200 rounded-tl-sm flex items-center gap-2">
                            <div className="flex gap-1">
                                {[0, 0.2, 0.4].map((d, i) => (
                                    <motion.div key={i} animate={{ opacity: [0.3, 1, 0.3] }} transition={{ repeat: Infinity, duration: 1, delay: d }}
                                        className="w-2 h-2 rounded-full bg-blue-400" />
                                ))}
                            </div>
                            <span className="text-xs text-slate-500 font-medium">{typingStatus}</span>
                        </div>
                    </div>
                )}
                <div ref={chatEndRef} />
            </div>

            {/* Quick Replies */}
            <QuickReplyChips chips={quickReplies} onSelect={handleSend} />

            {/* Input */}
            <div className="p-3 bg-white/90 border-t border-slate-100 shrink-0">
                <form onSubmit={(e) => { e.preventDefault(); handleSend(localInput); }} className="flex items-center gap-2">
                    <input
                        ref={inputRef}
                        type="text"
                        value={localInput}
                        onChange={e => setLocalInput(e.target.value)}
                        placeholder="Talk to Docty to build your financial model..."
                        disabled={isLoading}
                        className="flex-1 bg-white border border-slate-200 rounded-full px-4 py-2.5 text-[13.5px] focus:outline-none focus:ring-2 focus:ring-blue-500/40 focus:border-blue-500/60 transition-all text-slate-700 placeholder:text-slate-400 shadow-sm disabled:opacity-60"
                    />
                    <button type="submit" disabled={!localInput.trim() || isLoading}
                        className="w-10 h-10 rounded-full bg-blue-600 hover:bg-blue-700 text-white flex items-center justify-center disabled:opacity-40 transition-colors shadow-md shrink-0">
                        <Send size={16} />
                    </button>
                </form>
                <p className="mt-1.5 text-center text-[10px] text-slate-400">Powered by Vercel AI Gateway · Agentic CA-grade model builder</p>
            </div>
        </div>
    );
}
