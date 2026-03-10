/**
 * POST /api/export-model
 * ======================
 * Accepts the full UI model state as JSON and generates a fresh .xlsx
 * that exactly mirrors what the UI shows — same calculations, same layout.
 *
 * No template required. Generates from scratch using ExcelJS.
 */
import ExcelJS from 'exceljs';

export const maxDuration = 60;

// ─── Shared calculation functions (mirrors DoctyFinancialModelFull.jsx) ────────
const YEARS = ['2025-26', '2026-27', '2027-28', '2028-29', '2029-30'];

function calcItemRevYearly(item) {
    if (!item) return YEARS.map(() => 0);
    return YEARS.map((_, yi) => {
        const qty = Number(item.qty) || (Number(item.qtyDay) * 30) || 0;
        const base = qty * (Number(item.price) || 0) * 12;
        if (!base) return 0;
        let v = base;
        for (let i = 0; i < yi; i++) v *= (1 + (Number(item[`gY${i + 1}`]) || 0));
        return v;
    });
}

function calcItemOpexYearly(item) {
    if (!item) return YEARS.map(() => 0);
    return YEARS.map((_, yi) => {
        const base = (Number(item.qty) || 1) * (Number(item.cost) || 0) * 12;
        if (!base) return 0;
        let v = base;
        for (let i = 0; i < yi; i++) v *= (1 + (Number(item[`gY${i + 1}`]) || 0));
        return v;
    });
}

function calcRevYearly(groups = []) {
    return (Array.isArray(groups) ? groups : []).map(g => {
        const items = (Array.isArray(g.items) ? g.items : []).map(it => ({
            ...it,
            yearlyTotals: calcItemRevYearly(it),
        }));
        return {
            ...g,
            items,
            yearlyTotals: YEARS.map((_, yi) => items.reduce((s, it) => s + (it.yearlyTotals[yi] || 0), 0)),
        };
    });
}

function calcOpexYearly(groups = []) {
    return (Array.isArray(groups) ? groups : []).map(g => {
        const items = (Array.isArray(g.items) ? g.items : []).map(it => ({
            ...it,
            yearlyTotals: calcItemOpexYearly(it),
        }));
        return {
            ...g,
            items,
            yearlyTotals: YEARS.map((_, yi) => items.reduce((s, it) => s + (it.yearlyTotals[yi] || 0), 0)),
        };
    });
}

function fmtCr(v) {
    if (!v && v !== 0) return '—';
    const abs = Math.abs(v);
    const neg = v < 0;
    let s;
    if (abs >= 10000000) s = `₹${(abs / 10000000).toFixed(2)} Cr`;
    else if (abs >= 100000) s = `₹${(abs / 100000).toFixed(2)} L`;
    else s = `₹${Math.round(abs).toLocaleString('en-IN')}`;
    return neg ? `(${s})` : s;
}

// ─── Styling helpers ───────────────────────────────────────────────────────────
const COLORS = {
    headerBg: '0C1426',
    sectionBg: '0C1830',
    totalBg: '091422',
    gold: 'C4972A',
    goldL: 'E8C96B',
    teal: '2A9E9E',
    green: '2EA870',
    red: 'C84040',
    text0: 'E0EAF8',
    text1: '8AAAC8',
    text2: '4A6888',
    white: 'FFFFFF',
    lightGray: 'F4F6FA',
    midGray: 'D0D8E8',
};

function headerStyle(hex = COLORS.headerBg, textHex = COLORS.gold) {
    return {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${hex}` } },
        font: { bold: true, color: { argb: `FF${textHex}` }, size: 10, name: 'Calibri' },
        alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
        border: {
            bottom: { style: 'thin', color: { argb: `FF${COLORS.goldL}` } },
        },
    };
}

function bodyStyle(hex = COLORS.lightGray, textHex = COLORS.text0, right = false) {
    return {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${hex}` } },
        font: { color: { argb: `FF${textHex}` }, size: 10, name: 'Calibri' },
        alignment: { vertical: 'middle', horizontal: right ? 'right' : 'left' },
    };
}

function totalStyle(textHex = COLORS.goldL) {
    return {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.totalBg}` } },
        font: { bold: true, color: { argb: `FF${textHex}` }, size: 10, name: 'Calibri' },
        alignment: { vertical: 'middle', horizontal: 'right' },
        border: { top: { style: 'thin', color: { argb: `FF${COLORS.gold}` } } },
    };
}

function numFmt(v) {
    if (v == null || v === '') return '';
    const n = Number(v);
    if (!isFinite(n)) return String(v);
    return n;
}

// ─── Sheet builders ────────────────────────────────────────────────────────────

function buildBasicsSheet(wb, basics = {}, loan1 = {}, loan2 = {}, totalProjectCost = {}) {
    const ws = wb.addWorksheet('1. Basics');
    ws.columns = [
        { key: 'sno', width: 8 },
        { key: 'field', width: 42 },
        { key: 'value', width: 50 },
    ];

    // Title row
    ws.mergeCells('A1:C1');
    const titleCell = ws.getCell('A1');
    titleCell.value = '1. Basic Information — OnEasy Financial Model';
    titleCell.style = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } },
        font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' },
        alignment: { vertical: 'middle', horizontal: 'left' },
    };
    ws.getRow(1).height = 32;

    const rows = [
        ['1', 'Legal Name of the Business', basics.legalName || ''],
        ['2', 'Trade Name of the Business', basics.tradeName || ''],
        ['3', 'Registered Office Address', basics.address || ''],
        ['4', 'Official Email Id', basics.email || ''],
        ['5', 'Official Contact Number', basics.contact || ''],
        ['6.a', 'Total number of Promoters', basics.promoters || ''],
        ['7', 'Tentative Start Date of Phase 1', basics.startDateP1 || ''],
        ['8', 'Tentative Start Date of Phase 2', basics.startDateP2 || ''],
        ['9', 'Company Description', basics.description || ''],
        ['10', 'Burning Desire of the Company', basics.burningDesire || ''],
        ['—', '', ''],
        ['L1', 'Loan 1 — Amount (₹)', numFmt(loan1?.amount)],
        ['L1', 'Loan 1 — Interest Rate (%)', numFmt(loan1?.rate)],
        ['L1', 'Loan 1 — Duration (months)', numFmt(loan1?.duration)],
        ['L2', 'Loan 2 — Amount (₹)', numFmt(loan2?.amount)],
        ['L2', 'Loan 2 — Interest Rate (%)', numFmt(loan2?.rate)],
        ['L2', 'Loan 2 — Duration (months)', numFmt(loan2?.duration)],
        ['—', '', ''],
        ['PC', 'Total Project Cost (₹)', numFmt(totalProjectCost?.total)],
        ['PC', 'Promoter Contribution (₹)', numFmt(totalProjectCost?.promoterContrib)],
        ['PC', 'Term Loan (₹)', numFmt(totalProjectCost?.termLoan)],
        ['PC', 'Working Capital Loan (₹)', numFmt(totalProjectCost?.wcLoan)],
    ];

    // Header row
    const hRow = ws.addRow(['#', 'Field', 'Value']);
    hRow.eachCell(c => Object.assign(c, { style: headerStyle() }));
    hRow.height = 22;

    rows.forEach(([sno, field, value], i) => {
        const r = ws.addRow([sno, field, value]);
        r.height = 20;
        const bg = i % 2 === 0 ? 'EEF1F8' : COLORS.white;
        r.getCell(1).style = bodyStyle(bg, COLORS.text2);
        r.getCell(2).style = bodyStyle(bg, COLORS.text0);
        r.getCell(3).style = {
            ...bodyStyle(bg, COLORS.teal),
            font: { bold: !!value, color: { argb: `FF${COLORS.teal}` }, size: 10, name: 'Calibri' },
        };
    });
}

function buildRevenueSheet(wb, revP1 = []) {
    const ws = wb.addWorksheet('A.I Revenue Streams - P1');
    const computed = calcRevYearly(revP1);

    // Title
    ws.mergeCells('A1:N1');
    const t = ws.getCell('A1');
    t.value = 'A.I Revenue Streams — Phase 1 (Monthly × 12 → Annual)';
    t.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' } };
    ws.getRow(1).height = 30;

    ws.columns = [
        { key: 'id', width: 7 },
        { key: 'sub', width: 34 },
        { key: 'qty', width: 14 },
        { key: 'price', width: 14 },
        { key: 'gY1', width: 10 },
        { key: 'gY2', width: 10 },
        { key: 'gY3', width: 10 },
        { key: 'gY4', width: 10 },
        { key: 'gY5', width: 10 },
        ...YEARS.map(() => ({ width: 16 })),
    ];

    // Header row
    const hdr = ws.addRow(['ID', 'Item / Sub-Service', 'Qty/Month', 'Unit Price (₹)', 'Growth Y1', 'Growth Y2', 'Growth Y3', 'Growth Y4', 'Growth Y5', ...YEARS.map(y => `Annual Rev ${y}`)]);
    hdr.height = 28;
    hdr.eachCell(c => Object.assign(c, { style: headerStyle() }));

    computed.forEach((g) => {
        // Group header row
        const gr = ws.addRow([g.id, g.header || '—', '', '', '', '', '', '', '', ...g.yearlyTotals.map(v => fmtCr(v))]);
        gr.height = 22;
        gr.getCell(1).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.gold}` }, size: 10, name: 'Calibri' } };
        gr.getCell(2).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.text0}` }, size: 10, name: 'Calibri' } };
        for (let ci = 3; ci <= 9; ci++) gr.getCell(ci).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } } };
        for (let yi = 0; yi < YEARS.length; yi++) {
            const c = gr.getCell(10 + yi);
            c.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.teal}` }, size: 10, name: 'Calibri' }, alignment: { horizontal: 'right' } };
        }

        // Item rows
        (g.items || []).filter(it => it.sub || Number(it.qty) > 0).forEach((it, ii) => {
            const bg = ii % 2 === 0 ? 'F2F5FA' : COLORS.white;
            const r = ws.addRow([
                it.id, it.sub || '—',
                Number(it.qty) || 0,
                Number(it.price) || 0,
                `${((Number(it.gY1) || 0) * 100).toFixed(1)}%`,
                `${((Number(it.gY2) || 0) * 100).toFixed(1)}%`,
                `${((Number(it.gY3) || 0) * 100).toFixed(1)}%`,
                `${((Number(it.gY4) || 0) * 100).toFixed(1)}%`,
                `${((Number(it.gY5) || 0) * 100).toFixed(1)}%`,
                ...it.yearlyTotals.map(v => fmtCr(v)),
            ]);
            r.height = 18;
            r.getCell(1).style = bodyStyle(bg, COLORS.text2);
            r.getCell(2).style = bodyStyle(bg, COLORS.text1);
            r.getCell(3).style = { ...bodyStyle(bg, COLORS.text0, true) };
            r.getCell(4).style = { ...bodyStyle(bg, COLORS.text0, true) };
            for (let ci = 5; ci <= 9; ci++) r.getCell(ci).style = bodyStyle(bg, COLORS.text2, true);
            for (let yi = 0; yi < YEARS.length; yi++) {
                r.getCell(10 + yi).style = { ...bodyStyle(bg, COLORS.text0, true), font: { color: { argb: `FF${COLORS.text0}` }, size: 10, name: 'Calibri' } };
            }
        });

        // Grand Total row
        const gt = ws.addRow(['', 'Grand Total', '', '', '', '', '', '', '', ...g.yearlyTotals.map(v => fmtCr(v))]);
        gt.height = 20;
        for (let ci = 1; ci <= 9; ci++) gt.getCell(ci).style = totalStyle(COLORS.goldL);
        for (let yi = 0; yi < YEARS.length; yi++) {
            gt.getCell(10 + yi).style = { ...totalStyle(COLORS.gold), alignment: { horizontal: 'right' } };
        }
        ws.addRow([]); // spacer
    });

    // Yearly Grand Total
    const totals = YEARS.map((_, yi) => computed.reduce((s, g) => s + g.yearlyTotals[yi], 0));
    const ygt = ws.addRow(['', 'YEARLY GRAND TOTAL', '', '', '', '', '', '', '', ...totals.map(v => fmtCr(v))]);
    ygt.height = 24;
    for (let ci = 1; ci <= 9; ci++) {
        ygt.getCell(ci).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF07101E` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 11, name: 'Calibri' }, border: { top: { style: 'medium', color: { argb: `FF${COLORS.gold}` } }, bottom: { style: 'medium', color: { argb: `FF${COLORS.gold}` } } } };
    }
    for (let yi = 0; yi < YEARS.length; yi++) {
        ygt.getCell(10 + yi).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF07101E` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 11, name: 'Calibri' }, alignment: { horizontal: 'right' }, border: { top: { style: 'medium', color: { argb: `FF${COLORS.gold}` } }, bottom: { style: 'medium', color: { argb: `FF${COLORS.gold}` } } } };
    }
}

function buildOpexSheet(wb, opexP1 = []) {
    const ws = wb.addWorksheet('A.II OPEX');
    const computed = calcOpexYearly(opexP1);

    ws.mergeCells('A1:N1');
    const t = ws.getCell('A1');
    t.value = 'A.II OPEX — Operating Expenses (Monthly × 12 → Annual)';
    t.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' } };
    ws.getRow(1).height = 30;

    ws.columns = [
        { key: 'id', width: 7 }, { key: 'sub', width: 34 },
        { key: 'qty', width: 12 }, { key: 'cost', width: 14 },
        { key: 'gY1', width: 10 }, { key: 'gY2', width: 10 }, { key: 'gY3', width: 10 }, { key: 'gY4', width: 10 }, { key: 'gY5', width: 10 },
        ...YEARS.map(() => ({ width: 16 })),
    ];

    const hdr = ws.addRow(['ID', 'Item / Category', 'Headcount / Units', 'Monthly Cost ₹', 'Growth Y1', 'Growth Y2', 'Growth Y3', 'Growth Y4', 'Growth Y5', ...YEARS.map(y => `Annual OPEX ${y}`)]);
    hdr.height = 28;
    hdr.eachCell(c => Object.assign(c, { style: headerStyle(COLORS.headerBg, COLORS.red) }));

    computed.forEach((g) => {
        const gr = ws.addRow([g.id, g.header || '—', '', '', '', '', '', '', '', ...g.yearlyTotals.map(v => fmtCr(v))]);
        gr.height = 22;
        gr.getCell(1).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.gold}` }, size: 10, name: 'Calibri' } };
        gr.getCell(2).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.text0}` }, size: 10, name: 'Calibri' } };
        for (let ci = 3; ci <= 9; ci++) gr.getCell(ci).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } } };
        for (let yi = 0; yi < YEARS.length; yi++) {
            gr.getCell(10 + yi).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.red}` }, size: 10, name: 'Calibri' }, alignment: { horizontal: 'right' } };
        }

        (g.items || []).filter(it => it.sub || Number(it.cost) > 0).forEach((it, ii) => {
            const bg = ii % 2 === 0 ? 'F2F5FA' : COLORS.white;
            const totalMonthly = (Number(it.qty) || 1) * (Number(it.cost) || 0);
            const r = ws.addRow([
                it.id, it.sub || '—',
                Number(it.qty) || 1,
                totalMonthly,
                `${((Number(it.gY1) || 0) * 100).toFixed(1)}%`,
                `${((Number(it.gY2) || 0) * 100).toFixed(1)}%`,
                `${((Number(it.gY3) || 0) * 100).toFixed(1)}%`,
                `${((Number(it.gY4) || 0) * 100).toFixed(1)}%`,
                `${((Number(it.gY5) || 0) * 100).toFixed(1)}%`,
                ...it.yearlyTotals.map(v => fmtCr(v)),
            ]);
            r.height = 18;
            r.getCell(1).style = bodyStyle(bg, COLORS.text2);
            r.getCell(2).style = bodyStyle(bg, COLORS.text1);
            r.getCell(3).style = { ...bodyStyle(bg, COLORS.text0, true) };
            r.getCell(4).style = { ...bodyStyle(bg, COLORS.text0, true) };
            for (let ci = 5; ci <= 9; ci++) r.getCell(ci).style = bodyStyle(bg, COLORS.text2, true);
            for (let yi = 0; yi < YEARS.length; yi++) {
                r.getCell(10 + yi).style = { ...bodyStyle(bg, COLORS.red, true) };
            }
        });

        const gt = ws.addRow(['', 'Grand Total', '', '', '', '', '', '', '', ...g.yearlyTotals.map(v => fmtCr(v))]);
        gt.height = 20;
        for (let ci = 1; ci <= 9; ci++) gt.getCell(ci).style = totalStyle(COLORS.goldL);
        for (let yi = 0; yi < YEARS.length; yi++) gt.getCell(10 + yi).style = { ...totalStyle(COLORS.red), alignment: { horizontal: 'right' } };
        ws.addRow([]);
    });

    const totals = YEARS.map((_, yi) => computed.reduce((s, g) => s + g.yearlyTotals[yi], 0));
    const ygt = ws.addRow(['', 'TOTAL OPEX', '', '', '', '', '', '', '', ...totals.map(v => fmtCr(v))]);
    ygt.height = 24;
    for (let ci = 1; ci <= 14; ci++) {
        ygt.getCell(ci).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF07101E` } }, font: { bold: true, color: { argb: `FF${COLORS.red}` }, size: 11, name: 'Calibri' }, alignment: { horizontal: ci > 9 ? 'right' : 'left' }, border: { top: { style: 'medium', color: { argb: `FF${COLORS.red}` } }, bottom: { style: 'medium', color: { argb: `FF${COLORS.red}` } } } };
    }
}

function buildPLSheet(wb, revP1 = [], opexP1 = [], loan1 = {}, loan2 = {}) {
    const ws = wb.addWorksheet('4. P&L');

    ws.mergeCells('A1:G1');
    const t = ws.getCell('A1');
    t.value = '4. Profit & Loss Statement — 5-Year Projection';
    t.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' } };
    ws.getRow(1).height = 30;

    ws.columns = [{ width: 6 }, { width: 36 }, ...YEARS.map(() => ({ width: 18 }))];

    const hdr = ws.addRow(['#', 'Metric', ...YEARS]);
    hdr.height = 26;
    hdr.eachCell(c => Object.assign(c, { style: headerStyle() }));

    const revByYear = YEARS.map((_, yi) => calcRevYearly(revP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
    const opexByYear = YEARS.map((_, yi) => calcOpexYearly(opexP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
    const ebitda = YEARS.map((_, yi) => revByYear[yi] - opexByYear[yi]);
    const deprn = YEARS.map(() => 75000 * 12);
    const ebit = YEARS.map((_, yi) => ebitda[yi] - deprn[yi]);
    const interest = YEARS.map((_, yi) => [0, (Number(loan1?.amount) || 0) * (Number(loan1?.rate) || 0) / 100, (Number(loan2?.amount) || 0) * (Number(loan2?.rate) || 0) / 100, 0, 0][yi] || 0);
    const pbt = YEARS.map((_, yi) => ebit[yi] - interest[yi]);
    const tax = YEARS.map((_, yi) => Math.max(0, pbt[yi] * 0.25));
    const pat = YEARS.map((_, yi) => pbt[yi] - tax[yi]);
    const ebitdaMgn = YEARS.map((_, yi) => revByYear[yi] > 0 ? ebitda[yi] / revByYear[yi] : 0);

    const plRows = [
        { sno: '1', label: 'Total Revenue', vals: revByYear, isTotal: true, color: COLORS.teal },
        { sno: '', label: 'Total OPEX', vals: opexByYear, isNeg: true, color: COLORS.red },
        { sno: '2', label: 'EBITDA', vals: ebitda, isTotal: true, color: COLORS.gold },
        { sno: '', label: 'EBITDA Margin', vals: ebitdaMgn, isPct: true, color: COLORS.goldL },
        { sno: '', label: 'Depreciation', vals: deprn, color: COLORS.text2 },
        { sno: '3', label: 'EBIT', vals: ebit, isTotal: true, color: COLORS.gold },
        { sno: '', label: 'Interest', vals: interest, color: COLORS.text2 },
        { sno: '4', label: 'PBT', vals: pbt, isTotal: true, color: COLORS.gold },
        { sno: '', label: 'Tax (25%)', vals: tax, color: COLORS.text2 },
        { sno: '5', label: 'Net Profit (PAT)', vals: pat, isTotal: true, color: COLORS.green },
    ];

    plRows.forEach((row, i) => {
        const bg = row.isTotal ? COLORS.totalBg : (i % 2 === 0 ? 'F2F5FA' : COLORS.white);
        const fgColor = { argb: `FF${bg}` };
        const fontColor = { argb: `FF${row.color || COLORS.text0}` };
        const r = ws.addRow([
            row.sno,
            row.label,
            ...row.vals.map(v => row.isPct ? `${(v * 100).toFixed(1)}%` : fmtCr(v)),
        ]);
        r.height = row.isTotal ? 22 : 18;
        let ci = 0;
        r.eachCell(c => {
            ci++;
            c.style = {
                fill: { type: 'pattern', pattern: 'solid', fgColor },
                font: { bold: row.isTotal, color: fontColor, size: 10, name: 'Calibri' },
                // First 2 columns left-aligned, rest right-aligned
                alignment: { horizontal: ci <= 2 ? 'left' : 'right', vertical: 'middle' },
            };
            if (row.isTotal) c.style.border = { top: { style: 'thin', color: { argb: `FF${COLORS.gold}` } } };
        });
    });
}

function buildScenariosSheet(wb, revP1 = [], opexP1 = []) {
    const ws = wb.addWorksheet('Scenarios');

    ws.mergeCells('A1:H1');
    const t = ws.getCell('A1');
    t.value = 'Scenario Analysis — 5-Year Projections (Pessimistic / Baseline / Optimistic)';
    t.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' } };
    ws.getRow(1).height = 30;

    ws.columns = [{ width: 18 }, { width: 18 }, ...YEARS.map(() => ({ width: 18 })), { width: 14 }];

    const SCENARIOS = [
        { label: '🔻 Pessimistic', revMult: 0.5, opexMult: 1.2, color: COLORS.red },
        { label: '📊 Baseline', revMult: 1.0, opexMult: 1.0, color: COLORS.gold },
        { label: '🚀 Optimistic', revMult: 1.6, opexMult: 0.85, color: COLORS.green },
    ];

    const calcScenario = (rMult, oMult) => {
        const adjRev = (revP1 || []).map(g => ({
            ...g, items: (g.items || []).map(it => ({
                ...it,
                gY1: (Number(it.gY1) || 0) * rMult,
                gY2: (Number(it.gY2) || 0) * rMult,
                gY3: (Number(it.gY3) || 0) * rMult,
                gY4: (Number(it.gY4) || 0) * rMult,
                gY5: (Number(it.gY5) || 0) * rMult,
            }))
        }));
        const adjOpex = (opexP1 || []).map(g => ({
            ...g, items: (g.items || []).map(it => ({
                ...it,
                gY1: (Number(it.gY1) || 0) * oMult,
                gY2: (Number(it.gY2) || 0) * oMult,
                gY3: (Number(it.gY3) || 0) * oMult,
                gY4: (Number(it.gY4) || 0) * oMult,
                gY5: (Number(it.gY5) || 0) * oMult,
            }))
        }));
        const rev = YEARS.map((_, yi) => calcRevYearly(adjRev).reduce((s, g) => s + g.yearlyTotals[yi], 0));
        const opex = YEARS.map((_, yi) => calcOpexYearly(adjOpex).reduce((s, g) => s + g.yearlyTotals[yi], 0));
        const ebitda = YEARS.map((_, yi) => rev[yi] - opex[yi]);
        const deprn = YEARS.map(() => 75000 * 12);
        const pat = YEARS.map((_, yi) => ebitda[yi] - deprn[yi] - Math.max(0, (ebitda[yi] - deprn[yi]) * 0.25));
        return { rev, opex, ebitda, pat };
    };

    const hdr = ws.addRow(['Scenario', 'Metric', ...YEARS, '5Y Total']);
    hdr.height = 26;
    hdr.eachCell(c => Object.assign(c, { style: headerStyle() }));

    SCENARIOS.forEach(scenario => {
        const { rev, opex, ebitda, pat } = calcScenario(scenario.revMult, scenario.opexMult);
        const fgColor = { argb: `FF${COLORS.totalBg}` };
        const fontColor = { argb: `FF${scenario.color}` };

        [
            ['Revenue', rev],
            ['OPEX', opex.map(v => -v)],
            ['EBITDA', ebitda],
            ['Net Profit', pat],
        ].forEach(([metric, vals], mi) => {
            const r = ws.addRow([
                mi === 0 ? scenario.label : '',
                metric,
                ...vals.map(v => fmtCr(v)),
                fmtCr(vals.reduce((a, b) => a + b, 0)),
            ]);
            r.height = 20;
            r.getCell(1).style = { fill: { type: 'pattern', pattern: 'solid', fgColor }, font: { bold: true, color: fontColor, size: 10, name: 'Calibri' } };
            r.getCell(2).style = { fill: { type: 'pattern', pattern: 'solid', fgColor }, font: { color: { argb: `FF${COLORS.text1}` }, size: 10, name: 'Calibri' } };
            for (let ci = 3; ci <= 3 + YEARS.length; ci++) {
                const v = (ci - 3 < YEARS.length) ? vals[ci - 3] : vals.reduce((a, b) => a + b, 0);
                r.getCell(ci).style = {
                    fill: { type: 'pattern', pattern: 'solid', fgColor },
                    font: { bold: mi === 2, color: { argb: `FF${v >= 0 ? scenario.color : COLORS.red}` }, size: 10, name: 'Calibri' },
                    alignment: { horizontal: 'right' },
                };
            }
        });
        ws.addRow([]); // spacer
    });
}

// ─── Main Route ───────────────────────────────────────────────────────────────
export async function POST(request) {
    try {
        const state = await request.json();
        const { basics = {}, revP1 = [], opexP1 = [], loan1 = {}, loan2 = {}, totalProjectCost = {} } = state;

        const wb = new ExcelJS.Workbook();
        wb.creator = 'OnEasy Financial Model';
        wb.created = new Date();
        wb.modified = new Date();

        buildBasicsSheet(wb, basics, loan1, loan2, totalProjectCost);
        buildRevenueSheet(wb, revP1);
        buildOpexSheet(wb, opexP1);
        buildPLSheet(wb, revP1, opexP1, loan1, loan2);
        buildScenariosSheet(wb, revP1, opexP1);

        const buffer = await wb.xlsx.writeBuffer();

        const bizName = String(basics?.tradeName || basics?.legalName || 'OnEasy')
            .replace(/[^a-zA-Z0-9]/g, '_').slice(0, 30);

        return new Response(buffer, {
            status: 200,
            headers: {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': `attachment; filename="${bizName}_Financial_Model.xlsx"`,
                'Content-Length': buffer.byteLength.toString(),
            },
        });
    } catch (err) {
        console.error('[export-model]', err);
        return Response.json({ error: err.message }, { status: 500 });
    }
}
