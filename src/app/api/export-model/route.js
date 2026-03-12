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

const MONTHS_Y1 = ["Apr '26", "May '26", "Jun '26", "Jul '26", "Aug '26", "Sep '26", "Oct '26", "Nov '26", "Dec '26", "Jan '27", "Feb '27", "Mar '27"];

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

function fmtPct(v) {
    return v == null ? '—' : `${(v * 100).toFixed(1)}%`;
}

// ─── Styling helpers ───────────────────────────────────────────────────────────
const COLORS = {
    headerBg: '0C1426',     // Dark navy for headers
    sectionBg: '0C1830',    // Dark navy for section headers
    totalBg: '091422',      // Very dark for totals
    gold: 'C4972A',
    goldL: 'E8C96B',
    teal: '2A9E9E',
    green: '2EA870',
    red: 'C84040',
    // Body Text Colors (Must be dark for visibility on white/light-gray backgrounds)
    text0: '1A202C',        // Very dark gray/navy (for primary values/labels)
    text1: '2D3748',        // Dark gray (secondary labels)
    text2: '4A5568',        // Medium-dark gray (tertiary labels, empty states)
    white: 'FFFFFF',
    lightGray: 'F4F6FA',    // Background for alternate rows
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

function buildRevenueSheetP2(wb, revP2 = []) {
    const ws = wb.addWorksheet('A.I Revenue Streams - P2');
    const computed = calcRevYearly(revP2);

    ws.mergeCells('A1:N1');
    const t = ws.getCell('A1');
    t.value = 'A.I Revenue Streams — Phase 2 (Daily × 30 × 12 → Annual)';
    t.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' } };
    ws.getRow(1).height = 30;

    ws.columns = [
        { key: 'id', width: 7 }, { key: 'sub', width: 34 }, { key: 'qtyDay', width: 14 }, { key: 'price', width: 14 },
        { key: 'gY1', width: 10 }, { key: 'gY2', width: 10 }, { key: 'gY3', width: 10 }, { key: 'gY4', width: 10 }, { key: 'gY5', width: 10 },
        ...YEARS.map(() => ({ width: 16 })),
    ];

    const hdr = ws.addRow(['ID', 'Item / Sub-Service', 'Qty/Day', 'Unit Price (₹)', 'Growth Y1', 'Growth Y2', 'Growth Y3', 'Growth Y4', 'Growth Y5', ...YEARS.map(y => `Annual Rev ${y}`)]);
    hdr.height = 28;
    hdr.eachCell(c => Object.assign(c, { style: headerStyle() }));

    computed.forEach((g) => {
        const gr = ws.addRow([g.id, g.header || '—', '', '', '', '', '', '', '', ...g.yearlyTotals.map(v => fmtCr(v))]);
        gr.height = 22;
        gr.getCell(1).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.gold}` }, size: 10, name: 'Calibri' } };
        gr.getCell(2).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.text0}` }, size: 10, name: 'Calibri' } };
        for (let ci = 3; ci <= 9; ci++) gr.getCell(ci).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } } };
        for (let yi = 0; yi < YEARS.length; yi++) gr.getCell(10 + yi).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.teal}` }, size: 10, name: 'Calibri' }, alignment: { horizontal: 'right' } };

        (g.items || []).filter(it => it.sub || Number(it.qtyDay) > 0).forEach((it, ii) => {
            const bg = ii % 2 === 0 ? 'F2F5FA' : COLORS.white;
            const r = ws.addRow([
                it.id, it.sub || '—', Number(it.qtyDay) || 0, Number(it.price) || 0,
                `${((Number(it.gY1) || 0) * 100).toFixed(1)}%`, `${((Number(it.gY2) || 0) * 100).toFixed(1)}%`, `${((Number(it.gY3) || 0) * 100).toFixed(1)}%`, `${((Number(it.gY4) || 0) * 100).toFixed(1)}%`, `${((Number(it.gY5) || 0) * 100).toFixed(1)}%`,
                ...it.yearlyTotals.map(v => fmtCr(v)),
            ]);
            r.height = 18;
            r.getCell(1).style = bodyStyle(bg, COLORS.text2); r.getCell(2).style = bodyStyle(bg, COLORS.text1);
            r.getCell(3).style = { ...bodyStyle(bg, COLORS.text0, true) }; r.getCell(4).style = { ...bodyStyle(bg, COLORS.text0, true) };
            for (let ci = 5; ci <= 9; ci++) r.getCell(ci).style = bodyStyle(bg, COLORS.text2, true);
            for (let yi = 0; yi < YEARS.length; yi++) r.getCell(10 + yi).style = { ...bodyStyle(bg, COLORS.text0, true) };
        });
        const gt = ws.addRow(['', 'Grand Total', '', '', '', '', '', '', '', ...g.yearlyTotals.map(v => fmtCr(v))]);
        gt.height = 20;
        for (let ci = 1; ci <= 9; ci++) gt.getCell(ci).style = totalStyle(COLORS.goldL);
        for (let yi = 0; yi < YEARS.length; yi++) gt.getCell(10 + yi).style = { ...totalStyle(COLORS.gold), alignment: { horizontal: 'right' } };
        ws.addRow([]);
    });
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

function buildCapexSheet(wb, capex = []) {
    const ws = wb.addWorksheet('A.III CAPEX');
    ws.mergeCells('A1:I1');
    const t = ws.getCell('A1');
    t.value = 'A.III CAPEX — Capital Expenditure';
    t.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' } };
    ws.getRow(1).height = 30;

    ws.columns = [
        { key: 'sno', width: 6 }, { key: 'cat', width: 34 }, { key: 'total', width: 14 },
        { key: 'y1', width: 14 }, { key: 'y2', width: 14 }, { key: 'y3', width: 14 }, { key: 'y4', width: 14 }, { key: 'y5', width: 14 },
    ];

    const hdr = ws.addRow(['#', 'Nature of Expense', 'Total (₹)', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5']);
    hdr.height = 26;
    hdr.eachCell(c => Object.assign(c, { style: headerStyle() }));

    (capex || []).forEach((cat) => {
        const catTotal = (cat.items || []).reduce((s, it) => s + (it.total || 0), 0);
        const yTotals = [1, 2, 3, 4, 5].map(n => (cat.items || []).reduce((s, it) => s + (it[`y${n}`] || 0), 0));

        const cr = ws.addRow([cat.sno, cat.category || '—', fmtCr(catTotal), ...yTotals.map(v => fmtCr(v))]);
        cr.height = 22;
        cr.getCell(1).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.gold}` }, size: 10, name: 'Calibri' } };
        cr.getCell(2).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.text0}` }, size: 10, name: 'Calibri' } };
        for (let ci = 3; ci <= 8; ci++) cr.getCell(ci).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.teal}` }, size: 10, name: 'Calibri' }, alignment: { horizontal: 'right' } };

        (cat.items || []).forEach((it, ii) => {
            const bg = ii % 2 === 0 ? 'F2F5FA' : COLORS.white;
            const r = ws.addRow(['', it.name, fmtCr(it.total || 0), ...[1, 2, 3, 4, 5].map(n => fmtCr(it[`y${n}`] || 0))]);
            r.height = 18;
            r.getCell(1).style = bodyStyle(bg, COLORS.text2);
            r.getCell(2).style = bodyStyle(bg, COLORS.text1);
            for (let ci = 3; ci <= 8; ci++) r.getCell(ci).style = { ...bodyStyle(bg, COLORS.text0, true) };
        });
        ws.addRow([]);
    });
}

function buildCapitalCostingSheet(wb, capex = []) {
    const ws = wb.addWorksheet('3b. Costing - Cap Exp');
    ws.mergeCells('A1:H1');
    const t = ws.getCell('A1');
    t.value = '3b. Costing — Capital Expenditure (Summary)';
    t.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' } };
    ws.getRow(1).height = 30;

    ws.columns = [
        { key: 'sno', width: 6 }, { key: 'cat', width: 34 }, { key: 'total', width: 14 },
        { key: 'y1', width: 14 }, { key: 'y2', width: 14 }, { key: 'y3', width: 14 }, { key: 'y4', width: 14 }, { key: 'y5', width: 14 },
    ];

    const hdr = ws.addRow(['#', 'Nature of Expense', 'Total (₹)', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5']);
    hdr.height = 26;
    hdr.eachCell(c => Object.assign(c, { style: headerStyle() }));

    (capex || []).forEach((cat) => {
        const catTotal = (cat.items || []).reduce((s, it) => s + (it.total || 0), 0);
        const yTotals = [1, 2, 3, 4, 5].map(n => (cat.items || []).reduce((s, it) => s + (it[`y${n}`] || 0), 0));

        const cr = ws.addRow([cat.sno, cat.category || '—', fmtCr(catTotal), ...yTotals.map(v => fmtCr(v))]);
        cr.height = 22;
        cr.getCell(1).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.gold}` }, size: 10, name: 'Calibri' } };
        cr.getCell(2).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.text0}` }, size: 10, name: 'Calibri' } };
        for (let ci = 3; ci <= 8; ci++) cr.getCell(ci).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.teal}` }, size: 10, name: 'Calibri' }, alignment: { horizontal: 'right' } };

        (cat.items || []).forEach((it, ii) => {
            const bg = ii % 2 === 0 ? 'F2F5FA' : COLORS.white;
            const r = ws.addRow(['', it.name, fmtCr(it.total || 0), ...[1, 2, 3, 4, 5].map(n => fmtCr(it[`y${n}`] || 0))]);
            r.height = 18;
            r.getCell(1).style = bodyStyle(bg, COLORS.text2);
            r.getCell(2).style = bodyStyle(bg, COLORS.text1);
            for (let ci = 3; ci <= 8; ci++) r.getCell(ci).style = { ...bodyStyle(bg, COLORS.text0, true) };
        });
        ws.addRow([]);
    });
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
                alignment: { horizontal: ci <= 2 ? 'left' : 'right', vertical: 'middle' },
            };
            if (row.isTotal) c.style.border = { top: { style: 'thin', color: { argb: `FF${COLORS.gold}` } } };
        });
    });
}

function buildSalesSheet(wb, groups = [], phaseName = 'P1') {
    const ws = wb.addWorksheet(`B.I Sales - ${phaseName}`);
    const computed = calcRevYearly(groups);

    ws.mergeCells(`A1:P1`);
    const t = ws.getCell('A1');
    t.value = `B.I Sales — ${phaseName} Monthly View (Year 1: 2026-27)`;
    t.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' } };
    ws.getRow(1).height = 30;

    ws.columns = [
        { key: 'id', width: 6 }, { key: 'cat', width: 22 }, { key: 'sub', width: 22 },
        ...MONTHS_Y1.map(() => ({ width: 12 })), { key: 'annual', width: 16 }
    ];

    const hdr = ws.addRow(['#', 'Services', 'Sub Services', ...MONTHS_Y1, 'Annual Total']);
    hdr.height = 26;
    hdr.eachCell(c => Object.assign(c, { style: headerStyle() }));

    computed.forEach((g) => {
        // Group header
        const gr = ws.addRow([g.id, g.header || '—', '', ...MONTHS_Y1.map(() => ''), fmtCr(g.yearlyTotals[0] || 0)]);
        gr.height = 22;
        gr.getCell(1).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.gold}` }, size: 10, name: 'Calibri' } };
        gr.getCell(2).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.text0}` }, size: 10, name: 'Calibri' } };
        for (let ci = 3; ci <= 16; ci++) {
            gr.getCell(ci).style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.teal}` }, size: 10, name: 'Calibri' }, alignment: { horizontal: 'right' } };
        }

        (g.items || []).filter(it => it.sub || (it.yearlyTotals && it.yearlyTotals[0] > 0)).forEach((it, ii) => {
            const bg = ii % 2 === 0 ? 'F2F5FA' : COLORS.white;
            const ann = it.yearlyTotals[0] || 0;
            const mVal = ann / 12; // Flat monthly allocation for simplicity matching UI
            const r = ws.addRow([it.id, '', it.sub || '—', ...MONTHS_Y1.map(() => fmtCr(mVal)), fmtCr(ann)]);
            r.height = 18;
            r.getCell(1).style = bodyStyle(bg, COLORS.text2);
            r.getCell(3).style = bodyStyle(bg, COLORS.text1);
            for (let ci = 4; ci <= 16; ci++) r.getCell(ci).style = { ...bodyStyle(bg, COLORS.text0, true) };
        });
        ws.addRow([]);
    });
}

function buildBalanceSheet(wb, revP1 = [], opexP1 = []) {
    const ws = wb.addWorksheet('5. Balance sheet');
    ws.mergeCells('A1:G1');
    const t = ws.getCell('A1');
    t.value = '5. Balance Sheet — 5-Year Projection';
    t.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' } };
    ws.getRow(1).height = 30;
    ws.columns = [{ width: 36 }, ...YEARS.map(() => ({ width: 18 }))];

    const hdr = ws.addRow(['Particulars', ...YEARS]);
    hdr.height = 26;
    hdr.eachCell(c => Object.assign(c, { style: headerStyle() }));

    const revByYear = YEARS.map((_, yi) => calcRevYearly(revP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
    const opexByYear = YEARS.map((_, yi) => calcOpexYearly(opexP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
    const pat = YEARS.map((_, yi) => { const e = revByYear[yi] - opexByYear[yi] - 75000 * 12; return e - Math.max(0, e * 0.25); });
    const retainedEarnings = YEARS.map((_, yi) => pat.slice(0, yi + 1).reduce((s, v) => s + v, 0));
    const otherCL = [3509865, 4211838, 5054205, 6065046, 7278056];
    const fixedAsset = [425000, 403750, 339521, 453865, 466535];
    const investments = [4100000, 4300000, 5700000, 5800000, 6300000];
    const secDep = [1500000, 1500000, 1500000, 1500000, 1500000];
    const curAdv = [965000, 1302750, 1758712, 1934583, 2176406];
    const otherCA = [1805430, 4262000, 4773440, 4964377, 5460815];

    const sections = [
        { label: 'LIABILITIES', type: 'header' },
        { label: 'Shareholder Funds', type: 'section' },
        { label: 'Capital Account', vals: YEARS.map(() => 0) },
        { label: 'Add: Net Profit / Retained Earnings', vals: retainedEarnings },
        { label: 'Shareholder Funds', vals: retainedEarnings, type: 'total' },
        { type: 'spacer' },
        { label: 'Current Liabilities', type: 'section' },
        { label: 'Other Current Liabilities', vals: otherCL },
        { label: 'Total Liabilities', vals: YEARS.map((_, yi) => retainedEarnings[yi] + otherCL[yi]), type: 'total' },
        { type: 'spacer' },
        { label: 'ASSETS', type: 'header' },
        { label: 'Fixed Assets', type: 'section' },
        { label: 'Net Fixed Assets', vals: fixedAsset },
        { label: 'Investments', vals: investments },
        { type: 'spacer' },
        { label: 'Current Assets', type: 'section' },
        { label: 'Security Deposits', vals: secDep },
        { label: 'Current Advances', vals: curAdv },
        { label: 'Other Current Assets', vals: otherCA },
        { label: 'Cash & Bank Balance', vals: YEARS.map((_, yi) => Math.max(0, retainedEarnings[yi] + otherCL[yi] - fixedAsset[yi] - investments[yi] - secDep[yi] - curAdv[yi] - otherCA[yi])) },
        { label: 'Total Assets', vals: YEARS.map((_, yi) => fixedAsset[yi] + investments[yi] + secDep[yi] + curAdv[yi] + otherCA[yi]), type: 'total' },
    ];

    sections.forEach((row, i) => {
        if (row.type === 'spacer') { ws.addRow([]); return; }
        if (row.type === 'header') {
            const r = ws.addRow([row.label, ...YEARS.map(() => '')]);
            r.eachCell(c => c.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.sectionBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.gold}` }, size: 10 } });
            return;
        }
        if (row.type === 'section') {
            const r = ws.addRow([row.label, ...YEARS.map(() => '')]);
            r.eachCell(c => c.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { italic: true, color: { argb: `FF${COLORS.text1}` }, size: 10 } });
            return;
        }

        const bg = row.type === 'total' ? COLORS.totalBg : (i % 2 === 0 ? 'F2F5FA' : COLORS.white);
        const r = ws.addRow([row.label, ...row.vals.map(v => fmtCr(v))]);
        r.height = row.type === 'total' ? 22 : 18;
        let ci = 0;
        r.eachCell(c => {
            ci++;
            c.style = {
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${bg}` } },
                font: { bold: row.type === 'total', color: { argb: `FF${row.type === 'total' ? COLORS.text0 : COLORS.text1}` }, size: 10 },
                alignment: { horizontal: ci === 1 ? 'left' : 'right' },
            };
            if (row.type === 'total') c.style.border = { top: { style: 'thin', color: { argb: `FF${COLORS.gold}` } } };
        });
    });
}

function buildRatiosSheet(wb, revP1 = [], opexP1 = []) {
    const ws = wb.addWorksheet('6. Ratios');
    ws.mergeCells('A1:G1');
    const t = ws.getCell('A1');
    t.value = '6. Analysis of Ratios';
    t.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${COLORS.headerBg}` } }, font: { bold: true, color: { argb: `FF${COLORS.goldL}` }, size: 13, name: 'Calibri' } };
    ws.getRow(1).height = 30;
    ws.columns = [{ width: 36 }, ...YEARS.map(() => ({ width: 18 }))];

    const hdr = ws.addRow(['Particulars', ...YEARS]);
    hdr.height = 26;
    hdr.eachCell(c => Object.assign(c, { style: headerStyle() }));

    const revByYear = YEARS.map((_, yi) => calcRevYearly(revP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
    const opexByYear = YEARS.map((_, yi) => calcOpexYearly(opexP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
    const pat = YEARS.map((_, yi) => { const e = revByYear[yi] - opexByYear[yi] - 75000 * 12; return e - Math.max(0, e * 0.25); });
    const fixedAssets = [425000, 403750, 339521, 453865, 466535];

    const rows = [
        { label: 'Gross Receipts', vals: revByYear, type: 'currency' },
        { label: 'Net Profit After Deprn Before Tax', vals: YEARS.map((_, yi) => revByYear[yi] - opexByYear[yi] - 75000 * 12), type: 'currency' },
        { label: 'Fixed Assets', vals: fixedAssets, type: 'currency' },
        { label: 'Net Profit Ratio', vals: YEARS.map((_, yi) => revByYear[yi] > 0 ? pat[yi] / revByYear[yi] : 0), type: 'pct' },
        { label: 'Net Sales / Fixed Assets', vals: YEARS.map((_, yi) => fixedAssets[yi] > 0 ? revByYear[yi] / fixedAssets[yi] : 0), type: 'multiple' },
    ];

    rows.forEach((row, i) => {
        const bg = i % 2 === 0 ? 'F2F5FA' : COLORS.white;
        const r = ws.addRow([
            row.label,
            ...row.vals.map(v => row.type === 'pct' ? fmtPct(v) : row.type === 'multiple' ? `${v.toFixed(2)}x` : fmtCr(v))
        ]);
        r.height = 18;
        let ci = 0;
        r.eachCell(c => {
            ci++;
            c.style = {
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${bg}` } },
                font: { color: { argb: `FF${COLORS.text0}` }, size: 10 },
                alignment: { horizontal: ci === 1 ? 'left' : 'right' },
            };
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
        const { basics = {}, revP1 = [], revP2 = [], opexP1 = [], capex = [], loan1 = {}, loan2 = {}, totalProjectCost = {} } = state;

        const wb = new ExcelJS.Workbook();
        wb.creator = 'OnEasy Financial Model';
        wb.created = new Date();
        wb.modified = new Date();

        buildBasicsSheet(wb, basics, loan1, loan2, totalProjectCost);
        buildRevenueSheet(wb, revP1);
        buildRevenueSheetP2(wb, revP2);
        buildOpexSheet(wb, opexP1);
        buildCapexSheet(wb, capex);
        buildCapitalCostingSheet(wb, capex);
        buildSalesSheet(wb, revP1, 'P1');
        buildSalesSheet(wb, revP2, 'P2');
        buildPLSheet(wb, revP1, opexP1, loan1, loan2);
        buildBalanceSheet(wb, revP1, opexP1);
        buildRatiosSheet(wb, revP1, opexP1);
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
