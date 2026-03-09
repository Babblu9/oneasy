import { HyperFormula } from 'hyperformula';

function makeSafeSheetName(name) {
    const base = String(name || 'Sheet').replace(/[^A-Za-z0-9_]/g, '_').replace(/^_+/, '') || 'Sheet';
    return /^[A-Za-z_]/.test(base) ? base : `S_${base}`;
}

function buildSheetMap(sheetNames) {
    const used = new Set();
    const map = {};

    for (const name of sheetNames) {
        let safe = makeSafeSheetName(name);
        let n = 1;
        while (used.has(safe)) {
            safe = `${makeSafeSheetName(name)}_${n++}`;
        }
        used.add(safe);
        map[name] = safe;
    }

    return map;
}

function replaceSheetRefs(formula, sheetMap) {
    let out = formula;
    const entries = Object.entries(sheetMap).sort((a, b) => b[0].length - a[0].length);

    for (const [original, safe] of entries) {
        const escaped = original.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&');
        out = out.replace(new RegExp(`'${escaped}'!`, 'g'), `${safe}!`);
        out = out.replace(new RegExp(`(^|[^A-Za-z0-9_'])${escaped}!`, 'g'), `$1${safe}!`);
    }

    return out;
}

function transpileFormula(rawFormula, sheetMap) {
    if (rawFormula == null) return null;
    let formula = String(rawFormula).trim();
    if (!formula.startsWith('=')) formula = '=' + formula;

    formula = formula
        .replace(/_xlfn\./gi, '')
        .replace(/_xlws\./gi, '')
        .replace(/@/g, '')
        .replace(/;/g, ',');

    return replaceSheetRefs(formula, sheetMap);
}

function normalizeCellValue(cell) {
    if (cell.formula || cell.sharedFormula || (cell.value && typeof cell.value === 'object' && cell.value.formula)) {
        const rawFormula = cell.formula || cell.sharedFormula || cell.value.formula;
        return { __formula__: rawFormula };
    }

    const val = cell.value;
    if (val == null) return null;

    if (typeof val === 'object') {
        if (val.richText) return val.richText.map(rt => rt.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return val.result;
        if (val.error) return null;
    }

    return val;
}

export function excelAddressToCoords(address) {
    const m = /^([A-Z]{1,3})(\d+)$/i.exec(String(address || '').trim());
    if (!m) return null;

    const colLetters = m[1].toUpperCase();
    const row = Number(m[2]) - 1;
    let col = 0;
    for (const ch of colLetters) col = col * 26 + (ch.charCodeAt(0) - 64);

    return { row, col: col - 1 };
}

function getNumericRowValues(hf, sheetId, row, maxCols = 80) {
    const vals = [];
    for (let c = 0; c < maxCols; c++) {
        const v = hf.getCellValue({ sheet: sheetId, row, col: c });
        if (typeof v === 'number' && Number.isFinite(v)) vals.push(v);
    }
    return vals;
}

function findRowByLabel(sheetData, regex) {
    for (let r = 0; r < sheetData.length; r++) {
        for (let c = 0; c < Math.min(6, (sheetData[r] || []).length); c++) {
            const v = sheetData[r]?.[c];
            if (typeof v === 'string' && regex.test(v.trim())) return r;
        }
    }
    return -1;
}

function sumPositive(values) {
    return values.filter(v => typeof v === 'number' && v > 0).reduce((a, b) => a + b, 0);
}

function findFirstPositiveMonth(values) {
    const idx = values.findIndex(v => typeof v === 'number' && v > 0);
    return idx >= 0 ? idx + 1 : null;
}

function pickYears(values) {
    const positives = values.filter(v => typeof v === 'number' && v > 0);
    if (positives.length >= 2) return { y1: positives[0], y2: positives[1] };
    if (positives.length === 1) return { y1: positives[0], y2: positives[0] * 1.2 };
    return { y1: 0, y2: 0 };
}

function pickBestRevenueRow(hf, sheetId, maxRows = 30, maxCols = 120) {
    let best = { row: -1, score: -1, vals: [] };
    for (let r = 0; r < maxRows; r++) {
        const vals = getNumericRowValues(hf, sheetId, r, maxCols);
        if (vals.length < 3) continue;
        const positives = vals.filter(v => v > 0);
        if (positives.length < 2) continue;
        const monotonicCount = positives.slice(1).filter((v, i) => v >= positives[i]).length;
        const score = sumPositive(positives) + monotonicCount * 1e6;
        if (score > best.score) best = { row: r, score, vals: positives };
    }
    return best;
}

export function buildHyperFormulaFromWorkbook(workbook) {
    const sheetNames = workbook.worksheets.map(ws => ws.name);
    const sheetMap = buildSheetMap(sheetNames);

    const rawDataByOriginalSheet = {};
    const hfInputBySafeSheet = {};

    for (const ws of workbook.worksheets) {
        const safeName = sheetMap[ws.name];
        const rows = Math.max(ws.rowCount, 1);
        const cols = Math.max(ws.columnCount, 1);

        const rawGrid = [];
        const hfGrid = [];

        for (let r = 1; r <= rows; r++) {
            const rowRaw = [];
            const rowHf = [];
            const rowObj = ws.getRow(r);

            for (let c = 1; c <= cols; c++) {
                const cell = rowObj.getCell(c);
                const normalized = normalizeCellValue(cell);
                if (normalized && typeof normalized === 'object' && normalized.__formula__) {
                    const hfFormula = transpileFormula(normalized.__formula__, sheetMap);
                    rowHf.push(hfFormula);
                    rowRaw.push(String(cell.value?.result ?? ''));
                } else {
                    rowHf.push(normalized);
                    rowRaw.push(normalized);
                }
            }

            rawGrid.push(rowRaw);
            hfGrid.push(rowHf);
        }

        rawDataByOriginalSheet[ws.name] = rawGrid;
        hfInputBySafeSheet[safeName] = hfGrid;
    }

    const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });

    for (const safeName of Object.values(sheetMap)) {
        hf.addSheet(safeName);
        const sheetId = hf.getSheetId(safeName);
        hf.setCellContents({ sheet: sheetId, row: 0, col: 0 }, hfInputBySafeSheet[safeName]);
    }

    return { hf, sheetMap, rawDataByOriginalSheet };
}

export function getHfCellValue(hf, sheetMap, originalSheetName, address) {
    const safeName = sheetMap[originalSheetName] || sheetMap[Object.keys(sheetMap).find(k => k.toLowerCase() === originalSheetName.toLowerCase())];
    if (!safeName) return null;

    const sheetId = hf.getSheetId(safeName);
    if (sheetId == null) return null;

    const coords = excelAddressToCoords(address);
    if (!coords) return null;

    const v = hf.getCellValue({ sheet: sheetId, row: coords.row, col: coords.col });
    return typeof v === 'number' && Number.isFinite(v) ? v : null;
}

export function applyUpdatesToHyperFormula({ hf, sheetMap, updates }) {
    if (!Array.isArray(updates) || updates.length === 0) return;

    for (const u of updates) {
        const originalSheet = String(u?.sheet || '');
        const safeName =
            sheetMap[originalSheet] ||
            sheetMap[Object.keys(sheetMap).find(k => k.toLowerCase() === originalSheet.toLowerCase())];
        if (!safeName) continue;

        const sheetId = hf.getSheetId(safeName);
        if (sheetId == null) continue;

        const coords = excelAddressToCoords(u?.cell);
        if (!coords) continue;

        hf.setCellContents(
            { sheet: sheetId, row: coords.row, col: coords.col },
            [[u.value]]
        );
    }
}

export function extractDashboardMetrics({ hf, sheetMap, rawDataByOriginalSheet }) {
    const pnlSheet = Object.keys(sheetMap).find(n => /p&l/i.test(n)) || '4. P&L';
    const cashSheet = Object.keys(sheetMap).find(n => /cash\s*flow/i.test(n)) || '';

    const pnlSafe = sheetMap[pnlSheet];
    const pnlId = pnlSafe ? hf.getSheetId(pnlSafe) : null;
    const pnlRaw = rawDataByOriginalSheet[pnlSheet] || [];

    let year1Revenue = 0;
    let year2Revenue = 0;
    let ebitda = 0;
    let breakEvenMonth = null;

    if (pnlId != null) {
        const revRow = findRowByLabel(pnlRaw, /total\s+revenue|revenue\s+total|sales\s+total/i);
        const ebitdaRow = findRowByLabel(pnlRaw, /ebitda|ebidta/i);
        const patRow = findRowByLabel(pnlRaw, /net\s+profit|pat/i);

        if (revRow >= 0) {
            const revVals = getNumericRowValues(hf, pnlId, revRow, 120);
            const years = pickYears(revVals);
            year1Revenue = years.y1;
            year2Revenue = years.y2;
        } else {
            const bestRevenue = pickBestRevenueRow(hf, pnlId, 30, 120);
            if (bestRevenue.row >= 0) {
                const years = pickYears(bestRevenue.vals);
                year1Revenue = years.y1;
                year2Revenue = years.y2;
            }
        }

        if (ebitdaRow >= 0) {
            const eVals = getNumericRowValues(hf, pnlId, ebitdaRow, 120);
            ebitda = (eVals.find(v => v !== 0) ?? 0);
            if (!breakEvenMonth) breakEvenMonth = findFirstPositiveMonth(eVals);
        }

        if (!breakEvenMonth && patRow >= 0) {
            const pVals = getNumericRowValues(hf, pnlId, patRow, 120);
            breakEvenMonth = findFirstPositiveMonth(pVals);
        }
    }

    let cashRunwayMonths = 18;
    if (cashSheet && sheetMap[cashSheet]) {
        const cashId = hf.getSheetId(sheetMap[cashSheet]);
        const cashRaw = rawDataByOriginalSheet[cashSheet] || [];
        const closeCashRow = findRowByLabel(cashRaw, /closing\s+cash|cash\s+balance/i);
        if (closeCashRow >= 0 && cashId != null) {
            const cVals = getNumericRowValues(hf, cashId, closeCashRow, 120);
            const positiveMonths = cVals.filter(v => typeof v === 'number' && v > 0).length;
            if (positiveMonths > 0) cashRunwayMonths = positiveMonths;
        }
    }

    if (!breakEvenMonth) breakEvenMonth = ebitda > 0 ? 12 : 24;

    // Fallback if labels are not found in template variant
    if (!year1Revenue || !year2Revenue) {
        const revSheet = Object.keys(sheetMap).find(n => /revenue\s*streams/i.test(n));
        if (revSheet) {
            const revId = hf.getSheetId(sheetMap[revSheet]);
            const sumRows = [];
            for (let r = 0; r < 200; r++) {
                const vals = getNumericRowValues(hf, revId, r, 20);
                if (vals.length >= 2) sumRows.push(vals);
            }
            const rowTotals = sumRows.map(sumPositive).filter(v => v > 0).slice(-20);
            year1Revenue = year1Revenue || (rowTotals[0] || 0);
            year2Revenue = year2Revenue || (rowTotals[1] || year1Revenue * 1.2);
        }
    }

    return {
        year1Revenue: Math.round(year1Revenue || 0),
        year2Revenue: Math.round(year2Revenue || 0),
        ebitda: Math.round(ebitda || 0),
        breakEvenMonth,
        cashRunwayMonths,
    };
}
