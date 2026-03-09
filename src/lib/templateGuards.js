const ALLOWED_SHEETS = new Set([
    '1. Basics',
    'Basics',
    'A.I Revenue Streams - P1',
    'A.I Revenue Streams - P2',
    'A.II OPEX',
    'A.IIOPEX',
    '2.Total Project Cost',
    'Total Project Cost',
    '4. P&L',
    'P&L',
    '5. Balance Sheet',
    '5. Balance sheet',
    'Balance Sheet',
]);

const SHEET_ALIASES = {
    'Basics': '1. Basics',
    'A.II OPEX': 'A.IIOPEX',
    'Total Project Cost': '2.Total Project Cost',
    'P&L': '4. P&L',
    'Balance Sheet': '5. Balance sheet',
    '5. Balance Sheet': '5. Balance sheet',
    '5. balance sheet': '5. Balance sheet',
};

export function normalizeSheetName(name) {
    const raw = String(name || '').trim();
    return SHEET_ALIASES[raw] || raw;
}

export function isAllowedSheet(name) {
    return ALLOWED_SHEETS.has(String(name || '').trim());
}

export function isValidCellAddress(cell) {
    return /^[A-Z]{1,3}[1-9]\d{0,5}$/.test(String(cell || '').trim());
}

export function normalizeInputValue(value) {
    if (value === '' || value == null) return '';
    if (typeof value === 'number') return value;
    const s = String(value).trim();
    if (s === '') return '';
    const n = Number(s);
    if (!Number.isNaN(n) && Number.isFinite(n)) return n;
    return s;
}

export function isFormulaCell(cell) {
    return !!(cell?.formula || cell?.sharedFormula || (cell?.value && typeof cell.value === 'object' && (cell.value.formula || cell.value.sharedFormula)));
}

export function resolveWorksheet(workbook, name) {
    const normalized = normalizeSheetName(name);
    return (
        workbook.getWorksheet(normalized) ||
        workbook.getWorksheet(String(name || '').trim()) ||
        workbook.worksheets.find(ws => ws.name.toLowerCase() === normalized.toLowerCase()) ||
        workbook.worksheets.find(ws => ws.name.toLowerCase() === String(name || '').trim().toLowerCase()) ||
        null
    );
}
