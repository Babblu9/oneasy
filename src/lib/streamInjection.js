export const SLOT_RANGES = {
    revenue: { startRow: 10, endRow: 200 },
    opex: { startRow: 10, endRow: 200 },
};

const DEFAULT_COLUMNS = {
    revenue: { labelCol: 'E', secondaryLabelCol: 'F', qtyCol: 'H', valueCol: 'J' },
    opex: { labelCol: 'E', secondaryLabelCol: 'D', qtyCol: 'G', valueCol: 'I' },
};

function toColLetter(col) {
    let n = Number(col) || 1;
    let s = '';
    while (n > 0) {
        const rem = (n - 1) % 26;
        s = String.fromCharCode(65 + rem) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s || 'A';
}

export function normalizeStreamKey(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ');
}

export function createDeterministicRowAllocator(range = SLOT_RANGES.revenue) {
    const byName = new Map();
    let next = range.startRow;

    return function allocate(name) {
        const key = normalizeStreamKey(name);
        if (!key) return null;
        if (byName.has(key)) return byName.get(key);
        if (next > range.endRow) return null;
        const row = next++;
        byName.set(key, row);
        return row;
    };
}

function cleanHeader(value) {
    return String(value || '')
        .toLowerCase()
        .replace(/[^a-z0-9\s]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function hasAny(text, keywords) {
    return keywords.some(k => text.includes(k));
}

function detectColumnsFromReader({ type, readCell }) {
    const defaults = DEFAULT_COLUMNS[type];
    const headerRows = [8, 9, 4, 3];
    const maxCol = 16;

    const qtyKeywords = type === 'revenue'
        ? ['number of products sold', 'units', 'quantity', 'volume']
        : ['units', 'qty', 'quantity'];
    const valueKeywords = type === 'revenue'
        ? ['sale price', 'price', 'rate']
        : ['monthly cost', 'cost', 'expense', 'amount'];
    const labelKeywords = type === 'revenue'
        ? ['minor header', 'sub service', 'services', 'product', 'stream']
        : ['minor head', 'sub service', 'category', 'expense', 'particular'];

    let qtyCol = null;
    let valueCol = null;
    let labelCol = null;

    for (const r of headerRows) {
        for (let c = 1; c <= maxCol; c += 1) {
            const addr = `${toColLetter(c)}${r}`;
            const text = cleanHeader(readCell(addr));
            if (!text) continue;
            if (!qtyCol && hasAny(text, qtyKeywords)) qtyCol = toColLetter(c);
            if (!valueCol && hasAny(text, valueKeywords)) valueCol = toColLetter(c);
            if (!labelCol && hasAny(text, labelKeywords)) labelCol = toColLetter(c);
        }
    }

    return {
        labelCol: labelCol || defaults.labelCol,
        secondaryLabelCol: defaults.secondaryLabelCol,
        qtyCol: qtyCol || defaults.qtyCol,
        valueCol: valueCol || defaults.valueCol,
    };
}

export function detectRevenueColumnsFromExcelJs(ws) {
    return detectColumnsFromReader({
        type: 'revenue',
        readCell: (addr) => ws?.getCell(addr)?.value,
    });
}

export function detectOpexColumnsFromExcelJs(ws) {
    return detectColumnsFromReader({
        type: 'opex',
        readCell: (addr) => ws?.getCell(addr)?.value,
    });
}

export function detectRevenueColumnsFromPopulate(sheet) {
    return detectColumnsFromReader({
        type: 'revenue',
        readCell: (addr) => sheet?.cell(addr)?.value(),
    });
}

export function detectOpexColumnsFromPopulate(sheet) {
    return detectColumnsFromReader({
        type: 'opex',
        readCell: (addr) => sheet?.cell(addr)?.value(),
    });
}

export function getDefaultColumns(type) {
    return DEFAULT_COLUMNS[type];
}

