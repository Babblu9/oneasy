const fs = require('fs');
const path = require('path');

function makeSafeSheetName(name) {
  const base = String(name || 'Sheet').replace(/[^A-Za-z0-9_]/g, '_').replace(/^_+/, '') || 'Sheet';
  return /^[A-Za-z_]/.test(base) ? base : `S_${base}`;
}

function buildSheetMap(sheetNames) {
  const map = {};
  const used = new Set();

  for (const original of sheetNames || []) {
    let safe = makeSafeSheetName(original);
    let i = 1;
    while (used.has(safe)) {
      safe = `${makeSafeSheetName(original)}_${i++}`;
    }
    used.add(safe);
    map[original] = safe;
  }

  return map;
}

function replaceSheetRefs(formula, sheetMap) {
  let out = formula;
  const entries = Object.entries(sheetMap).sort((a, b) => b[0].length - a[0].length);

  for (const [original, safe] of entries) {
    const escaped = original.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&');

    const quotedRef = new RegExp(`'${escaped}'!`, 'g');
    out = out.replace(quotedRef, `${safe}!`);

    const unquotedRef = new RegExp(`(^|[^A-Za-z0-9_'])${escaped}!`, 'g');
    out = out.replace(unquotedRef, `$1${safe}!`);
  }

  return out;
}

function transpileFormula(rawFormula, sheetMap) {
  if (rawFormula == null) return rawFormula;
  let formula = String(rawFormula).trim();
  if (!formula.startsWith('=')) formula = '=' + formula;

  // Excel-specific prefixes and operators not understood by HyperFormula
  formula = formula
    .replace(/_xlfn\./gi, '')
    .replace(/_xlws\./gi, '')
    .replace(/@/g, '');

  // Locale normalization fallback (most templates already use comma)
  // Convert semicolon argument separators to comma when present.
  formula = formula.replace(/;/g, ',');

  formula = replaceSheetRefs(formula, sheetMap);

  return formula;
}

function transpileRawFormulaJson(inputPath, outputPath) {
  const raw = JSON.parse(fs.readFileSync(inputPath, 'utf8'));
  const sheetMap = buildSheetMap(raw.sheets || []);

  const transpiledFormulas = (raw.formulas || []).map((item) => ({
    ...item,
    hfSheet: sheetMap[item.sheet] || makeSafeSheetName(item.sheet),
    hfFormula: transpileFormula(item.formula, sheetMap),
  }));

  const output = {
    source: path.basename(inputPath),
    generatedAt: new Date().toISOString(),
    totalFormulas: transpiledFormulas.length,
    sheetMap,
    sheets: (raw.sheets || []).map((name) => ({ original: name, safe: sheetMap[name] || makeSafeSheetName(name) })),
    formulas: transpiledFormulas,
  };

  fs.writeFileSync(outputPath, JSON.stringify(output, null, 2));
  return output;
}

if (require.main === module) {
  const inputPath = process.argv[2]
    ? path.resolve(process.argv[2])
    : path.join(process.cwd(), 'info', 'raw_formulas.json');
  const outputPath = process.argv[3]
    ? path.resolve(process.argv[3])
    : path.join(process.cwd(), 'info', 'hyperformula_formulas.json');

  if (!fs.existsSync(inputPath)) {
    console.error(`Input file not found: ${inputPath}`);
    process.exit(1);
  }

  const result = transpileRawFormulaJson(inputPath, outputPath);
  console.log(`Transpiled ${result.totalFormulas} formulas -> ${outputPath}`);
}

module.exports = {
  buildSheetMap,
  transpileFormula,
  transpileRawFormulaJson,
};
