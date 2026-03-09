import { NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs';
import XlsxPopulate from 'xlsx-populate';
import { findOpexRow, findRevenueRow } from '@/lib/excelCellMap';
import {
  createDeterministicRowAllocator,
  SLOT_RANGES,
  detectRevenueColumnsFromPopulate,
  detectOpexColumnsFromPopulate
} from '@/lib/streamInjection';

const SLOT_CONFIG = SLOT_RANGES;

function toNumber(value, fallback = 0) {
  if (value == null || value === '') return fallback;
  if (typeof value === 'number') return Number.isFinite(value) ? value : fallback;
  const s = String(value).replace(/[,$₹\s]/g, '');
  const n = Number(s);
  return Number.isFinite(n) ? n : fallback;
}

function asString(value) {
  return String(value || '').trim();
}

function isSkippableRowLabel(text) {
  const t = String(text || '').toLowerCase();
  return t.includes('total') || t.includes('grand') || t.includes('phase');
}

function getCellFormula(cell) {
  try {
    return cell?.formula?.() || null;
  } catch {
    return null;
  }
}

function isFormulaCell(cell) {
  return !!getCellFormula(cell);
}

function safeWriteCell(sheet, address, value, report) {
  const cell = sheet.cell(address);
  if (isFormulaCell(cell)) {
    report.formulaWriteBlocks.push(address);
    return false;
  }
  cell.value(value);
  return true;
}

function getCandidateRows(sheet, cfg) {
  const rows = [];
  for (let r = cfg.startRow; r <= cfg.endRow; r += 1) {
    const d = asString(sheet.cell(`D${r}`).value());
    const e = asString(sheet.cell(`E${r}`).value());
    if (isSkippableRowLabel(d) || isSkippableRowLabel(e)) continue;
    rows.push(r);
  }
  return rows;
}

function extractRevenueItems(financialData) {
  const out = [];
  const streams = Array.isArray(financialData?.revenueStreams) ? financialData.revenueStreams : [];

  for (const stream of streams) {
    if (Array.isArray(stream?.subStreams)) {
      for (const sub of stream.subStreams) {
        if (Array.isArray(sub?.products)) {
          for (const prod of sub.products) {
            const name = asString(prod?.name || prod?.productName || sub?.name || stream?.name);
            if (!name) continue;
            out.push({
              name,
              qty: toNumber(prod?.units ?? prod?.quantity, 1),
              price: toNumber(prod?.price ?? prod?.value, 0),
              cellRef: prod?.cellRef || null,
            });
          }
        }
      }
    }

    if (!Array.isArray(stream?.subStreams) && asString(stream?.label || stream?.name)) {
      out.push({
        name: asString(stream?.label || stream?.name),
        qty: toNumber(stream?.units ?? stream?.quantity, 1),
        price: toNumber(stream?.price ?? stream?.value, 0),
        cellRef: stream?.cellRef || null,
      });
    }
  }

  return out;
}

function extractOpexItems(financialData) {
  const out = [];
  const items = Array.isArray(financialData?.opex) ? financialData.opex : [];

  for (const item of items) {
    const name = asString(item?.subCategory || item?.category || item?.label || item?.name);
    if (!name) continue;
    out.push({
      name,
      qty: toNumber(item?.units ?? item?.quantity, 1),
      cost: toNumber(item?.monthlyCost ?? item?.price ?? item?.value, 0),
      cellRef: item?.cellRef || null,
    });
  }

  return out;
}

function injectFallbackRows({ sheet, items, cfg, report, type, columns, allocator }) {
  if (!sheet || !items.length) return [];

  const candidateRows = new Set(getCandidateRows(sheet, cfg));
  const unresolved = [];

  for (const item of items) {
    let assigned = false;
    const row = allocator(item.name);
    if (row == null || !candidateRows.has(row)) {
      unresolved.push(item);
      continue;
    }

    const labelAddr = `${columns.labelCol}${row}`;
    const qtyAddr = `${columns.qtyCol}${row}`;
    const valueAddr = `${columns.valueCol}${row}`;

    const okLabel = safeWriteCell(sheet, labelAddr, item.name, report);
    const okQty = safeWriteCell(sheet, qtyAddr, item.qty, report);
    const okValue = safeWriteCell(sheet, valueAddr, type === 'revenue' ? item.price : item.cost, report);
    const okSecondaryLabel = columns.secondaryLabelCol
      ? safeWriteCell(sheet, `${columns.secondaryLabelCol}${row}`, item.name, report)
      : true;

    if (okLabel && okQty && okValue && okSecondaryLabel) {
      report.fallbackAssignments.push({ type, name: item.name, row });
      assigned = true;
    }

    if (!assigned) {
      unresolved.push(item);
    }
  }

  return unresolved;
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { template, financialData } = body;

    if (!template || !financialData || typeof financialData !== 'object') {
      return NextResponse.json(
        { error: 'Invalid request. "template" and "financialData" object are required.' },
        { status: 400 }
      );
    }

    let templatePath = path.join(process.cwd(), 'excel-templates', template);
    if (!fs.existsSync(templatePath)) {
      const alt = path.join(process.cwd(), 'templates', template);
      if (fs.existsSync(alt)) templatePath = alt;
    }

    let workbook;
    try {
      workbook = await XlsxPopulate.fromFileAsync(templatePath);
    } catch (err) {
      console.error('Error loading template:', err);
      return NextResponse.json({ error: `Template "${template}" not found or invalid.` }, { status: 404 });
    }

    const report = {
      formulaWriteBlocks: [],
      unresolvedRevenue: [],
      unresolvedOpex: [],
      fallbackAssignments: [],
      directAssignments: { revenue: 0, opex: 0 },
    };

    const fd = financialData || {};
    const basicInfo = fd.businessInfo || {};
    const branchCount = toNumber(fd.branchCount, fd.branches?.length || 1);
    const revenueItems = extractRevenueItems(fd);
    const opexItems = extractOpexItems(fd);

    const basicsSheet = workbook.sheet('1. Basics');
    if (basicsSheet) {
      if (basicInfo.legalName) safeWriteCell(basicsSheet, 'D2', basicInfo.legalName, report);
      if (basicInfo.tradeName) safeWriteCell(basicsSheet, 'D4', basicInfo.tradeName, report);
      if (basicInfo.address) safeWriteCell(basicsSheet, 'D6', basicInfo.address, report);
      if (basicInfo.email) safeWriteCell(basicsSheet, 'D8', basicInfo.email, report);
      if (basicInfo.phone) safeWriteCell(basicsSheet, 'D10', basicInfo.phone, report);
    }

    const revSheet = workbook.sheet('A.I Revenue Streams - P1');
    if (revSheet) {
      safeWriteCell(revSheet, 'H7', branchCount, report);
    }

    const opexSheet = workbook.sheet('A.IIOPEX');
    if (opexSheet) {
      safeWriteCell(opexSheet, 'I8', branchCount, report);
    }

    const revenueAllocator = createDeterministicRowAllocator(SLOT_CONFIG.revenue);
    const opexAllocator = createDeterministicRowAllocator(SLOT_CONFIG.opex);

    const unresolvedRevenue = [];
    if (revSheet) {
      for (const item of revenueItems) {
        const fallback = item.name ? findRevenueRow(item.name) : null;
        const qtyCell = item?.cellRef?.qty || fallback?.qty;
        const priceCell = item?.cellRef?.price || fallback?.price;

        let wrote = false;
        if (qtyCell) wrote = safeWriteCell(revSheet, qtyCell, item.qty, report) || wrote;
        if (priceCell) wrote = safeWriteCell(revSheet, priceCell, item.price, report) || wrote;

        if (wrote) {
          report.directAssignments.revenue += 1;
        } else {
          unresolvedRevenue.push(item);
        }
      }
    }

    const unresolvedOpex = [];
    if (opexSheet) {
      for (const item of opexItems) {
        const fallback = item.name ? findOpexRow(item.name) : null;
        const costCell = item?.cellRef?.cost || fallback?.cost;

        let wrote = false;
        if (costCell) wrote = safeWriteCell(opexSheet, costCell, item.cost, report) || wrote;

        if (wrote) {
          report.directAssignments.opex += 1;
        } else {
          unresolvedOpex.push(item);
        }
      }
    }

    const revenueColumns = revSheet ? detectRevenueColumnsFromPopulate(revSheet) : { labelCol: 'E', secondaryLabelCol: 'F', qtyCol: 'H', valueCol: 'J' };
    const opexColumns = opexSheet ? detectOpexColumnsFromPopulate(opexSheet) : { labelCol: 'E', secondaryLabelCol: 'D', qtyCol: 'G', valueCol: 'I' };
    report.columnDetection = { revenue: revenueColumns, opex: opexColumns };

    report.unresolvedRevenue = injectFallbackRows({
      sheet: revSheet,
      items: unresolvedRevenue,
      cfg: SLOT_CONFIG.revenue,
      report,
      type: 'revenue',
      columns: revenueColumns,
      allocator: revenueAllocator,
    }).map(({ name, qty, price }) => ({ name, qty, price }));

    report.unresolvedOpex = injectFallbackRows({
      sheet: opexSheet,
      items: unresolvedOpex,
      cfg: SLOT_CONFIG.opex,
      report,
      type: 'opex',
      columns: opexColumns,
      allocator: opexAllocator,
    }).map(({ name, qty, cost }) => ({ name, qty, cost }));

    if (report.formulaWriteBlocks.length > 0) {
      return NextResponse.json({
        error: 'Export blocked: attempted write to formula/protected cells.',
        report,
      }, { status: 409 });
    }

    if (report.unresolvedRevenue.length > 0 || report.unresolvedOpex.length > 0) {
      return NextResponse.json({
        error: 'Export blocked: some streams could not be injected into template slots.',
        report,
      }, { status: 422 });
    }

    const uint8Array = await workbook.outputAsync();
    const response = new NextResponse(uint8Array, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="financial_model_output.xlsx"',
        'Content-Length': uint8Array.byteLength.toString(),
        'X-Fina-Direct-Revenue': String(report.directAssignments.revenue),
        'X-Fina-Direct-Opex': String(report.directAssignments.opex),
        'X-Fina-Fallback-Assignments': String(report.fallbackAssignments.length),
      },
    });

    return response;
  } catch (error) {
    console.error('Error generating financial model:', error);
    return NextResponse.json({ error: 'Internal Server Error' }, { status: 500 });
  }
}
