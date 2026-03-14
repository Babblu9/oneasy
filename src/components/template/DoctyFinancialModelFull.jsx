'use client';

import { useState, useRef, useEffect } from "react";
import { dataActionToPatches } from "@/lib/excelCellMap.js";
import { modelStateToPatches } from "@/lib/modelStateToPatches.js";
import { getIndustryTemplate, getGrowthRates } from "@/lib/engine/industryTemplates";
import ScenarioDashboard from "@/components/template/ScenarioDashboard";
import { Activity, Bot, Send, Download, Sparkles, MessageSquare, TrendingUp, DollarSign, Building, CreditCard, Scale, FileText, PieChart, BarChart3, Layout, ChevronDown, ChevronRight, Calendar, Settings } from "lucide-react";

// ─── THEME ────────────────────────────────────────────────────────────────────
const C = {
  bg0: "#F3F6F9", bg1: "#FFFFFF", bg2: "#F8F9FA", bg3: "#E9ECEF",
  nav: "#FFFFFF", navB: "#F1F5F9",
  border: "#E2E8F0", borderLight: "#F1F5F9",
  gold: "#B7791F", goldL: "#D69E2E", goldD: "#975A16",
  teal: "#2C7A7B", tealL: "#38B2AC",
  green: "#2F855A", greenL: "#48BB78",
  red: "#C53030", redL: "#E53E3E",
  blue: "#2B6CB0", blueL: "#4299E1",
  purple: "#6B46C1",
  text0: "#1A202C", text1: "#2D3748", text2: "#4A5568", text3: "#718096",
  inputBlue: "#2B6CB0",
  sectionBg: "#EDF2F7", totalBg: "#E2E8F0", headerBg: "#E2E8F0",
};

// ─── UTILS ────────────────────────────────────────────────────────────────────
const fmtINR = (v, compact = false) => {
  if (v === null || v === undefined) return "—";
  if (typeof v === "string") return v;
  if (v === 0) return "—";
  const abs = Math.abs(v);
  const neg = v < 0;
  let s;
  if (compact) {
    if (abs >= 10000000) s = `₹${(abs / 10000000).toFixed(2)}Cr`;
    else if (abs >= 100000) s = `₹${(abs / 100000).toFixed(2)}L`;
    else if (abs >= 1000) s = `₹${(abs / 1000).toFixed(1)}K`;
    else s = `₹${abs.toLocaleString("en-IN")}`;
  } else {
    s = `₹${abs.toLocaleString("en-IN", { maximumFractionDigits: 0 })}`;
  }
  return neg ? `(${s})` : s;
};
const fmtPct = (v) => v == null ? "—" : `${(v * 100).toFixed(1)}%`;
const fmtNum = (v) => v == null ? "—" : v === 0 ? "—" : v.toLocaleString("en-IN", { maximumFractionDigits: 0 });
const fmtDate = (d) => { try { return new Date(d).toLocaleDateString("en-IN"); } catch { return d || "—"; } };

const YEARS = ["2026-27", "2027-28", "2028-29", "2029-30", "2030-31"];
const GROWTH_KEYS_REV = ["gY1", "gY2", "gY3", "gY4", "gY5"];
const GROWTH_KEYS_OPEX = ["gY1", "gY2", "gY3", "gY4", "gY5"];
const MONTHS_Y1 = ["Apr '26", "May '26", "Jun '26", "Jul '26", "Aug '26", "Sep '26", "Oct '26", "Nov '26", "Dec '26", "Jan '27", "Feb '27", "Mar '27"];
const MONTHS_Y1_DATES = ["2026-04-30", "2026-05-31", "2026-06-30", "2026-07-31", "2026-08-31", "2026-09-30", "2026-10-31", "2026-11-30", "2026-12-31", "2027-01-31", "2027-02-28", "2027-03-31"];

const extractDataTags = (text) => {
  const out = [];
  // More flexible regex to handle various formats
  const re = /\[DATA:\s*(\{[\s\S]*?\})\]/g;
  let m;
  while ((m = re.exec(String(text || ""))) !== null) {
    try {
      let jsonStr = m[1];
      // Handle escaped quotes
      jsonStr = jsonStr.replace(/\\"/g, '"').replace(/\\\\/g, '\\');
      out.push(JSON.parse(jsonStr));
    } catch (e) {
      console.log("[DEBUG] Failed to parse DATA tag:", m[1]?.substring(0, 100));
    }
  }
  return out;
};

const cleanForDisplay = (text) => String(text || "").replace(/\[DATA:\s*\{[\s\S]*?\}\]/g, "").replace(/\[SUGGESTIONS:\s*\[[\s\S]*?\]\]/g, "").trim();

const extractSuggestions = (text) => {
  const re = /\[SUGGESTIONS:\s*(\[[\s\S]*?\])\]/g;
  const match = re.exec(String(text || ""));
  if (match) {
    try {
      return JSON.parse(match[1].replace(/"/g, '"').replace(/'/g, '"'));
    } catch {
      return null;
    }
  }
  return null;
};

// ─── DRAGGABLE SCROLL WRAPPER ────────────────────────────────────────────────
const DraggableScroll = ({ children, style = {}, className = "" }) => {
  const scrollRef = useRef(null);
  const [isDragging, setIsDragging] = useState(false);
  const [startX, setStartX] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);

  const onMouseDown = (e) => {
    const tag = e.target.tagName?.toLowerCase();
    if (['input', 'select', 'textarea', 'button'].includes(tag)) return;
    if (!scrollRef.current) return;
    setIsDragging(true);
    setStartX(e.pageX - scrollRef.current.offsetLeft);
    setScrollLeft(scrollRef.current.scrollLeft);
  };
  const onMouseLeave = () => setIsDragging(false);
  const onMouseUp = () => setIsDragging(false);
  const onMouseMove = (e) => {
    if (!isDragging || !scrollRef.current) return;
    e.preventDefault();
    const x = e.pageX - scrollRef.current.offsetLeft;
    const walk = (x - startX) * 1.5;
    scrollRef.current.scrollLeft = scrollLeft - walk;
  };

  return (
    <div
      ref={scrollRef}
      onMouseDown={onMouseDown}
      onMouseLeave={onMouseLeave}
      onMouseUp={onMouseUp}
      onMouseMove={onMouseMove}
      style={{ overflowX: "auto", cursor: isDragging ? "grabbing" : "auto", ...style }}
      className={className}
    >
      {children}
    </div>
  );
};

// ─── INITIAL DATA ─────────────────────────────────────────────────────────────
const INIT = {
  basics: {
    legalName: "", tradeName: "",
    address: "",
    email: "", contact: "", promoters: "",
    startDateP1: "", startDateP2: "",
    description: "",
    pitchDeck: "", burningDesire: "",
  },
  revP1: [
    {
      id: "1", header: "", items: [
        { id: "1a", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "1b", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "1c", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "1d", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "1e", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
      ]
    },
  ],
  revP2: [
    {
      id: "1", header: "", items: [
        { id: "1a", sub: "", qtyDay: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "1b", sub: "", qtyDay: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
      ]
    },
  ],
  opexP1: [
    {
      id: "1", header: "", items: [
        { id: "1a", sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "1b", sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "1c", sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
      ]
    },
  ],
  capex: [
    {
      sno: 1, category: "", items: [
        { name: "", total: 0, y1: 0, y2: 0, y3: 0, y4: 0, y5: 0 },
      ]
    },
  ],
  totalProjectCost: { total: 0, promoterContrib: 0, termLoan: 0, wcLoan: 0 },
  loan1: { amount: 0, duration: 0, rate: 0, startDate: "2026-04-01" },
  loan2: { amount: 0, duration: 0, rate: 0, startDate: "2027-04-01" },
  fixedAssets: [
    { name: "", rate: 0.10, opening: 0, addAbove: 0, addBelow: 0 },
    { name: "", rate: 0.15, opening: 0, addAbove: 0, addBelow: 0 },
    { name: "", rate: 0.40, opening: 0, addAbove: 0, addBelow: 0 },
  ],
};


const zeroGroupItems = (groups, isOpex = false, perDay = false) =>
  (groups || []).map((g) => ({
    ...g,
    header: "",
    items: (g.items || []).map((it) => ({
      ...it,
      sub: "",
      qty: isOpex ? 0 : 0,
      qtyDay: perDay ? 0 : it.qtyDay,
      price: 0,
      cost: 0,
      gY1: 0,
      gY2: 0,
      gY3: 0,
      gY4: 0,
      gY5: 0,
    })),
  }));

const buildEmptyState = () => ({
  ...INIT,
  basics: {
    ...INIT.basics,
    legalName: "",
    tradeName: "",
    address: "",
    email: "",
    contact: "",
    promoters: "",
    startDateP1: "",
    startDateP2: "",
    description: "",
    pitchDeck: "",
    burningDesire: "",
  },
  revP1: zeroGroupItems(INIT.revP1, false, false),
  revP2: zeroGroupItems(INIT.revP2, false, true),
  opexP1: zeroGroupItems(INIT.opexP1, true, false),
  capex: (INIT.capex || []).map((c) => ({
    ...c,
    items: (c.items || []).map((it) => ({
      ...it,
      name: "",
      total: 0,
      y1: 0,
      y2: 0,
      y3: 0,
      y4: 0,
      y5: 0,
    })),
  })),
  totalProjectCost: { total: 0, promoterContrib: 0, termLoan: 0, wcLoan: 0 },
  loan1: { ...INIT.loan1, amount: 0, duration: 0, rate: 0, startDate: "" },
  loan2: { ...INIT.loan2, amount: 0, duration: 0, rate: 0, startDate: "" },
  fixedAssets: (INIT.fixedAssets || []).map((fa) => ({
    ...fa,
    name: "",
    opening: 0,
    addAbove: 0,
    addBelow: 0,
    rate: 0,
  })),
});

// ─── CALCULATIONS ─────────────────────────────────────────────────────────────
function calcRevYearly(groups, perDay = false) {
  return groups.map(g => ({
    ...g,
    yearlyTotals: YEARS.map((_, yi) => {
      return g.items.reduce((sum, item) => {
        const base = perDay ? (item.qtyDay || 0) * (item.price || 0) * 30 * 12 : (item.qty || 0) * (item.price || 0) * 12;
        if (!base) return sum;
        let v = base;
        for (let i = 0; i < yi; i++) v *= (1 + (item[GROWTH_KEYS_REV[i]] || 0));
        return sum + v;
      }, 0);
    }),
    items: g.items.map(item => ({
      ...item,
      yearlyTotals: YEARS.map((_, yi) => {
        const base = perDay ? (item.qtyDay || 0) * (item.price || 0) * 30 * 12 : (item.qty || 0) * (item.price || 0) * 12;
        if (!base) return 0;
        let v = base;
        for (let i = 0; i < yi; i++) v *= (1 + (item[GROWTH_KEYS_REV[i]] || 0));
        return v;
      })
    }))
  }));
}

function calcOpexYearly(groups) {
  return groups.map(g => ({
    ...g,
    yearlyTotals: YEARS.map((_, yi) => {
      return g.items.reduce((sum, item) => {
        const base = (item.qty || 1) * (item.cost || 0) * 12;
        if (!base) return sum;
        let v = base;
        for (let i = 0; i < yi; i++) v *= (1 + (item[GROWTH_KEYS_OPEX[i]] || 0));
        return sum + v;
      }, 0);
    }),
    items: g.items.map(item => ({
      ...item,
      yearlyTotals: YEARS.map((_, yi) => {
        const base = (item.qty || 1) * (item.cost || 0) * 12;
        if (!base) return 0;
        let v = base;
        for (let i = 0; i < yi; i++) v *= (1 + (item[GROWTH_KEYS_OPEX[i]] || 0));
        return v;
      })
    }))
  }));
}

function calcLoan(amount, durationMonths, ratePA, startDate) {
  const monthly = ratePA / 12 / 100;
  const emi = amount * monthly * Math.pow(1 + monthly, durationMonths) / (Math.pow(1 + monthly, durationMonths) - 1);
  const rows = [];
  let bal = amount;
  const sd = new Date(startDate);
  for (let i = 0; i < durationMonths; i++) {
    const date = new Date(sd.getFullYear(), sd.getMonth() + i + 1, 1);
    const interest = bal * monthly;
    const principal = emi - interest;
    const closing = bal - principal;
    rows.push({ no: i + 1, date: date.toLocaleDateString("en-IN"), opening: bal, principal, interest, closing: Math.max(0, closing), emi });
    bal = Math.max(0, closing);
  }
  return { emi, totalInterest: emi * durationMonths - amount, rows };
}

function calcFA(assets) {
  return assets.map(a => {
    const dep1 = (a.opening + a.addAbove + a.addBelow * 0.5) * a.rate;
    const closing1 = a.opening + a.addAbove + a.addBelow - dep1;
    const dep2 = closing1 * a.rate;
    const closing2 = closing1 - dep2;
    return { ...a, dep1, closing1, dep2, closing2 };
  });
}

// ─── EDITABLE CELL ────────────────────────────────────────────────────────────
function EC({ v, onChange, type = "num", style = {}, onFocus = null }) {
  const [ed, setEd] = useState(false);
  const [val, setVal] = useState("");
  const commit = () => {
    setEd(false);
    const n = parseFloat(val.replace(/[₹,%()CrLK]/g, "").replace(/,/g, "")) || 0;
    onChange(type === "pct" ? n / 100 : n);
  };
  const display = type === "pct" ? fmtPct(v) : type === "currency" ? fmtINR(v, true) : fmtNum(v);
  if (ed) return <input autoFocus value={val} onChange={e => setVal(e.target.value)} onBlur={commit} onFocus={() => onFocus && onFocus()} onKeyDown={e => { if (e.key === "Enter") commit(); if (e.key === "Escape") setEd(false); }} style={{ background: "#0A2040", border: `1px solid ${C.teal}`, borderRadius: 3, color: C.inputBlue, fontSize: 14, fontFamily: "monospace", padding: "2px 6px", textAlign: "right", width: "100%", outline: "none", boxSizing: "border-box", ...style }} />;
  return <div onClick={() => { setEd(true); setVal(type === "pct" ? ((v || 0) * 100).toFixed(1) : String(v || 0)); if (onFocus) onFocus(); }} title="Click to edit" style={{ color: C.inputBlue, fontSize: 14, fontFamily: "monospace", textAlign: "right", cursor: "pointer", padding: "2px 6px", borderRadius: 3, border: "1px solid transparent", ...style }}>{display}</div>;
}

function TI({ v, onChange, placeholder = "", style = {}, onFocus = null }) {
  return <input value={v || ""} onChange={e => onChange(e.target.value)} onFocus={() => onFocus && onFocus()} placeholder={placeholder} style={{ background: "transparent", border: "none", outline: "none", color: C.inputBlue, fontSize: 14, fontFamily: "sans-serif", width: "100%", ...style }} />;
}

// ─── SHARED TABLE STYLES ──────────────────────────────────────────────────────
const th = (extra = {}) => ({ padding: "8px 12px", fontSize: 13, color: C.gold, fontWeight: 700, letterSpacing: "0.05em", textTransform: "uppercase", borderBottom: `2px solid ${C.goldD}`, background: C.headerBg, whiteSpace: "nowrap", ...extra });
const td0 = (extra = {}) => ({ padding: "6px 12px", fontSize: 14, borderBottom: `1px solid ${C.border}`, ...extra });

// ─── SHEET: 1. BASICS ────────────────────────────────────────────────────────
function Basics({ d, setD, onFocus }) {
  const fields = [
    ["1", "Legal Name of the Business", "legalName"],
    ["2", "Trade Name of the Business", "tradeName"],
    ["3", "Registered Office Address", "address"],
    ["4", "Official Email Id", "email"],
    ["5", "Official Contact Number", "contact"],
    ["6.a", "Total number of Promoters", "promoters"],
    ["7", "Tentative Start Date of Phase 1", "startDateP1"],
    ["8", "Company Description", "description"],
    ["9", "Link to Company Pitch Deck", "pitchDeck"],
    ["10", "Burning Desire of the company", "burningDesire"],
    ["11", "Tentative Start Date of Phase 2", "startDateP2"],
  ];
  return (
    <div>
      <SheetHeader title="1. Basic Information" sub="Company Details & Project Basics" />
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <tbody>
          {fields.map(([no, label, key], i) => (
            <tr key={key} style={{ background: i % 2 === 0 ? C.bg2 : C.bg1 }}>
              <td style={{ ...td0(), width: 50, color: C.gold, fontFamily: "monospace", fontWeight: 700, borderRight: `1px solid ${C.border}` }}>{no}</td>
              <td style={{ ...td0(), width: 300, color: C.text1, borderRight: `1px solid ${C.border}` }}>{label}</td>
              <td style={td0()}>
                <TI
                  v={d[key]}
                  onChange={v => setD(p => ({ ...p, [key]: v }))}
                  onFocus={() => onFocus && onFocus("Basics", label, key, d[key])}
                  placeholder={`Enter ${label.toLowerCase()}...`}
                  style={{ color: C.inputBlue, fontSize: 12, padding: "2px 4px" }}
                />
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ─── SHEET: A. DATA NEEDED ───────────────────────────────────────────────────
function DataNeeded({ setSheet }) {
  const items = [
    ["I", "Revenue Streams", "Define all revenue streams with qty, price, growth rates", "A.I Revenue Streams - P1"],
    ["II", "Sales Price per Revenue Streams", "Set pricing for each sub-service", "A.I Revenue Streams - P1"],
    ["III", "Costing headers (OPEX)", "List all operating expense categories and costs", "A.IIOPEX"],
    ["IIV", "Capex Information", "Define capital expenditure by category and year", "A.III CAPEX"],
  ];
  return (
    <div>
      <SheetHeader title="A. Data Needed" sub="Major heads of data required for the business plan" />
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr><th style={th()}>No.</th><th style={th({ textAlign: "left" })}>Data Category</th><th style={th({ textAlign: "left" })}>Description</th><th style={th()}>Go To Sheet</th></tr>
        </thead>
        <tbody>
          {items.map(([no, cat, desc, link], i) => (
            <tr key={no} style={{ background: i % 2 === 0 ? C.bg2 : C.bg1 }}>
              <td style={{ ...td0(), color: C.gold, fontFamily: "monospace", fontWeight: 700, textAlign: "center" }}>{no}</td>
              <td style={{ ...td0(), color: C.text0, fontWeight: 600 }}>{cat}</td>
              <td style={{ ...td0(), color: C.text1 }}>{desc}</td>
              <td style={{ ...td0(), textAlign: "center" }}>
                <button onClick={() => setSheet(link)} style={{ padding: "3px 12px", background: C.navB, border: `1px solid ${C.border}`, borderRadius: 4, color: C.teal, cursor: "pointer", fontSize: 11 }}>Open →</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ─── SHEET: REVENUE STREAMS ──────────────────────────────────────────────────
function RevStreams({ groups, setGroups, phase, perDay = false, onFocus }) {
  const computed = calcRevYearly(groups, perDay);
  const gKeys = ["gY1", "gY2", "gY3", "gY4", "gY5"];
  const gLabels = ["Monthly Growth Y1", "Yearly Growth Y2→Y3", "Growth Y3→Y4", "Growth Y4→Y5", "Growth Y5→Y6"];
  const upd = (gi, ii, k, v) => setGroups(p => p.map((g, gi2) => gi2 !== gi ? g : { ...g, items: g.items.map((it, ii2) => ii2 !== ii ? it : { ...it, [k]: v }) }));
  const updG = (gi, k, v) => setGroups(p => p.map((g, gi2) => gi2 !== gi ? g : { ...g, [k]: v }));
  const grandTotals = YEARS.map((_, yi) => computed.reduce((s, g) => s + g.yearlyTotals[yi], 0));

  return (
    <DraggableScroll>
      <SheetHeader title={`A.I Revenue Streams — ${phase}`} sub={`This sheet assists in getting data related to Sales (${perDay ? "qty per day" : "qty per month"})`} />
      <table style={{ borderCollapse: "collapse", minWidth: 1200 }}>
        <thead>
          <tr>
            {["#", "Major Header", "Sub Service", perDay ? "Qty/Day" : "Qty/Month", "Price (₹)", ...YEARS, ...gLabels].map((h, i) => (
              <th key={i} style={th({ textAlign: i >= 3 ? "right" : "left", minWidth: i === 0 ? 40 : i <= 2 ? 150 : 90 })}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {computed.map((g, gi) => [
            <tr key={`h${gi}`} style={{ background: C.sectionBg }}>
              <td style={{ ...td0(), color: C.gold, fontFamily: "monospace", fontWeight: 700, borderRight: `1px solid ${C.border}` }}>{g.id}</td>
              <td colSpan={2} style={td0()}><TI v={g.header} onChange={v => updG(gi, "header", v)} placeholder="Category name..." style={{ color: C.text0, fontWeight: 700, fontSize: 12 }} /></td>
              <td style={td0()} />
              <td style={td0()} />
              {YEARS.map((_, yi) => <td key={yi} style={{ ...td0(), color: C.tealL, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(g.yearlyTotals[yi], true)}</td>)}
              {gLabels.map((_, i) => <td key={i} style={td0()} />)}
            </tr>,
            ...g.items.map((it, ii) => (
              <tr key={`${gi}-${ii}`} style={{ background: ii % 2 === 0 ? C.bg1 : C.bg0 }}>
                <td style={{ ...td0(), color: C.text2, fontFamily: "monospace", fontSize: 10 }}>{it.id}</td>
                <td style={td0()} />
                <td style={td0()}>
                  <TI
                    v={it.sub}
                    onChange={v => upd(gi, ii, "sub", v)}
                    onFocus={() => onFocus && onFocus(phase, it.sub, "Sub Service", it.sub)}
                    placeholder="Sub service..."
                    style={{ color: C.text1 }}
                  />
                </td>
                <td style={td0()}>
                  <EC
                    v={perDay ? it.qtyDay : it.qty}
                    type="num"
                    onChange={v => upd(gi, ii, perDay ? "qtyDay" : "qty", v)}
                    onFocus={() => onFocus && onFocus(phase, it.sub, perDay ? "Qty/Day" : "Qty/Month", perDay ? it.qtyDay : it.qty)}
                  />
                </td>
                <td style={td0()}>
                  <EC
                    v={it.price}
                    type="num"
                    onChange={v => upd(gi, ii, "price", v)}
                    onFocus={() => onFocus && onFocus(phase, it.sub, "Price", it.price)}
                  />
                </td>
                {YEARS.map((_, yi) => <td key={yi} style={{ ...td0(), color: it.yearlyTotals[yi] > 0 ? C.text0 : C.text2, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(it.yearlyTotals[yi], true)}</td>)}
                {gKeys.map((k, i) => (
                  <td key={i} style={td0()}>
                    <EC
                      v={it[k]}
                      type="pct"
                      onChange={v => upd(gi, ii, k, v)}
                      onFocus={() => onFocus && onFocus(phase, it.sub, gLabels[i], it[k])}
                    />
                  </td>
                ))}
              </tr>
            )),
            <tr key={`gt${gi}`} style={{ background: C.totalBg }}>
              <td colSpan={3} style={{ ...td0({ borderTop: `1px solid ${C.borderLight}` }), color: C.gold, fontWeight: 700, fontSize: 10, letterSpacing: "0.04em" }}>GRAND TOTAL</td>
              <td colSpan={2} style={{ borderBottom: `1px solid ${C.border}`, borderTop: `1px solid ${C.borderLight}` }} />
              {YEARS.map((_, yi) => <td key={yi} style={{ ...td0({ borderTop: `1px solid ${C.borderLight}` }), color: C.gold, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(g.yearlyTotals[yi], true)}</td>)}
              {gLabels.map((_, i) => <td key={i} style={{ borderBottom: `1px solid ${C.border}`, borderTop: `1px solid ${C.borderLight}` }} />)}
            </tr>,
            <tr key={`sp${gi}`}><td colSpan={20} style={{ height: 8 }} /></tr>
          ])}
          <tr style={{ background: "#07101E" }}>
            <td colSpan={5} style={{ padding: "8px 10px", color: C.goldL, fontWeight: 700, fontSize: 12, borderTop: `2px solid ${C.gold}`, borderBottom: `2px solid ${C.gold}` }}>TOTAL REVENUE (Annual)</td>
            {grandTotals.map((t, yi) => <td key={yi} style={{ padding: "8px 10px", color: C.goldL, fontFamily: "monospace", textAlign: "right", fontWeight: 700, fontSize: 12, borderTop: `2px solid ${C.gold}`, borderBottom: `2px solid ${C.gold}` }}>{fmtINR(t, true)}</td>)}
            {gLabels.map((_, i) => <td key={i} style={{ borderTop: `2px solid ${C.gold}`, borderBottom: `2px solid ${C.gold}` }} />)}
          </tr>
        </tbody>
      </table>
    </DraggableScroll>
  );
}

// ─── SHEET: OPEX ─────────────────────────────────────────────────────────────
function OpexSheet({ groups, setGroups, onFocus }) {
  const computed = calcOpexYearly(groups);
  const gKeys = ["gY1", "gY2", "gY3", "gY4", "gY5"];
  const gLabels = ["Growth Y1", "Growth Y2", "Growth Y3", "Growth Y4", "Growth Y5"];
  const upd = (gi, ii, k, v) => setGroups(p => p.map((g, gi2) => gi2 !== gi ? g : { ...g, items: g.items.map((it, ii2) => ii2 !== ii ? it : { ...it, [k]: v }) }));
  const updG = (gi, k, v) => setGroups(p => p.map((g, gi2) => gi2 !== gi ? g : { ...g, [k]: v }));
  const totals = YEARS.map((_, yi) => computed.reduce((s, g) => s + g.yearlyTotals[yi], 0));

  return (
    <DraggableScroll>
      <SheetHeader title="A.II OPEX — Operating Expenses" sub="Write down all major and minor headers for costing — quantity and price beside each header" />
      <table style={{ borderCollapse: "collapse", minWidth: 1200 }}>
        <thead>
          <tr>
            {["#", "Major Head", "Sub Service", "Qty", "Cost/Month (₹)", ...YEARS, ...gLabels].map((h, i) => (
              <th key={i} style={th({ textAlign: i >= 3 ? "right" : "left", minWidth: i === 0 ? 40 : i <= 2 ? 180 : 90 })}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {computed.map((g, gi) => [
            <tr key={`h${gi}`} style={{ background: C.sectionBg }}>
              <td style={{ ...td0(), color: C.gold, fontFamily: "monospace", fontWeight: 700 }}>{g.id}</td>
              <td colSpan={2} style={td0()}><TI v={g.header} onChange={v => updG(gi, "header", v)} placeholder="Category name..." style={{ color: C.text0, fontWeight: 700, fontSize: 12 }} /></td>
              {YEARS.map((_, yi) => <td key={yi} style={{ ...td0(), color: C.redL, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(g.yearlyTotals[yi], true)}</td>)}
              {gLabels.map((_, i) => <td key={i} style={td0()} />)}
            </tr>,
            ...g.items.map((it, ii) => (
              <tr key={`${gi}-${ii}`} style={{ background: ii % 2 === 0 ? C.bg1 : C.bg0 }}>
                <td style={{ ...td0(), color: C.text2, fontFamily: "monospace", fontSize: 10 }}>{it.id}</td>
                <td style={td0()} />
                <td style={td0()}>
                  <TI
                    v={it.sub}
                    onChange={v => upd(gi, ii, "sub", v)}
                    onFocus={() => onFocus && onFocus("OPEX", it.sub, "Sub Service", it.sub)}
                    placeholder="Sub service..."
                    style={{ color: C.text1 }}
                  />
                </td>
                <td style={td0()}>
                  <EC
                    v={it.qty}
                    type="num"
                    onChange={v => upd(gi, ii, "qty", v)}
                    onFocus={() => onFocus && onFocus("OPEX", it.sub, "Quantity", it.qty)}
                  />
                </td>
                <td style={td0()}>
                  <EC
                    v={it.cost}
                    type="num"
                    onChange={v => upd(gi, ii, "cost", v)}
                    onFocus={() => onFocus && onFocus("OPEX", it.sub, "Cost/Month", it.cost)}
                  />
                </td>
                {YEARS.map((_, yi) => <td key={yi} style={{ ...td0(), color: it.yearlyTotals[yi] > 0 ? "#E87878" : C.text2, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(it.yearlyTotals[yi], true)}</td>)}
                {gKeys.map((k, i) => (
                  <td key={i} style={td0()}>
                    <EC
                      v={it[k]}
                      type="pct"
                      onChange={v => upd(gi, ii, k, v)}
                      onFocus={() => onFocus && onFocus("OPEX", it.sub, gLabels[i], it[k])}
                    />
                  </td>
                ))}
              </tr>
            )),
            <tr key={`gt${gi}`} style={{ background: C.totalBg }}>
              <td colSpan={3} style={{ ...td0({ borderTop: `1px solid ${C.borderLight}` }), color: C.gold, fontWeight: 700, fontSize: 10 }}>GRAND TOTAL</td>
              <td colSpan={2} style={{ borderBottom: `1px solid ${C.border}`, borderTop: `1px solid ${C.borderLight}` }} />
              {YEARS.map((_, yi) => <td key={yi} style={{ ...td0({ borderTop: `1px solid ${C.borderLight}` }), color: C.gold, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(g.yearlyTotals[yi], true)}</td>)}
              {gLabels.map((_, i) => <td key={i} style={{ borderBottom: `1px solid ${C.border}`, borderTop: `1px solid ${C.borderLight}` }} />)}
            </tr>,
            <tr key={`sp${gi}`}><td colSpan={20} style={{ height: 8 }} /></tr>
          ])}
          <tr style={{ background: "#150808" }}>
            <td colSpan={5} style={{ padding: "8px 10px", color: "#FF9999", fontWeight: 700, fontSize: 12, borderTop: `2px solid ${C.red}`, borderBottom: `2px solid ${C.red}` }}>TOTAL OPEX (Annual)</td>
            {totals.map((t, yi) => <td key={yi} style={{ padding: "8px 10px", color: "#FF9999", fontFamily: "monospace", textAlign: "right", fontWeight: 700, fontSize: 12, borderTop: `2px solid ${C.red}`, borderBottom: `2px solid ${C.red}` }}>{fmtINR(t, true)}</td>)}
            {gLabels.map((_, i) => <td key={i} style={{ borderTop: `2px solid ${C.red}`, borderBottom: `2px solid ${C.red}` }} />)}
          </tr>
        </tbody>
      </table>
    </DraggableScroll>
  );
}

// ─── SHEET: CAPEX ─────────────────────────────────────────────────────────────
function Capex({ data, setData, onFocus }) {
  const updItem = (ci, ii, k, v) => setData(p => p.map((cat, ci2) => ci2 !== ci ? cat : { ...cat, items: cat.items.map((it, ii2) => ii2 !== ii ? it : { ...it, [k]: v }) }));
  const updCat = (ci, k, v) => setData(p => p.map((cat, ci2) => ci2 !== ci ? cat : { ...cat, [k]: v }));
  const yLabels = ["Year 1 (31/03/2027)", "Year 2 (31/03/2028)", "Year 3 (31/03/2029)", "Year 4 (31/03/2030)", "Year 5 (31/03/2031)"];
  return (
    <DraggableScroll>
      <SheetHeader title="A.III CAPEX — Capital Expenditures" sub="List all long-term assets and one-time setup costs" />
      <table style={{ borderCollapse: "collapse", minWidth: 900 }}>
        <thead>
          <tr>
            {["S.No", "Nature of Expense", "Total Amount (₹)", ...yLabels].map((h, i) => (
              <th key={i} style={th({ textAlign: i >= 2 ? "right" : "left", minWidth: i === 1 ? 200 : 90 })}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.map((cat, ci) => [
            <tr key={`c${ci}`} style={{ background: C.sectionBg }}>
              <td style={{ ...td0(), color: C.gold, fontFamily: "monospace", fontWeight: 700 }}>{cat.sno}</td>
              <td style={td0()}><TI v={cat.category} onChange={v => updCat(ci, "category", v)} placeholder="Category..." style={{ color: C.text0, fontWeight: 700, fontSize: 12 }} /></td>
              <td style={{ ...td0(), color: C.tealL, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(cat.items.reduce((s, it) => s + (it.total || 0), 0), true)}</td>
              {["y1", "y2", "y3", "y4", "y5"].map(k => <td key={k} style={{ ...td0(), color: C.tealL, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(cat.items.reduce((s, it) => s + (it[k] || 0), 0), true)}</td>)}
            </tr>,
            ...cat.items.map((it, ii) => (
              <tr key={`${ci}-${ii}`} style={{ background: ii % 2 === 0 ? C.bg1 : C.bg0 }}>
                <td style={{ ...td0(), color: C.text2, fontSize: 10 }} />
                <td style={{ ...td0(), paddingLeft: 24 }}><TI v={it.name} onChange={v => updItem(ci, ii, "name", v)} onFocus={() => onFocus && onFocus("CAPEX", it.name, "Asset Name", it.name)} placeholder="Asset name..." style={{ color: C.text1 }} /></td>
                <td style={td0()}><EC v={it.total} type="num" onChange={v => updItem(ci, ii, "total", v)} onFocus={() => onFocus && onFocus("CAPEX", it.name, "Total Amount", it.total)} /></td>
                {["y1", "y2", "y3", "y4", "y5"].map(k => <td key={k} style={td0()}><EC v={it[k]} type="num" onChange={v => updItem(ci, ii, k, v)} onFocus={() => onFocus && onFocus("CAPEX", it.name, `Year ${k.slice(1)}`, it[k])} /></td>)}
              </tr>
            )),
            <tr key={`sp${ci}`}><td colSpan={8} style={{ height: 8 }} /></tr>
          ])}
        </tbody>
      </table>
    </DraggableScroll>
  );
}

// ─── SHEET: B.I SALES (monthly view) ─────────────────────────────────────────
function SalesSheet({ groups, phase }) {
  const computed = calcRevYearly(groups, false);
  return (
    <DraggableScroll>
      <SheetHeader title={`B.I Sales — ${phase} Monthly View (Year 1: 2026-27)`} sub="Do not enter any data in this sheet — calculated from Revenue Streams assumptions" readOnly />
      <table style={{ borderCollapse: "collapse", minWidth: 1400 }}>
        <thead>
          <tr style={{ background: C.headerBg }}>
            <th style={th({ minWidth: 40 })}>#</th>
            <th style={th({ textAlign: "left", minWidth: 160 })}>Services</th>
            <th style={th({ textAlign: "left", minWidth: 150 })}>Sub Services</th>
            {MONTHS_Y1.map(m => <th key={m} style={th({ minWidth: 90 })}>{m}</th>)}
            <th style={{ ...th({ minWidth: 110 }), borderLeft: `2px solid ${C.gold}` }}>Annual</th>
          </tr>
          <tr style={{ background: C.headerBg }}>
            <td colSpan={3} style={{ padding: "4px 10px", fontSize: 10, color: C.text2 }}>Phase-I Revenues</td>
            {MONTHS_Y1_DATES.map((d, i) => <td key={i} style={{ padding: "3px 8px", fontSize: 9, color: C.text2, textAlign: "right" }}>{d}</td>)}
            <td style={{ borderLeft: `2px solid ${C.gold}` }} />
          </tr>
        </thead>
        <tbody>
          {computed.map((g, gi) => {
            const annualTotal = g.yearlyTotals[0];
            const monthlyBase = annualTotal / 12;
            return [
              <tr key={`h${gi}`} style={{ background: C.sectionBg }}>
                <td style={{ ...td0(), color: C.gold, fontFamily: "monospace", fontWeight: 700 }}>{g.id}</td>
                <td style={{ ...td0(), color: C.text0, fontWeight: 700 }} colSpan={2}>{g.header || `Service ${g.id}`}</td>
                {MONTHS_Y1.map((_, mi) => <td key={mi} style={{ ...td0(), color: C.tealL, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(monthlyBase * (1 + mi * 0.01), true)}</td>)}
                <td style={{ ...td0({ borderLeft: `2px solid ${C.gold}` }), color: C.tealL, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(annualTotal, true)}</td>
              </tr>,
              ...g.items.filter(it => it.sub || it.qty > 0).map((it, ii) => (
                <tr key={`${gi}-${ii}`} style={{ background: ii % 2 === 0 ? C.bg1 : C.bg0 }}>
                  <td style={{ ...td0(), color: C.text2, fontSize: 10, fontFamily: "monospace" }}>{it.id}</td>
                  <td style={td0()} />
                  <td style={{ ...td0(), color: C.text1 }}>{it.sub || "—"}</td>
                  {MONTHS_Y1.map((_, mi) => {
                    const mv = it.yearlyTotals[0] / 12 * (1 + mi * 0.01);
                    return <td key={mi} style={{ ...td0(), color: it.qty > 0 ? C.text0 : C.text2, fontFamily: "monospace", textAlign: "right" }}>{it.qty > 0 ? fmtINR(mv, true) : "—"}</td>;
                  })}
                  <td style={{ ...td0({ borderLeft: `2px solid ${C.gold}` }), color: C.text0, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(it.yearlyTotals[0], true)}</td>
                </tr>
              )),
              <tr key={`gt${gi}`} style={{ background: C.totalBg }}>
                <td colSpan={3} style={{ ...td0({ borderTop: `1px solid ${C.borderLight}` }), color: C.gold, fontWeight: 700, fontSize: 10 }}>Grand Total</td>
                {MONTHS_Y1.map((_, mi) => <td key={mi} style={{ ...td0({ borderTop: `1px solid ${C.borderLight}` }), color: C.gold, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(monthlyBase * (1 + mi * 0.01), true)}</td>)}
                <td style={{ ...td0({ borderTop: `1px solid ${C.borderLight}`, borderLeft: `2px solid ${C.gold}` }), color: C.gold, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(annualTotal, true)}</td>
              </tr>,
              <tr key={`sp${gi}`}><td colSpan={20} style={{ height: 8 }} /></tr>
            ];
          })}
          <tr style={{ background: "#07101E" }}>
            <td colSpan={3} style={{ padding: "8px 10px", color: C.goldL, fontWeight: 700, fontSize: 12, borderTop: `2px solid ${C.gold}`, borderBottom: `2px solid ${C.gold}` }}>Yearly Grand Total</td>
            {MONTHS_Y1.map((_, mi) => {
              const t = computed.reduce((s, g) => s + g.yearlyTotals[0] / 12 * (1 + mi * 0.01), 0);
              return <td key={mi} style={{ padding: "8px 10px", color: C.goldL, fontFamily: "monospace", textAlign: "right", fontWeight: 700, borderTop: `2px solid ${C.gold}`, borderBottom: `2px solid ${C.gold}` }}>{fmtINR(t, true)}</td>;
            })}
            <td style={{ padding: "8px 10px", color: C.goldL, fontFamily: "monospace", textAlign: "right", fontWeight: 700, borderTop: `2px solid ${C.gold}`, borderBottom: `2px solid ${C.gold}`, borderLeft: `2px solid ${C.gold}` }}>
              {fmtINR(computed.reduce((s, g) => s + g.yearlyTotals[0], 0), true)}
            </td>
          </tr>
        </tbody>
      </table>
    </DraggableScroll>
  );
}

// ─── SHEET: TOTAL PROJECT COST ────────────────────────────────────────────────
function TotalProjectCost({ data, setData, onFocus }) {
  const { total, promoterContrib, termLoan, wcLoan } = data;
  return (
    <div>
      <SheetHeader title="2. Statement of Total Project Cost" sub="Project Summary" />
      <table style={{ borderCollapse: "collapse", width: 600 }}>
        <thead>
          <tr><th style={th({ textAlign: "left", minWidth: 300 })}>Description</th><th style={th({ minWidth: 160 })}>Amount (₹)</th></tr>
        </thead>
        <tbody>
          {[
            ["Total", "total"], ["Less: 20% Promoter Contribution", "promoterContrib"],
            ["Balance Term Loan Application", "termLoan"], ["Working Capital Loan", "wcLoan"]
          ].map(([label, key], i) => (
            <tr key={key} style={{ background: i % 2 === 0 ? C.bg2 : C.bg1 }}>
              <td style={{ ...td0(), color: i === 0 ? C.text0 : C.text1, fontWeight: i === 0 ? 700 : 400 }}>{label}</td>
              <td style={td0()}><EC v={data[key]} type="num" onChange={v => setData(p => ({ ...p, [key]: v }))} onFocus={() => onFocus && onFocus("Project Cost", label, key, data[key])} /></td>
            </tr>
          ))}
          <tr style={{ background: C.totalBg }}>
            <td style={{ ...td0({ borderTop: `1px solid ${C.borderLight}` }), color: C.gold, fontWeight: 700 }}>Total Loan from Bank</td>
            <td style={{ ...td0({ borderTop: `1px solid ${C.borderLight}` }), color: C.gold, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(termLoan + wcLoan)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// ─── SHEET: P&L ──────────────────────────────────────────────────────────────
function PL({ revP1, opexP1, loan1, loan2 }) {
  const revByYear = YEARS.map((_, yi) => calcRevYearly(revP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const opexByYear = YEARS.map((_, yi) => calcOpexYearly(opexP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const grossProfit = YEARS.map((_, yi) => revByYear[yi] - opexByYear[yi] * 0.3);
  const ebitda = YEARS.map((_, yi) => revByYear[yi] - opexByYear[yi]);
  const deprn = YEARS.map(() => 75000 * 12);
  const ebit = YEARS.map((_, yi) => ebitda[yi] - deprn[yi]);
  
  // Dynamic interest calculation: Y2 for Loan 1, Y3 for Loan 2 (matching export logic)
  const interest = YEARS.map((_, yi) => {
    if (yi === 1) return (Number(loan1?.amount) || 0) * (Number(loan1?.rate) || 0) / 100;
    if (yi === 2) return (Number(loan2?.amount) || 0) * (Number(loan2?.rate) || 0) / 100;
    return 0;
  });

  const pbt = YEARS.map((_, yi) => ebit[yi] - interest[yi]);
  const tax = YEARS.map((_, yi) => Math.max(0, pbt[yi] * 0.25));
  const pat = YEARS.map((_, yi) => pbt[yi] - tax[yi]);

  const rows = [
    { label: "REVENUE", type: "header" },
    { label: "Total Revenue", vals: revByYear, type: "total", color: C.tealL },
    { type: "spacer" },
    { label: "EXPENSES", type: "header" },
    { label: "Direct Costs (30% of OPEX)", vals: YEARS.map((_, yi) => -opexByYear[yi] * 0.3), type: "data" },
    { label: "Gross Profit", vals: grossProfit, type: "subtotal" },
    { label: "Gross Margin %", vals: YEARS.map((_, yi) => revByYear[yi] > 0 ? grossProfit[yi] / revByYear[yi] : 0), type: "pct" },
    { type: "spacer" },
    { label: "OPERATING EXPENSES", type: "header" },
    { label: "Total OPEX", vals: YEARS.map((_, yi) => -opexByYear[yi]), type: "data" },
    { label: "EBITDA", vals: ebitda, type: "subtotal" },
    { label: "EBITDA Margin %", vals: YEARS.map((_, yi) => revByYear[yi] > 0 ? ebitda[yi] / revByYear[yi] : 0), type: "pct" },
    { type: "spacer" },
    { label: "Depreciation & Amortisation", vals: YEARS.map(() => -75000 * 12), type: "data" },
    { label: "EBIT", vals: ebit, type: "subtotal" },
    { type: "spacer" },
    { label: "Interest & Finance Charges", vals: interest.map(v => -v), type: "data" },
    { label: "Profit Before Tax (PBT)", vals: pbt, type: "subtotal" },
    { type: "spacer" },
    { label: "Tax (25%)", vals: tax.map(v => -v), type: "data" },
    { label: "NET PROFIT AFTER TAX (PAT)", vals: pat, type: "total", color: null },
    { label: "PAT Margin %", vals: YEARS.map((_, yi) => revByYear[yi] > 0 ? pat[yi] / revByYear[yi] : 0), type: "pct" },
  ];

  return (
    <DraggableScroll>
      <SheetHeader title="4. Profit & Loss Statement" sub="Projected Annual P&L" readOnly />
      <table style={{ borderCollapse: "collapse", minWidth: 800 }}>
        <thead>
          <tr>
            <th style={th({ textAlign: "left", minWidth: 280 })}>Particulars</th>
            {YEARS.map(y => <th key={y} style={th({ minWidth: 130 })}>{y}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => {
            if (row.type === "spacer") return <tr key={i}><td colSpan={6} style={{ height: 8 }} /></tr>;
            if (row.type === "header") return (
              <tr key={i} style={{ background: C.sectionBg }}>
                <td colSpan={6} style={{ padding: "7px 10px", color: C.gold, fontSize: 10, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase" }}>{row.label}</td>
              </tr>
            );
            if (row.type === "pct") return (
              <tr key={i} style={{ background: C.bg0 }}>
                <td style={{ ...td0({ paddingLeft: 24 }), color: C.text2, fontSize: 11, fontStyle: "italic" }}>{row.label}</td>
                {row.vals.map((v, yi) => <td key={yi} style={{ ...td0(), color: v >= 0 ? C.greenL : C.redL, fontFamily: "monospace", textAlign: "right", fontSize: 11 }}>{fmtPct(v)}</td>)}
              </tr>
            );
            const isTotal = row.type === "total" || row.type === "subtotal";
            return (
              <tr key={i} style={{ background: isTotal ? C.totalBg : i % 2 === 0 ? C.bg1 : C.bg0 }}>
                <td style={{ ...td0(isTotal ? { borderTop: `1px solid ${C.borderLight}` } : {}), color: isTotal ? C.text0 : C.text1, fontWeight: isTotal ? 700 : 400, paddingLeft: isTotal ? 10 : 22 }}>{row.label}</td>
                {row.vals.map((v, yi) => {
                  const col = row.color || (v < 0 ? C.redL : isTotal ? C.text0 : C.text1);
                  if (!row.color && isTotal && row.label.includes("PAT")) {
                    const col2 = v < 0 ? C.redL : C.greenL;
                    return <td key={yi} style={{ ...td0(isTotal ? { borderTop: `1px solid ${C.borderLight}` } : {}), color: col2, fontFamily: "monospace", textAlign: "right", fontWeight: isTotal ? 700 : 400 }}>{fmtINR(v, true)}</td>;
                  }
                  return <td key={yi} style={{ ...td0(isTotal ? { borderTop: `1px solid ${C.borderLight}` } : {}), color: col, fontFamily: "monospace", textAlign: "right", fontWeight: isTotal ? 700 : 400 }}>{fmtINR(v, true)}</td>;
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </DraggableScroll>
  );
}

// ─── SHEET: BALANCE SHEET ────────────────────────────────────────────────────
function BalanceSheet({ revP1, opexP1, loan1, loan2 }) {
  const revByYear = YEARS.map((_, yi) => calcRevYearly(revP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const opexByYear = YEARS.map((_, yi) => calcOpexYearly(opexP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  
  const pat = YEARS.map((_, yi) => {
    const ebitda = revByYear[yi] - opexByYear[yi];
    const ebit = ebitda - 75000 * 12;
    const interest = (yi === 1 ? (Number(loan1?.amount) || 0) * (Number(loan1?.rate) || 0) / 100 : yi === 2 ? (Number(loan2?.amount) || 0) * (Number(loan2?.rate) || 0) / 100 : 0);
    const pbt = ebit - interest;
    return pbt - Math.max(0, pbt * 0.25);
  });
  const retainedEarnings = YEARS.map((_, yi) => pat.slice(0, yi + 1).reduce((s, v) => s + v, 0));
  const otherCL = [3509865, 4211838, 5054205, 6065046, 7278056];
  const fixedAsset = [425000, 403750, 339521, 453865, 466535];
  const investments = [4100000, 4300000, 5700000, 5800000, 6300000];
  const secDep = [1500000, 1500000, 1500000, 1500000, 1500000];
  const curAdv = [965000, 1302750, 1758712, 1934583, 2176406];
  const otherCA = [1805430, 4262000, 4773440, 4964377, 5460815];

  const sections = [
    { label: "LIABILITIES", type: "header" },
    { label: "Shareholder Funds", type: "section" },
    { label: "Capital Account", vals: YEARS.map(() => 0) },
    { label: "Add: Net Profit / Retained Earnings", vals: retainedEarnings, color: v => v < 0 ? C.redL : C.greenL },
    { label: "Shareholder Funds", vals: retainedEarnings, type: "subtotal" },
    { type: "spacer" },
    { label: "Loans Liability", type: "section" },
    { label: "Term Loan", vals: YEARS.map(() => 0) },
    { type: "spacer" },
    { label: "Current Liabilities", type: "section" },
    { label: "Sundry Creditors", vals: YEARS.map(() => 0) },
    { label: "Other Current Liabilities", vals: otherCL },
    { label: "Total Liabilities", vals: YEARS.map((_, yi) => retainedEarnings[yi] + otherCL[yi]), type: "total" },
    { type: "spacer" },
    { label: "ASSETS", type: "header" },
    { label: "Fixed Assets", type: "section" },
    { label: "Net Fixed Assets", vals: fixedAsset },
    { type: "spacer" },
    { label: "Investments", vals: investments },
    { type: "spacer" },
    { label: "Current Assets", type: "section" },
    { label: "Security Deposits", vals: secDep },
    { label: "Current Advances", vals: curAdv },
    { label: "Sundry Debtors", vals: YEARS.map(() => 0) },
    { label: "Other Current Assets", vals: otherCA },
    { label: "Cash & Bank Balance", vals: YEARS.map((_, yi) => Math.max(0, retainedEarnings[yi] + otherCL[yi] - fixedAsset[yi] - investments[yi] - secDep[yi] - curAdv[yi] - otherCA[yi])) },
    { label: "Total Assets", vals: YEARS.map((_, yi) => fixedAsset[yi] + investments[yi] + secDep[yi] + curAdv[yi] + otherCA[yi]), type: "total" },
  ];

  return (
    <DraggableScroll>
      <SheetHeader title="5. Balance Sheet" sub="M/S AIROC Hospitals Pvt Ltd — Projected Balance Sheet" readOnly />
      <table style={{ borderCollapse: "collapse", minWidth: 800 }}>
        <thead>
          <tr>
            <th style={th({ textAlign: "left", minWidth: 280 })}>Particulars</th>
            {YEARS.map((_, yi) => <th key={yi} style={th({ minWidth: 130 })}>31st Mar {2026 + yi}</th>)}
          </tr>
        </thead>
        <tbody>
          {sections.map((row, i) => {
            if (row.type === "spacer") return <tr key={i}><td colSpan={6} style={{ height: 8 }} /></tr>;
            if (row.type === "header") return <tr key={i} style={{ background: C.sectionBg }}><td colSpan={6} style={{ padding: "7px 10px", color: C.gold, fontWeight: 700, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.08em" }}>{row.label}</td></tr>;
            if (row.type === "section") return <tr key={i} style={{ background: "#0A1528" }}><td colSpan={6} style={{ padding: "6px 10px 4px 18px", color: C.text1, fontWeight: 600, fontSize: 11, fontStyle: "italic" }}>{row.label}</td></tr>;
            const isTotal = row.type === "total" || row.type === "subtotal";
            return (
              <tr key={i} style={{ background: isTotal ? C.totalBg : i % 2 === 0 ? C.bg1 : C.bg0 }}>
                <td style={{ ...td0(isTotal ? { borderTop: `1px solid ${C.borderLight}` } : {}), color: isTotal ? C.text0 : C.text1, fontWeight: isTotal ? 700 : 400, paddingLeft: isTotal ? 10 : 24 }}>{row.label}</td>
                {row.vals.map((v, yi) => {
                  const col = row.color ? row.color(v) : (v < 0 ? C.redL : isTotal ? C.text0 : C.text1);
                  return <td key={yi} style={{ ...td0(isTotal ? { borderTop: `1px solid ${C.borderLight}` } : {}), color: col, fontFamily: "monospace", textAlign: "right", fontWeight: isTotal ? 700 : 400 }}>{fmtINR(v, true)}</td>;
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </DraggableScroll>
  );
}

// ─── SHEET: RATIOS ────────────────────────────────────────────────────────────
function Ratios({ revP1, opexP1, loan1, loan2 }) {
  const revByYear = YEARS.map((_, yi) => calcRevYearly(revP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const opexByYear = YEARS.map((_, yi) => calcOpexYearly(opexP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const pat = YEARS.map((_, yi) => {
    const ebitda = revByYear[yi] - opexByYear[yi];
    const ebit = ebitda - 75000 * 12;
    const interest = (yi === 1 ? (Number(loan1?.amount) || 0) * (Number(loan1?.rate) || 0) / 100 : yi === 2 ? (Number(loan2?.amount) || 0) * (Number(loan2?.rate) || 0) / 100 : 0);
    const pbt = ebit - interest;
    return pbt - Math.max(0, pbt * 0.25);
  });
  const fixedAssets = [425000, 403750, 339521, 453865, 466535];
  const ratioRows = [
    { label: "Gross Receipts", vals: revByYear, type: "currency" },
    { label: "Net Profit After Depreciation Before Tax", vals: YEARS.map((_, yi) => revByYear[yi] - opexByYear[yi] - 75000 * 12), type: "currency" },
    { label: "Fixed Assets", vals: fixedAssets, type: "currency" },
    { type: "spacer" },
    { label: "Net Profit Ratio", vals: YEARS.map((_, yi) => revByYear[yi] > 0 ? pat[yi] / revByYear[yi] : 0), type: "pct" },
    { label: "Net Sales / Fixed Assets", vals: YEARS.map((_, yi) => fixedAssets[yi] > 0 ? revByYear[yi] / fixedAssets[yi] : 0), type: "multiple" },
    { label: "Optimum Coverage Ratio", vals: YEARS.map(() => 0.75), type: "num" },
  ];
  return (
    <DraggableScroll>
      <SheetHeader title="6. Analysis of Ratios" sub="M/S TREATMENT RANGE HOSPITAL PRIVATE LIMITED" readOnly />
      <table style={{ borderCollapse: "collapse", minWidth: 800 }}>
        <thead>
          <tr>
            <th style={th({ textAlign: "left", minWidth: 280 })}>Particulars</th>
            {YEARS.map(y => <th key={y} style={th({ minWidth: 130 })}>{y}</th>)}
          </tr>
        </thead>
        <tbody>
          {ratioRows.map((row, i) => {
            if (row.type === "spacer") return <tr key={i}><td colSpan={6} style={{ height: 8 }} /></tr>;
            return (
              <tr key={i} style={{ background: i % 2 === 0 ? C.bg1 : C.bg0 }}>
                <td style={{ ...td0(), color: C.text1 }}>{row.label}</td>
                {row.vals.map((v, yi) => (
                  <td key={yi} style={{ ...td0(), color: v < 0 ? C.redL : C.text0, fontFamily: "monospace", textAlign: "right" }}>
                    {row.type === "pct" ? fmtPct(v) : row.type === "multiple" ? `${v.toFixed(2)}x` : row.type === "currency" ? fmtINR(v, true) : v.toFixed(2)}
                  </td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
    </DraggableScroll>
  );
}

// ─── SHEET: DSCR ─────────────────────────────────────────────────────────────
function DSCR({ revP1, opexP1 }) {
  const revByYear = YEARS.map((_, yi) => calcRevYearly(revP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const opexByYear = YEARS.map((_, yi) => calcOpexYearly(opexP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const deprn = YEARS.map(() => 75000 * 12);
  const pat = YEARS.map((_, yi) => { const e = revByYear[yi] - opexByYear[yi] - deprn[yi]; return e - Math.max(0, e * 0.25); });
  const rows = [
    { label: "Profit After Tax", vals: pat },
    { label: "Int on Term Loan", vals: YEARS.map(() => 0) },
    { label: "Depreciation", vals: deprn },
    { type: "spacer" },
    { label: "Total (Numerator)", vals: YEARS.map((_, yi) => pat[yi] + deprn[yi]), type: "total" },
    { type: "spacer" },
    { label: "DEBT", type: "header" },
    { label: "Int on Term Loan", vals: YEARS.map(() => 0) },
    { label: "Term Loan Repayment", vals: YEARS.map(() => 0) },
    { label: "Total (Denominator)", vals: YEARS.map(() => 0), type: "total" },
    { type: "spacer" },
    { label: "D.S.C.R", vals: YEARS.map((_, yi) => { const d = 0; const n = pat[yi] + deprn[yi]; return d > 0 ? n / d : 0; }), type: "ratio" },
  ];
  return (
    <DraggableScroll>
      <SheetHeader title="DSCR — Debt Service Coverage Ratio" sub="M/S TREATMENT RANGE HOSPITAL PRIVATE LIMITED" readOnly />
      <table style={{ borderCollapse: "collapse", minWidth: 800 }}>
        <thead>
          <tr>
            <th style={th({ textAlign: "left", minWidth: 280 })}>Particulars</th>
            {YEARS.map(y => <th key={y} style={th({ minWidth: 130 })}>{y}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => {
            if (row.type === "spacer") return <tr key={i}><td colSpan={6} style={{ height: 8 }} /></tr>;
            if (row.type === "header") return <tr key={i} style={{ background: C.sectionBg }}><td colSpan={6} style={{ padding: "7px 10px", color: C.gold, fontWeight: 700, fontSize: 10, textTransform: "uppercase" }}>{row.label}</td></tr>;
            const isTotal = row.type === "total";
            return (
              <tr key={i} style={{ background: isTotal ? C.totalBg : i % 2 === 0 ? C.bg1 : C.bg0 }}>
                <td style={{ ...td0(isTotal ? { borderTop: `1px solid ${C.borderLight}` } : {}), color: isTotal ? C.text0 : C.text1, fontWeight: isTotal ? 700 : 400 }}>{row.label}</td>
                {row.vals.map((v, yi) => (
                  <td key={yi} style={{ ...td0(isTotal ? { borderTop: `1px solid ${C.borderLight}` } : {}), color: row.type === "ratio" ? (v >= 1.5 ? C.greenL : v > 0 ? C.gold : C.text2) : v < 0 ? C.redL : isTotal ? C.text0 : C.text1, fontFamily: "monospace", textAlign: "right", fontWeight: isTotal ? 700 : 400 }}>
                    {row.type === "ratio" ? (v > 0 ? `${v.toFixed(2)}x` : "—") : fmtINR(v, true)}
                  </td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
    </DraggableScroll>
  );
}

// ─── SHEET: REPAYMENT SCHEDULE ────────────────────────────────────────────────
function RepaymentSchedule({ loan1, loan2, onFocus }) {
  const l1 = calcLoan(loan1.amount, loan1.duration, loan1.rate, loan1.startDate);
  const l2 = calcLoan(loan2.amount, loan2.duration, loan2.rate, loan2.startDate);
  const [active, setActive] = useState(1);
  const loan = active === 1 ? l1 : l2;
  const cfg = active === 1 ? loan1 : loan2;
  return (
    <div>
      <SheetHeader title="Repayment Schedule" sub="Phase-wise loan amortisation table" />
      <div style={{ display: "flex", gap: 8, padding: "8px 12px", borderBottom: `1px solid ${C.border}` }}>
        {[1, 2].map(p => <button key={p} onClick={() => setActive(p)} style={{ padding: "4px 14px", fontSize: 11, background: active === p ? C.navB : "transparent", border: `1px solid ${active === p ? C.teal : C.border}`, borderRadius: 4, color: active === p ? C.teal : C.text2, cursor: "pointer" }}>Phase {p} Loan</button>)}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 1, padding: "10px 12px", background: C.bg2, borderBottom: `1px solid ${C.border}` }}>
        {[["Loan Amount", fmtINR(cfg.amount)], ["Duration", `${cfg.duration} months`], ["Rate", `${cfg.rate}% p.a.`], ["EMI", fmtINR(loan.emi)], ["Total Interest", fmtINR(loan.totalInterest, true)], ["Start Date", cfg.startDate]].map(([k, v]) => (
          <div key={k} style={{ padding: "6px 10px" }}>
            <div style={{ fontSize: 10, color: C.text2, marginBottom: 2 }}>{k}</div>
            <div style={{ fontSize: 13, color: C.gold, fontFamily: "monospace", fontWeight: 700 }}>{v}</div>
          </div>
        ))}
      </div>
      <div style={{ maxHeight: 400, overflowY: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%" }}>
          <thead style={{ position: "sticky", top: 0 }}>
            <tr>
              {["EMI No", "Date", "Opening Bal.", "Principal", "Interest", "EMI", "Closing Bal."].map(h => <th key={h} style={th({ minWidth: h === "Date" ? 100 : 120 })}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {loan.rows.map((row, i) => (
              <tr key={i} style={{ background: i % 2 === 0 ? C.bg1 : C.bg0 }}>
                <td style={{ ...td0(), color: C.text2, textAlign: "center" }}>{row.no}</td>
                <td style={{ ...td0(), color: C.text1, fontSize: 11 }}>{row.date}</td>
                <td style={{ ...td0(), color: C.text0, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(row.opening, true)}</td>
                <td style={{ ...td0(), color: C.blueL, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(row.principal, true)}</td>
                <td style={{ ...td0(), color: C.redL, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(row.interest, true)}</td>
                <td style={{ ...td0(), color: C.gold, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(row.emi, true)}</td>
                <td style={{ ...td0(), color: C.text0, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(row.closing, true)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── SHEET: FA SCHEDULE ───────────────────────────────────────────────────────
function FASchedule({ assets, setAssets, onFocus }) {
  const computed = calcFA(assets);
  const updAsset = (i, k, v) => setAssets(p => p.map((a, ai) => ai !== i ? a : { ...a, [k]: v }));
  return (
    <DraggableScroll>
      <SheetHeader title="FA Schedule — Fixed Assets Schedule" sub="Schedule - 01" />
      <table style={{ borderCollapse: "collapse", minWidth: 900 }}>
        <thead>
          <tr>
            {["Description of Asset", "Rate", "Opening WDV", "Addition >180d", "Addition <180d", "Dep (31.3.26)", "Closing (31.3.26)", "Addition Y2", "Dep Y2", "Closing (31.3.27)"].map((h, i) => (
              <th key={i} style={th({ textAlign: i >= 1 ? "right" : "left", minWidth: i === 0 ? 160 : 90 })}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {computed.map((a, i) => (
            <tr key={i} style={{ background: i % 2 === 0 ? C.bg1 : C.bg0 }}>
              <td style={td0()}><TI v={a.name} onChange={v => updAsset(i, "name", v)} onFocus={() => onFocus && onFocus("FA Schedule", a.name, "Asset Name", a.name)} placeholder="Asset name..." style={{ color: C.text1 }} /></td>
              <td style={td0()}><EC v={a.rate} type="pct" onChange={v => updAsset(i, "rate", v)} onFocus={() => onFocus && onFocus("FA Schedule", a.name, "Rate", a.rate)} /></td>
              <td style={td0()}><EC v={a.opening} type="num" onChange={v => updAsset(i, "opening", v)} onFocus={() => onFocus && onFocus("FA Schedule", a.name, "Opening WDV", a.opening)} /></td>
              <td style={td0()}><EC v={a.addAbove} type="num" onChange={v => updAsset(i, "addAbove", v)} onFocus={() => onFocus && onFocus("FA Schedule", a.name, "Addition >180d", a.addAbove)} /></td>
              <td style={td0()}><EC v={a.addBelow} type="num" onChange={v => updAsset(i, "addBelow", v)} onFocus={() => onFocus && onFocus("FA Schedule", a.name, "Addition <180d", a.addBelow)} /></td>
              <td style={{ ...td0(), color: C.redL, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(a.dep1, true)}</td>
              <td style={{ ...td0(), color: C.tealL, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(a.closing1, true)}</td>
              <td style={{ ...td0(), color: C.inputBlue, fontFamily: "monospace", textAlign: "right" }}>—</td>
              <td style={{ ...td0(), color: C.redL, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(a.dep2, true)}</td>
              <td style={{ ...td0(), color: C.tealL, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(a.closing2, true)}</td>
            </tr>
          ))}
          <tr style={{ background: C.totalBg }}>
            <td colSpan={2} style={{ ...td0({ borderTop: `1px solid ${C.borderLight}` }), color: C.gold, fontWeight: 700 }}>Total</td>
            {[
              computed.reduce((s, a) => s + a.opening, 0), computed.reduce((s, a) => s + a.addAbove, 0), computed.reduce((s, a) => s + a.addBelow, 0),
              computed.reduce((s, a) => s + a.dep1, 0), computed.reduce((s, a) => s + a.closing1, 0),
              0, computed.reduce((s, a) => s + a.dep2, 0), computed.reduce((s, a) => s + a.closing2, 0)
            ].map((v, i) => <td key={i} style={{ ...td0({ borderTop: `1px solid ${C.borderLight}` }), color: C.gold, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(v, true)}</td>)}
          </tr>
        </tbody>
      </table>
    </DraggableScroll>
  );
}

// ─── SHEET: 3B CAPITAL COSTING ────────────────────────────────────────────────
function CapitalCosting({ data, setData, onFocus }) {
  const updItem = (ci, ii, k, v) => setData(p => p.map((cat, ci2) => ci2 !== ci ? cat : { ...cat, items: cat.items.map((it, ii2) => ii2 !== ii ? it : { ...it, [k]: v }) }));
  const yLabels = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"];
  return (
    <DraggableScroll>
      <SheetHeader title="3b. Costing" sub="Direct costs associated with generating revenue (COGS)" />
      <table style={{ borderCollapse: "collapse", minWidth: 900 }}>
        <thead>
          <tr>
            {["S.No", "Nature of Expense", "Total (₹)", ...yLabels].map((h, i) => (
              <th key={i} style={th({ textAlign: i >= 2 ? "right" : "left", minWidth: i === 1 ? 200 : 90 })}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.map((cat, ci) => [
            <tr key={`c${ci}`} style={{ background: C.sectionBg }}>
              <td style={{ ...td0(), color: C.gold, fontWeight: 700, fontFamily: "monospace" }}>{cat.sno}</td>
              <td style={{ ...td0(), color: C.text0, fontWeight: 700 }}>{cat.category}</td>
              <td style={{ ...td0(), color: C.tealL, fontFamily: "monospace", textAlign: "right", fontWeight: 700 }}>{fmtINR(cat.items.reduce((s, it) => s + (it.total || 0), 0), true)}</td>
              {["y1", "y2", "y3", "y4", "y5"].map(k => <td key={k} style={{ ...td0(), color: C.tealL, fontFamily: "monospace", textAlign: "right" }}>{fmtINR(cat.items.reduce((s, it) => s + (it[k] || 0), 0), true)}</td>)}
            </tr>,
            ...cat.items.map((it, ii) => (
              <tr key={`${ci}-${ii}`} style={{ background: ii % 2 === 0 ? C.bg1 : C.bg0 }}>
                <td style={{ ...td0(), color: C.text2, fontSize: 10 }} />
                <td style={{ ...td0(), paddingLeft: 20 }}><TI v={it.name} onChange={v => updItem(ci, ii, "name", v)} onFocus={() => onFocus && onFocus("Capital Costing", it.name, "Asset Name", it.name)} style={{ color: C.text1 }} /></td>
                <td style={td0()}><EC v={it.total} type="num" onChange={v => updItem(ci, ii, "total", v)} onFocus={() => onFocus && onFocus("Capital Costing", it.name, "Total Amount", it.total)} /></td>
                {["y1", "y2", "y3", "y4", "y5"].map(k => <td key={k} style={td0()}><EC v={it[k]} type="num" onChange={v => updItem(ci, ii, k, v)} onFocus={() => onFocus && onFocus("Capital Costing", it.name, `Year ${k.slice(1)}`, it[k])} /></td>)}
              </tr>
            )),
            <tr key={`sp${ci}`}><td colSpan={8} style={{ height: 8 }} /></tr>
          ])}
        </tbody>
      </table>
    </DraggableScroll>
  );
}

// ─── GENERIC EMPTY SHEET ─────────────────────────────────────────────────────
function EmptySheet({ title, sub }) {
  return (
    <div>
      <SheetHeader title={title} sub={sub} />
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: 300, flexDirection: "column", gap: 12 }}>
        <div style={{ fontSize: 36 }}>📋</div>
        <div style={{ color: C.text2, fontSize: 13 }}>This sheet is referenced from other sheets</div>
        <div style={{ color: C.text3, fontSize: 11 }}>{sub}</div>
      </div>
    </div>
  );
}

// ─── SHARED COMPONENTS ────────────────────────────────────────────────────────
function SheetHeader({ title, sub, readOnly }) {
  return (
    <div style={{ padding: "8px 14px 6px", background: C.nav, borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
      <div>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.text0, letterSpacing: "-0.01em" }}>{title}</div>
        {sub && <div style={{ fontSize: 10, color: C.text2, marginTop: 1 }}>{sub}</div>}
      </div>
      {readOnly && <div style={{ fontSize: 10, color: C.teal, padding: "2px 8px", border: `1px solid rgba(42,158,158,0.3)`, borderRadius: 4 }}>Read-only — auto-calculated</div>}
    </div>
  );
}

// ─── AI SYSTEM PROMPT ────────────────────────────────────────────────────────
const AI_SYSTEM = `You are the Financial AI Strategist, a professional AI assistant specialized in building investor-grade financial models.

### CRITICAL: YOU MUST OUTPUT [DATA] TAGS
Every time the user confirms data (revenue, costs, funding), you MUST emit [DATA] tags to update the model. Never just say "I've added it" - you MUST include the data tag.

### REQUIRED DATA TAGS
When user confirms revenue:
[DATA: {"type":"addRevenueStream","streamName":"Subscription","productName":"SaaS Subscription","units":25,"price":30}]

When user confirms costs:
[DATA: {"type":"addOpex","category":"Team Salaries","subCategory":"Salaries","units":4,"cost":4000}]

When user confirms funding:
[DATA: {"type":"setFunding","amount":175000}]

When you need to ask questions:
[SUGGESTIONS: ["option1", "option2", "option3"]]

### RULES
- ALWAYS include [DATA] tags when user confirms any data
- Ask short questions (under 15 words)
- End responses with [SUGGESTIONS] tag

### DISCOVERY FLOW
1. Business type → 2. Revenue model → 3. Price → 4. Volume → 5. Costs → 6. Funding

### STYLE
- Short, crisp responses
- When user says "Yes, add it" or confirms data → MUST emit [DATA] tag`;

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function DoctyModel() {
  const MODEL_DB = {
    edtech: {
      label: "EdTech",
      streams: ["Course Sales", "Subscription Plans", "Corporate Training", "Certifications", "Live Workshops"],
      products: ["Online Courses", "Tutoring Platform", "Learning App", "Test Prep", "Certification Programs"],
      substreams: ["Recorded Courses", "Live Cohorts", "1-on-1 Tutoring", "Mobile App Access", "Test Series"],
      audience: ["K-12 Students", "College Students", "Working Professionals", "Exam Aspirants", "Schools/Institutions"],
      pricing: ["Subscription", "Per Course", "Freemium + Paid", "Cohort-based", "B2B Licensing"],
      opex: ["Instructor Payout", "Platform Hosting", "Marketing", "Support Team", "Content Production"]
    },
    saas: {
      label: "SaaS",
      streams: ["Starter Plan", "Pro Plan", "Enterprise Plan", "API Usage", "Add-ons"],
      products: ["Starter Plan", "Pro Plan", "Enterprise Plan", "API Access", "Add-ons"],
      substreams: ["Monthly Billing", "Annual Billing", "Per-seat Pricing", "Usage-based API", "Implementation Fee"],
      audience: ["SMBs", "Mid-market", "Enterprise", "Developers", "Agencies"],
      pricing: ["Monthly Subscription", "Annual Subscription", "Usage-based", "Tiered Pricing", "Hybrid"],
      opex: ["Engineering Team", "Cloud Infrastructure", "Sales Team", "Customer Success", "Marketing"]
    },
    healthcare: {
      label: "Healthcare",
      streams: ["Consultations", "Diagnostics", "Telemedicine", "Pharmacy", "Packages"],
      products: ["Consultations", "Diagnostics", "Telemedicine", "Pharmacy", "Health Packages"],
      substreams: ["General OPD", "Specialist OPD", "Lab Tests", "Home Collection", "Annual Health Plans"],
      audience: ["Urban Families", "Working Professionals", "Senior Citizens", "Corporate Employees"],
      pricing: ["Per Visit", "Per Test", "Package-based", "Membership", "Insurance-linked"],
      opex: ["Doctor Salaries", "Nursing Staff", "Rent", "Medical Consumables", "Admin"]
    },
    consulting: {
      label: "Consulting",
      streams: ["Retainers", "Project Consulting", "Workshops", "Advisory Calls", "Audits"],
      products: ["Retainers", "Project Consulting", "Workshops", "Advisory Calls", "Audits"],
      substreams: ["Monthly Retainer", "Strategy Projects", "Leadership Workshops", "CXO Advisory", "Process Audits"],
      audience: ["Startups", "SMEs", "Enterprises", "Founders", "CXOs"],
      pricing: ["Retainer", "Per Project", "Hourly", "Outcome-based", "Hybrid"],
      opex: ["Consultant Salaries", "Travel", "Office Rent", "Software Licenses", "Business Development"]
    },
    ecommerce: {
      label: "E-commerce",
      streams: ["Direct Sales", "Marketplace Sales", "Subscriptions", "Bundles", "Wholesale"],
      products: ["Direct Product Sales", "Marketplace Sales", "Subscriptions", "Bundles", "Wholesale"],
      substreams: ["Website Orders", "Marketplace Orders", "Repeat Subscriptions", "Combo Packs", "Bulk Orders"],
      audience: ["Retail Consumers", "D2C Buyers", "Resellers", "B2B Buyers"],
      pricing: ["Per Unit", "Bundle Pricing", "Subscription", "Tiered", "Promo-driven"],
      opex: ["Inventory", "Fulfillment", "Ad Spend", "Warehouse", "Returns"]
    }
  };


  const [d, setD] = useState(INIT);
  const [sheet, setSheet] = useState("1. Basics");
  const CONTEXT_QUESTIONS = [];
  const [contextStep, setContextStep] = useState(0);
  const [businessModel, setBusinessModel] = useState("consulting");
  const [msgs, setMsgs] = useState([{
    role: "assistant",
    text: "Hi! I'm your **Financial Strategist**. Tell me about your business - what do you do and who are your customers?"
  }]);
  const [cellSuggestion, setCellSuggestion] = useState(null);
  const [selectedSuggestion, setSelectedSuggestion] = useState("");
  const [llmSuggestions, setLlmSuggestions] = useState(["Healthcare Clinic", "EdTech Platform", "SaaS Startup", "E-commerce Store", "Consulting Agency", "Pharmacy", "Manufacturing"]);
  const [input, setInput] = useState("");
  const [refreshKey, setRefreshKey] = useState(0);
  const [loading, setLoading] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [chatOpen, setChatOpen] = useState(true);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const endRef = useRef(null);
  // Schema generation Modal State
  const [showImportModal, setShowImportModal] = useState(false);
  const [importText, setImportText] = useState("");
  const [importLoading, setImportLoading] = useState(false);

  const [stage, setStage] = useState("discovery");
  const [completion, setCompletion] = useState(0);

  useEffect(() => { endRef.current?.scrollIntoView({ behavior: "smooth" }); }, [msgs]);


  const handleImportPlan = async () => {
    if (!importText.trim()) return;
    setImportLoading(true);
    try {
      const res = await fetch("/api/generate-from-plan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ strategyPlan: importText })
      });
      if (!res.ok) {
        const errorData = await res.json();
        throw new Error(errorData.error || "Failed to parse plan");
      }
      const data = await res.json();

      // Update state with parsed data
      setD(prev => ({
        ...prev,
        basics: data.basics || prev.basics,
        revP1: data.revP1 || prev.revP1,
        revP2: data.revP2 || prev.revP2,
        opexP1: data.opexP1 || prev.opexP1,
        capex: data.capex || prev.capex,
        totalProjectCost: data.totalProjectCost || prev.totalProjectCost,
        loan1: data.loan1 || prev.loan1,
        loan2: data.loan2 || prev.loan2,
        fixedAssets: data.fixedAssets || prev.fixedAssets,
      }));

      setShowImportModal(false);
      setMsgs(p => [...p, { role: "assistant", text: "I've successfully imported your Strategic Execution Plan and updated all the financial projections across Revenue, OPEX, and CAPEX. Review the sheets and let me know if you want to tweak any numbers!" }]);
      setImportText(""); // Clear the textarea after successful import
    } catch (e) {
      setMsgs(p => [...p, { role: "assistant", text: `Could not import plan: ${(e && e.message) ? e.message : "Unknown error"}` }]);
    } finally {
      setImportLoading(false);
    }
  };

  const handleDownloadExcel = async () => {
    if (downloading) return;
    setDownloading(true);
    try {
      // POST the full model state → server generates Excel from scratch
      // using the same calcRevYearly / calcOpexYearly as the UI
      const res = await fetch("/api/export-model", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(d),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Export failed (${res.status})`);
      }
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const bizName = String(d?.basics?.tradeName || d?.basics?.legalName || "OnEasy_Financial_Agent").replace(/[^a-zA-Z0-9]/g, "_").slice(0, 30);
      a.download = `${bizName}_Financial_Model.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (e) {
      setMsgs((p) => [...p, { role: "assistant", text: `Could not download Excel: ${(e && e.message) ? e.message : "Unknown error"}` }]);
    } finally {
      setDownloading(false);
    }
  };


  const writePatchesToExcel = async (patches) => {
    const safe = Array.isArray(patches) ? patches.filter(p => p && p.sheet && p.cell) : [];
    console.log("[DEBUG] writePatchesToExcel received:", safe);
    if (!safe.length) {
      console.log("[DEBUG] No valid patches to write");
      return;
    }
    try {
      const res = await fetch("/api/excel-fill", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ patches: safe }),
      });
      const result = await res.json();
      console.log("[DEBUG] excel-fill response:", result);
      if (result.success) {
        setRefreshKey(k => k + 1);
      }
    } catch (e) {
      console.log("[DEBUG] excel-fill error:", e);
    }
  };

  const applyRevenueAction = (prev, action) => {
    const streamName = String(action.streamName || "Revenue").trim();
    const subName = String(action.subName || "General").trim();
    const productName = String(action.productName || action.label || "Service").trim();
    const finalSub = productName || subName;
    const qty = Math.max(0, Number(action.units) || 0);
    const price = Math.max(0, Number(action.price ?? action.value) || 0);

    const next = { ...prev, revP1: prev.revP1.map(g => ({ ...g, items: g.items.map(it => ({ ...it })) })) };

    let targetGroup = next.revP1.find(g => String(g.header || "").trim().toLowerCase() === streamName.toLowerCase());
    if (!targetGroup) {
      targetGroup = next.revP1.find(g => !String(g.header || "").trim() || g.header.startsWith("Revenue Stream"));
    }
    if (!targetGroup) {
      targetGroup = { id: String(next.revP1.length + 1), header: streamName, items: Array(5).fill(null).map((_, i) => ({ id: `${next.revP1.length + 1}${String.fromCharCode(97 + i)}`, sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 })) };
      next.revP1.push(targetGroup);
    }

    if (!targetGroup.header || targetGroup.header.startsWith("Revenue Stream")) {
      targetGroup.header = streamName;
    }

    let targetItem = targetGroup.items.find(it => String(it.sub || "").trim().toLowerCase() === finalSub.toLowerCase());
    if (!targetItem) {
      targetItem = targetGroup.items.find(it => !String(it.sub || "").trim() || it.sub.startsWith("Sub service"));
    }
    if (!targetItem) {
      targetItem = { id: `${targetGroup.id}${String.fromCharCode(97 + targetGroup.items.length)}`, sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 };
      targetGroup.items.push(targetItem);
    }

    targetItem.sub = finalSub;
    targetItem.qty = qty;
    targetItem.price = price;
    return next;
  };

  const applyOpexAction = (prev, action) => {
    const cat = String(action.category || "Expense").trim();
    const sub = String(action.subCategory || action.label || "General").trim();
    const units = Math.max(0, Number(action.units) || 1);
    const cost = Math.max(0, Number(action.price ?? action.value) || 0);

    const next = { ...prev, opexP1: prev.opexP1.map(g => ({ ...g, items: g.items.map(it => ({ ...it })) })) };

    let targetGroup = next.opexP1.find(g => String(g.header || "").trim().toLowerCase() === cat.toLowerCase());
    if (!targetGroup) {
      targetGroup = next.opexP1.find(g => !String(g.header || "").trim() || g.header.startsWith("Expense Category") || g.header.startsWith("OPEX Group"));
    }

    if (!targetGroup) {
      targetGroup = { id: String(next.opexP1.length + 1), header: cat, items: Array(5).fill(null).map((_, i) => ({ id: `${next.opexP1.length + 1}${String.fromCharCode(97 + i)}`, sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 })) };
      next.opexP1.push(targetGroup);
    }

    if (!targetGroup.header || targetGroup.header.startsWith("Expense Category") || targetGroup.header.startsWith("OPEX Group")) {
      targetGroup.header = cat;
    }

    let targetItem = targetGroup.items.find(it => String(it.sub || "").trim().toLowerCase() === sub.toLowerCase());
    if (!targetItem) {
      targetItem = targetGroup.items.find(it => !String(it.sub || "").trim() || it.sub.startsWith("Expense item") || it.sub.startsWith("Item..."));
    }
    if (!targetItem) {
      targetItem = { id: `${targetGroup.id}${String.fromCharCode(97 + targetGroup.items.length)}`, sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 };
      targetGroup.items.push(targetItem);
    }

    targetItem.sub = sub;
    targetItem.qty = units;
    targetItem.cost = cost;
    return next;
  };

  const applyDataAction = async (action) => {
    if (!action || typeof action !== "object") return;
    const t = String(action.type || "");

    if (t === "navigateTab" && action.tab) {
      setSheet(String(action.tab));
      return;
    }

    if (t === "setBusinessInfo") {
      setD(prev => ({
        ...prev,
        basics: {
          ...prev.basics,
          legalName: action.legalName ?? prev.basics.legalName,
          tradeName: action.tradeName ?? prev.basics.tradeName,
          address: action.address ?? prev.basics.address,
          email: action.email ?? prev.basics.email,
          contact: action.phone ?? action.contact ?? prev.basics.contact,
          description: action.description ?? prev.basics.description,
          startDateP1: action.startDate ?? prev.basics.startDateP1,
        }
      }));
      await writePatchesToExcel(dataActionToPatches({ type: "setBusinessInfo", ...action }));
      return;
    }

    if (t === "setFunding") {
      await writePatchesToExcel(dataActionToPatches({ type: "setFunding", ...action }));
      return;
    }

    if (t === "setAssumptions" || t === "setAssumption") {
      const mapped = t === "setAssumption"
        ? { [String(action.key || "")]: action.value }
        : action;
      await writePatchesToExcel(dataActionToPatches({ type: "setAssumptions", ...mapped }));
      return;
    }

    if (t === "editCell" && action.sheet && action.cell && action.value !== undefined) {
      try {
        await fetch("/api/edit-cell", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ sheet: action.sheet, cell: action.cell, value: action.value }),
        });
      } catch {
        // non-blocking
      }
      return;
    }

    if (t === "addRevenueStream") {
      console.log("[DEBUG] addRevenueStream action:", action);
      const patches = dataActionToPatches({ type: "addRevenueStream", ...action });
      console.log("[DEBUG] Revenue patches:", patches);
      setD(prev => applyRevenueAction(prev, action));
      await writePatchesToExcel(patches);
      setSheet("A.I Revenue Streams - P1");
      return;
    }

    if (t === "addOpex") {
      console.log("[DEBUG] addOpex action:", action);
      const patches = dataActionToPatches({ type: "addOpex", ...action });
      console.log("[DEBUG] OPEX patches:", patches);
      setD(prev => applyOpexAction(prev, action));
      await writePatchesToExcel(patches);
      setSheet("A.IIOPEX");
    }
    if (t === "setFunding") {
      console.log("[DEBUG] setFunding action:", action);
      const patches = dataActionToPatches({ type: "setFunding", ...action });
      console.log("[DEBUG] Funding patches:", patches);
      await writePatchesToExcel(patches);
      setSheet("1. Basics");
    }
  };

  const parseList = (text) => String(text || "")
    .split(/,|\n|;/)
    .map(s => s.trim())
    .filter(Boolean)
    .slice(0, 8);

  const parseNumbers = (text) => {
    const m = String(text || "").match(/\d[\d,]*(?:\.\d+)?/g) || [];
    return m.map(v => Number(String(v).replace(/,/g, ""))).filter(n => Number.isFinite(n) && n > 0);
  };

  const parsePercentages = (text) => {
    const matches = String(text || "").match(/(\d+(?:\.\d+)?)\s*%/g) || [];
    return matches
      .map(m => Number(String(m).replace("%", "").trim()))
      .filter(n => Number.isFinite(n) && n >= 0);
  };

  const parseSubstreamMap = (text, streams) => {
    const map = {};
    const normalizedStreams = (streams || []).map(s => String(s || "").trim()).filter(Boolean);
    const chunks = String(text || "").split(/\||\n/).map(s => s.trim()).filter(Boolean);
    for (const chunk of chunks) {
      const idx = chunk.indexOf(":");
      if (idx > 0) {
        const key = chunk.slice(0, idx).trim();
        const values = chunk
          .slice(idx + 1)
          .split(/,|;/)
          .map(s => s.trim())
          .filter(Boolean)
          .slice(0, 5);
        if (key && values.length) map[key.toLowerCase()] = values;
      }
    }
    if (Object.keys(map).length) return map;
    const fallback = parseList(text).slice(0, 5);
    if (fallback.length && normalizedStreams[0]) {
      map[normalizedStreams[0].toLowerCase()] = fallback;
    }
    return map;
  };

  const buildCleanRevenueState = (modelKey, products = []) => {
    const template = getIndustryTemplate(modelKey);
    const growth = getGrowthRates(template?.growthProfile);
    const seeded = Array.isArray(products) && products.length
      ? products.map(name => ({ stream: "Primary Revenue", name: String(name).trim(), price: 0, quantity: 0 }))
      : (template?.revenueStreams || []).map(item => ({ ...item }));

    const next = Array.from({ length: Math.max(1, Math.ceil(seeded.length / 5), INIT.revP1.length) }, (_, gi) => ({
      id: String(gi + 1),
      header: "",
      items: Array.from({ length: 5 }, (_, ii) => ({ id: `${gi + 1}${String.fromCharCode(97 + ii)}`, sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 }))
    }));

    seeded.slice(0, next.length * 5).forEach((item, idx) => {
      const gi = Math.floor(idx / 5);
      const ii = idx % 5;
      if (!next[gi] || !next[gi].items[ii]) return;
      next[gi].header = String(item.stream || next[gi].header || "Primary Revenue").trim();
      next[gi].items[ii] = {
        ...next[gi].items[ii],
        sub: String(item.name || "").trim(),
        qty: Number(item.quantity) || 0,
        price: Number(item.price) || 0,
        gY1: growth.y1 || 0,
        gY2: growth.y2 || 0,
        gY3: growth.y3 || 0,
        gY4: growth.y4 || 0,
        gY5: growth.y5 || 0,
      };
    });
    return next;
  };

  const buildRevenueFromStreams = (streams = []) => {
    const list = (Array.isArray(streams) ? streams : []).map(s => String(s || "").trim()).filter(Boolean).slice(0, 5);
    const next = INIT.revP1.map(g => ({
      ...g,
      header: "",
      items: g.items.map(it => ({ ...it, sub: "", qty: 0, price: 0 })),
    }));
    list.forEach((name, i) => {
      if (!next[i]) return;
      next[i].header = name;
    });
    return next;
  };

  const buildRevenueWithSubstreams = (streams = [], subMap = {}) => {
    const template = getIndustryTemplate(businessModel);
    const growth = getGrowthRates(template?.growthProfile);
    const priceLookup = new Map((template?.revenueStreams || []).map(item => [String(item.name || "").toLowerCase(), Number(item.price) || 0]));
    const next = buildRevenueFromStreams(streams);
    streams.forEach((streamName, si) => {
      if (!next[si]) return;
      const subs = subMap[String(streamName || "").toLowerCase()] || [];
      subs.slice(0, 5).forEach((sub, ii) => {
        if (!next[si].items[ii]) return;
        const key = String(sub || "").toLowerCase();
        next[si].items[ii] = {
          ...next[si].items[ii],
          sub: String(sub).trim(),
          qty: 0,
          price: priceLookup.get(key) || 0,
          gY1: growth.y1 || 0,
          gY2: growth.y2 || 0,
          gY3: growth.y3 || 0,
          gY4: growth.y4 || 0,
          gY5: growth.y5 || 0,
        };
      });
    });
    return next;
  };

  const buildCleanOpexState = (modelKey, costs = []) => {
    const template = getIndustryTemplate(modelKey);
    const growth = getGrowthRates(template?.growthProfile);
    const seeded = Array.isArray(costs) && costs.length
      ? costs.map(name => ({ category: "Operating Expense", name: String(name).trim(), monthlyCost: 0 }))
      : (template?.opex || []).map(item => ({ ...item }));

    const next = Array.from({ length: Math.max(1, Math.ceil(seeded.length / 7), INIT.opexP1.length) }, (_, gi) => ({
      id: String(gi + 1),
      header: "",
      items: Array.from({ length: 7 }, (_, ii) => ({ id: `${gi + 1}${String.fromCharCode(97 + ii)}`, sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 }))
    }));

    seeded.slice(0, next.length * 7).forEach((item, idx) => {
      const gi = Math.floor(idx / 7);
      const ii = idx % 7;
      if (!next[gi] || !next[gi].items[ii]) return;
      next[gi].header = String(item.category || next[gi].header || "Operating Expense").trim();
      next[gi].items[ii] = {
        ...next[gi].items[ii],
        sub: String(item.name || "").trim(),
        qty: 1,
        cost: Number(item.monthlyCost) || 0,
        gY1: growth.y1 || 0,
        gY2: growth.y2 || 0,
        gY3: growth.y3 || 0,
        gY4: growth.y4 || 0,
        gY5: growth.y5 || 0,
      };
    });
    return next;
  };

  const parseLaunchDate = (text) => {
    const raw = String(text || "").trim();
    if (!raw) return "";
    const low = raw.toLowerCase();
    const today = new Date();
    if (low.includes("next month") || low.includes("next moth")) {
      return new Date(today.getFullYear(), today.getMonth() + 1, 1).toISOString().slice(0, 10);
    }
    const d = new Date(raw);
    if (!Number.isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    return raw;
  };




  const handleCellFocus = async (sheetName, rowLabel, fieldName, currentValue) => {
    if (loading) return;
    setCellSuggestion({ sheet: sheetName, row: rowLabel, field: fieldName, loading: true });
    try {
      console.log("[AGENT] Action: Cell input suggestion (LLM)");
      const res = await fetch("/api/chat-simple", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          system: AI_SYSTEM,
          messages: [{ role: "user", text: `Suggest a realistic, industry-standard value for the following cell in a ${businessModel} financial model:\nSheet: ${sheetName}\nRow/Item: ${rowLabel}\nField: ${fieldName}\nCurrent Value: ${currentValue}\n\nRespond ONLY with a JSON object: {"suggestion": "value", "reason": "1-sentence explanation"}` }],
        })
      });
      const data = await res.json();
      if (!res.ok) throw new Error("Suggestion failed");
      let parsed;
      try {
        parsed = JSON.parse(data.text.replace(/```json|```/g, "").trim());
      } catch {
        parsed = { suggestion: data.text, reason: "" };
      }
      setCellSuggestion({ sheet: sheetName, row: rowLabel, field: fieldName, suggestion: parsed.suggestion, reason: parsed.reason, loading: false });
    } catch {
      setCellSuggestion(null);
    }
  };

  const SHEET_GROUPS = [
    {
      type: 'single',
      id: "1. Basics",
      label: "Model Overview (Cover)",
      icon: Layout
    },
    {
      type: 'group',
      label: "Your inputs",
      items: [
        { id: "A. Data Needed", label: "Timeline", icon: Calendar },
        { id: "A.I Revenue Streams - P1", label: "Sales (Phase 1)", icon: TrendingUp },
        { id: "A.I Revenue Streams - P2", label: "Sales (Phase 2)", icon: TrendingUp },
        { id: "A.IIOPEX", label: "SG&A / Costs", icon: DollarSign },
        { id: "A.III CAPEX", label: "Capital Expenditures", icon: Building },
        { id: "Repayment schedule", label: "Financing", icon: CreditCard },
      ]
    },
    {
      type: 'divider',
      label: "calculations and outputs"
    },
    {
      type: 'list',
      items: [
        { id: "B.I Sales - P1", label: "Sales Analysis (P1)", icon: TrendingUp },
        { id: "B.I Sales - P2", label: "Sales Analysis (P2)", icon: TrendingUp },
        { id: "B.II - OPEX", label: "SG&A Analysis", icon: DollarSign },
        { id: "2.Total Project Cost", label: "Project Cost", icon: Building },
        { id: "4. P&L", label: "P&L Statement", icon: FileText },
        { id: "5. Balance sheet", label: "Balance Sheet", icon: Scale },
        { id: "6. Ratios", label: "Financial Ratios", icon: PieChart },
        { id: "FA Schedule", label: "Fixed Assets", icon: Layout },
        { id: "Scenarios", label: "Model Dashboard & Scenarios", icon: BarChart3 },
      ]
    }
  ];

  // Flatten for internal logic compatibility
  const SHEETS = SHEET_GROUPS.flatMap(g => g.items || (g.id ? [g] : []));

  const Sidebar = () => {
    if (!sidebarOpen) {
      return (
        <div style={{ width: 32, height: "100%", background: "#FFF", borderLeft: `1px solid ${C.border}`, display: "flex", alignItems: "flex-start", justifyContent: "center", paddingTop: 14 }}>
          <button
            onClick={() => setSidebarOpen(true)}
            title="Open file sheets"
            style={{
              width: 22,
              height: 40,
              borderRadius: 999,
              border: `1px solid ${C.border}`,
              background: C.bg1,
              color: C.text2,
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              boxShadow: "0 1px 3px rgba(0,0,0,0.06)"
            }}
          >
            <ChevronRight size={14} strokeWidth={2.5} />
          </button>
        </div>
      );
    }

    return (
      <div style={{ width: 280, height: "100%", background: "#FFF", borderLeft: `1px solid ${C.border}`, display: "flex", flexDirection: "column", overflowY: "auto", padding: "16px 12px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12, paddingLeft: 8, gap: 8 }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: C.text3, textTransform: "uppercase", letterSpacing: "0.1em" }}>File Sheets</div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <button
              onClick={handleDownloadExcel}
              disabled={downloading}
              title="Export model"
              style={{
                display: "flex",
                alignItems: "center",
                gap: 4,
                padding: "6px 8px",
                borderRadius: 8,
                border: `1px solid ${C.border}`,
                background: C.bg1,
                color: C.text1,
                cursor: downloading ? "not-allowed" : "pointer",
                fontSize: 11,
                fontWeight: 600
              }}
            >
              <Download size={13} strokeWidth={2} />
              Export
            </button>
            <button
              onClick={() => setShowImportModal(true)}
              title="Import strategy plan"
              style={{
                display: "flex",
                alignItems: "center",
                gap: 4,
                padding: "6px 8px",
                borderRadius: 8,
                border: "none",
                background: `linear-gradient(135deg, ${C.goldL}, ${C.gold})`,
                color: "#fff",
                cursor: "pointer",
                fontSize: 11,
                fontWeight: 600
              }}
            >
              <Sparkles size={13} strokeWidth={2} />
              Import
            </button>
            <button
              onClick={() => setSidebarOpen(false)}
              title="Close file sheets"
              style={{
                width: 24,
                height: 36,
                borderRadius: 999,
                border: `1px solid ${C.border}`,
                background: C.bg1,
                color: C.text2,
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                justifyContent: "center"
              }}
            >
              <span style={{ fontSize: 14, fontWeight: 700 }}>&gt;</span>
            </button>
          </div>
        </div>
        
        {SHEET_GROUPS.map((group, gIdx) => {
          if (group.type === 'single') {
            const Icon = group.icon;
            const active = sheet === group.id;
            return (
              <div 
                key={group.id} 
                onClick={() => setSheet(group.id)}
                style={{ 
                  display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", borderRadius: 8, cursor: "pointer",
                  background: active ? "#F3F4F6" : "transparent", color: active ? C.text0 : C.text2,
                  marginBottom: 4, transition: "all 0.2s"
                }}
              >
                <Icon size={16} strokeWidth={active ? 2.5 : 2} />
                <span style={{ fontSize: 13, fontWeight: active ? 600 : 500 }}>{group.label}</span>
              </div>
            );
          }
          if (group.type === 'group') {
            return (
              <div key={gIdx} style={{ marginTop: 12 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 12px", borderRadius: 8, background: C.bg0, marginBottom: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                     <Layout size={14} color={C.text3} />
                     <span style={{ fontSize: 12, fontWeight: 700, color: C.text1, textTransform: "uppercase", letterSpacing: "0.02em" }}>{group.label}</span>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 4, opacity: 0.6 }}>
                    <div style={{ width: 6, height: 6, borderRadius: "50%", background: C.greenL }} />
                    <span style={{ fontSize: 9, fontWeight: 600, color: C.text3, textTransform: "uppercase" }}>Active</span>
                  </div>
                </div>
                <div style={{ paddingLeft: 4 }}>
                  {group.items.map(item => {
                    const Icon = item.icon;
                    const active = sheet === item.id;
                    return (
                      <div 
                        key={item.id} 
                        onClick={() => setSheet(item.id)}
                        style={{ 
                          display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 12px", borderRadius: 8, cursor: "pointer",
                          color: active ? C.blue : C.text2,
                          background: active ? `${C.blue}10` : "transparent",
                          marginBottom: 2
                        }}
                      >
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <span style={{ fontSize: 13, fontWeight: active ? 600 : 500 }}>{item.label}</span>
                        </div>
                        <div style={{ width: 32, height: 18, background: active ? C.blue : C.bg3, borderRadius: 10, position: "relative", cursor: "pointer" }}>
                          <div style={{ width: 14, height: 14, background: "#FFF", borderRadius: "50%", position: "absolute", top: 2, right: active ? 2 : 16, transition: "all 0.2s" }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          }
          if (group.type === 'divider') {
            return (
              <div key={gIdx} style={{ margin: "24px 0 16px", display: "flex", alignItems: "center", gap: 12 }}>
                <div style={{ flex: 1, height: 1, background: C.border }} />
                <span style={{ fontSize: 10, fontWeight: 700, color: C.text3, textTransform: "uppercase", letterSpacing: "0.05em" }}>{group.label}</span>
                <div style={{ flex: 1, height: 1, background: C.border }} />
              </div>
            );
          }
          if (group.type === 'list') {
            return (
              <div key={gIdx}>
                {group.items.map(item => {
                  const Icon = item.icon;
                  const active = sheet === item.id;
                  return (
                    <div 
                      key={item.id} 
                      onClick={() => setSheet(item.id)}
                      style={{ 
                        display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", borderRadius: 8, cursor: "pointer",
                        background: active ? "#F3F4F6" : "transparent", color: active ? C.text0 : C.text2,
                        marginBottom: 4
                      }}
                    >
                      <div style={{ width: 32, height: 32, borderRadius: 6, background: "#F3F4F6", display: "flex", alignItems: "center", justifyContent: "center", color: C.text3 }}>
                        <Icon size={16} />
                      </div>
                      <span style={{ fontSize: 13, fontWeight: active ? 600 : 500 }}>{item.label}</span>
                    </div>
                  );
                })}
              </div>
            );
          }
          return null;
        })}
      </div>
    );
  };

  const applyFinancialExtractionToState = async (extracted) => {
    if (!extracted) return;
    console.log("[AGENT] Applying structured extraction:", extracted);

    let nextSnapshot = null;
    setD(prev => {
      let next = { ...prev };
      if (extracted.basics) {
        next.basics = { ...next.basics, ...extracted.basics };
      }
      if (extracted.revenue_streams) {
        extracted.revenue_streams.forEach(stream => {
          (stream.items || []).forEach(item => {
            next = applyRevenueAction(next, {
              streamName: stream.header,
              productName: item.sub,
              units: item.qty,
              price: item.price
            });
          });
        });
      }
      if (extracted.opex_streams) {
        extracted.opex_streams.forEach(stream => {
          (stream.items || []).forEach(item => {
            next = applyOpexAction(next, {
              category: stream.header,
              subCategory: item.sub,
              units: item.qty,
              price: item.cost
            });
          });
        });
      }
      if (extracted.funding) {
        next.totalProjectCost = {
          ...next.totalProjectCost,
          promoterContrib: extracted.funding.promoterContrib ?? next.totalProjectCost.promoterContrib,
          termLoan: extracted.funding.termLoan ?? next.totalProjectCost.termLoan,
          wcLoan: extracted.funding.wcLoan ?? next.totalProjectCost.wcLoan,
        };
        if (extracted.funding.loan1) {
          next.loan1 = { ...next.loan1, ...extracted.funding.loan1 };
        }
      }
      nextSnapshot = next;
      return next;
    });

    if (nextSnapshot) {
      const patches = modelStateToPatches(nextSnapshot);
      await writePatchesToExcel(patches);
    }
  };

  const send = async () => {
    const text = input.trim();
    if (!text || loading) return;
    setInput(""); setLoading(true);
    
    // Add user message immediately
    const nextMsgs = [...msgs, { role: "user", text }];
    setMsgs(nextMsgs);

    try {
      console.log("[AGENT] Action: Structured Financial Agent Chat");
      const res = await fetch("/api/financial-chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          messages: nextMsgs.map(m => ({ role: m.role, content: m.text })),
          knowledgeGraph: d,
          stage: stage
        })
      });

      if (!res.ok) throw new Error("API request failed");

      // Handle Headers (Stage/Completion)
      const newStage = res.headers.get("X-Stage");
      const newCompletion = res.headers.get("X-Completion");
      if (newStage) setStage(newStage);
      if (newCompletion) setCompletion(Number(newCompletion));

      // Read the data stream (Protocol: 0:text, 2:extraction)
      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let assistantText = "";
      let buffer = "";
      
      setMsgs(prev => [...prev, { role: "assistant", text: "" }]);

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop(); 
        
        for (const line of lines) {
          if (line.startsWith("0:")) {
            try {
              const content = JSON.parse(line.substring(2));
              assistantText += content;
              setMsgs(prev => {
                const updated = [...prev];
                updated[updated.length - 1].text = assistantText;
                return updated;
              });
            } catch {}
          } else if (line.startsWith("2:")) {
            try {
              const extracted = JSON.parse(line.substring(2));
              await applyFinancialExtractionToState(extracted);
            } catch (e) {
              console.error("Failed to parse stream extraction:", e);
            }
          }
        }
      }
    } catch (err) {
      console.error("Chat error:", err);
      setMsgs(prev => [...prev, { role: "assistant", text: "I'm sorry, I encountered an error. Please try again." }]);
    } finally {
      setLoading(false);
    }
  };


  const renderSheet = () => {
    switch (sheet) {
      case "1. Basics": return <Basics d={d.basics} setD={v => setD(p => ({ ...p, basics: v(p.basics) }))} onFocus={handleCellFocus} />;
      case "A. Data Needed": return <DataNeeded setSheet={setSheet} />;
      case "A.I Revenue Streams - P1": return <RevStreams groups={d.revP1} setGroups={v => setD(p => ({ ...p, revP1: typeof v === "function" ? v(p.revP1) : v }))} phase="Phase 1" onFocus={handleCellFocus} />;
      case "A.I Revenue Streams - P2": return <RevStreams groups={d.revP2} setGroups={v => setD(p => ({ ...p, revP2: typeof v === "function" ? v(p.revP2) : v }))} phase="Phase 2" perDay onFocus={handleCellFocus} />;
      case "A.IIOPEX": return <OpexSheet groups={d.opexP1} setGroups={v => setD(p => ({ ...p, opexP1: typeof v === "function" ? v(p.opexP1) : v }))} onFocus={handleCellFocus} />;
      case "A.III CAPEX": return <Capex data={d.capex} setData={v => setD(p => ({ ...p, capex: typeof v === "function" ? v(p.capex) : v }))} onFocus={handleCellFocus} />;
      case "B.I Sales - P1": return <SalesSheet groups={d.revP1} phase="Phase 1" />;
      case "B.I Sales - P2": return <SalesSheet groups={d.revP2} phase="Phase 2" />;
      case "B.II - OPEX": return <OpexSheet groups={d.opexP1} setGroups={v => setD(p => ({ ...p, opexP1: typeof v === "function" ? v(p.opexP1) : v }))} onFocus={handleCellFocus} />;
      case "2.Total Project Cost": return <TotalProjectCost data={d.totalProjectCost} setData={v => setD(p => ({ ...p, totalProjectCost: typeof v === "function" ? v(p.totalProjectCost) : v }))} onFocus={handleCellFocus} />;
      case "B.II OPEX - P1": return <OpexSheet groups={d.opexP1} setGroups={v => setD(p => ({ ...p, opexP1: typeof v === "function" ? v(p.opexP1) : v }))} onFocus={handleCellFocus} />;
      case "4. P&L": return <PL revP1={d.revP1} opexP1={d.opexP1} loan1={d.loan1} loan2={d.loan2} />;
      case "3b. Costing - (Capital Exp)": return <CapitalCosting data={d.capex} setData={v => setD(p => ({ ...p, capex: typeof v === "function" ? v(p.capex) : v }))} onFocus={handleCellFocus} />;
      case "5. Balance sheet": return <BalanceSheet revP1={d.revP1} opexP1={d.opexP1} loan1={d.loan1} loan2={d.loan2} />;
      case "6. Ratios": return <Ratios revP1={d.revP1} opexP1={d.opexP1} loan1={d.loan1} loan2={d.loan2} />;
      case "DSCR": return <DSCR revP1={d.revP1} opexP1={d.opexP1} loan1={d.loan1} loan2={d.loan2} />;
      case "Repayment schedule": return <RepaymentSchedule loan1={d.loan1} loan2={d.loan2} onFocus={handleCellFocus} />;
      case "FA Schedule": return <FASchedule assets={d.fixedAssets} setAssets={v => setD(p => ({ ...p, fixedAssets: typeof v === "function" ? v(p.fixedAssets) : v }))} onFocus={handleCellFocus} />;
      case "Phase 1": return <RepaymentSchedule loan1={d.loan1} loan2={d.loan2} onFocus={handleCellFocus} />;
      case "Phase 2": return <RepaymentSchedule loan1={d.loan1} loan2={d.loan2} onFocus={handleCellFocus} />;
      case "Costing": return <EmptySheet title="Costing" sub="Referenced from other sheets — no direct data entry" />;
      case "Scenarios": return <ScenarioDashboard revP1={d.revP1} opexP1={d.opexP1} />;
      default: return <EmptySheet title={sheet} sub="Sheet content" />;
    }
  };

  const revY1 = calcRevYearly(d.revP1).reduce((s, g) => s + g.yearlyTotals[0], 0);
  const opexY1 = calcOpexYearly(d.opexP1).reduce((s, g) => s + g.yearlyTotals[0], 0);

  return (
    <div style={{ display: "flex", height: "100vh", background: C.bg0, fontFamily: "'Poppins', sans-serif", overflow: "hidden" }}>
      {chatOpen && (
        <div style={{ width: 450, flexShrink: 0, display: "flex", flexDirection: "column", background: C.bg1, borderRight: `1px solid ${C.border}` }}>
          <div style={{ padding: "12px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 30, height: 30, borderRadius: 8, background: "linear-gradient(135deg,#2B6CB0,#4299E1)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", boxShadow: "0 2px 4px rgba(43,108,176,0.2)" }}>
                  <Bot size={17} strokeWidth={2.5} />
                </div>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: C.text0 }}>Fina AI Strategist</div>
                  <div style={{ fontSize: 9, color: C.teal, fontWeight: 600, textTransform: "uppercase" }}>
                    {stage.replace("_", " ")}
                  </div>
                </div>
              </div>
              <div style={{ fontSize: 13, fontWeight: 700, color: C.teal }}>{completion}%</div>
            </div>
            
            {/* Progress Bar (Polished) */}
            <div style={{ 
              width: "100%", height: 10, background: C.bg0, borderRadius: 10, 
              overflow: "hidden", border: `1px solid ${C.border}`, padding: 2,
              marginTop: 4
            }}>
              <div 
                style={{ 
                  width: `${completion}%`, 
                  height: "100%", 
                  background: `linear-gradient(90deg, ${C.blue}, ${C.blueL})`, 
                  borderRadius: 10, 
                  transition: "width 1.2s cubic-bezier(0.34, 1.56, 0.64, 1)",
                  boxShadow: `0 0 12px ${C.blue}44`
                }} 
              />
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 2px" }}>
                <span style={{ fontSize: 9, color: completion >= 33 ? C.blue : C.text3, fontWeight: 700, textTransform: "uppercase" }}>Discovery</span>
                <span style={{ fontSize: 9, color: completion >= 66 ? C.blue : C.text3, fontWeight: 700, textTransform: "uppercase" }}>Drafting</span>
                <span style={{ fontSize: 9, color: completion >= 100 ? C.blue : C.text3, fontWeight: 700, textTransform: "uppercase" }}>Review</span>
            </div>

          </div>
          <div style={{ flex: 1, overflowY: "auto", padding: 10 }}>
            {msgs.map((m, i) => (
              <div key={i} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start", marginBottom: 16 }}>
                <div style={{
                  maxWidth: "92%",
                  padding: "12px 16px",
                  borderRadius: m.role === "user" ? "16px 16px 4px 16px" : "4px 16px 16px 16px",
                  background: m.role === "user" ? "#042B3D" : C.bg2,
                  border: `1px solid ${m.role === "user" ? "transparent" : C.border}`,
                  fontSize: 13,
                  lineHeight: 1.6,
                  color: m.role === "user" ? "#FFFFFF" : C.text0,
                  whiteSpace: "pre-wrap",
                  boxShadow: m.role === "user" ? "0 2px 4px rgba(4,43,61,0.15)" : "none"
                }}>
                  {m.text.split(/(\*\*.*?\*\*)/).map((p, j) => p.startsWith("**") && p.endsWith("**") ? <strong key={j} style={{ color: m.role === "user" ? "#FFFFFF" : C.teal }}>{p.slice(2, -2)}</strong> : p)}
                </div>
              </div>
            ))}
            {loading && (
              <div style={{ display: "flex", justifyContent: "flex-start", marginBottom: 8 }}>
                <div style={{ padding: "8px 13px", borderRadius: "2px 10px 10px 10px", background: C.bg2, border: `1px solid ${C.border}`, display: "flex", gap: 4 }}>
                  {[0, 1, 2].map(i => <div key={i} style={{ width: 5, height: 5, borderRadius: "50%", background: C.teal, animation: `p 1.2s ${i * 0.2}s infinite` }} />)}
                  <style>{`@keyframes p{0%,80%,100%{opacity:.3;transform:scale(.8)}40%{opacity:1;transform:scale(1)}}`}</style>
                </div>
              </div>
            )}
            <div ref={endRef} />
          </div>
          <div style={{ padding: "8px 10px", borderTop: `1px solid ${C.border}` }}>
            {cellSuggestion && (
              <div style={{ marginBottom: 10, padding: 10, background: C.bg2, border: `1px solid ${C.teal}`, borderRadius: 8 }}>
                <div style={{ fontSize: 10, color: C.teal, fontWeight: 700, textTransform: "uppercase", marginBottom: 4 }}>AI Suggestion for: {cellSuggestion.field}</div>
                {cellSuggestion.loading ? (
                  <div style={{ fontSize: 11, color: C.text2 }}>Analyzing industry benchmarks...</div>
                ) : (
                  <>
                    <div style={{ fontSize: 13, color: C.gold, fontWeight: 700, marginBottom: 2 }}>{cellSuggestion.suggestion}</div>
                    <div style={{ fontSize: 10, color: C.text1, lineHeight: 1.4, marginBottom: 8 }}>{cellSuggestion.reason}</div>
                    <button
                      onClick={() => setInput(String(cellSuggestion.suggestion))}
                      style={{ padding: "4px 10px", fontSize: 10, background: C.teal, border: "none", borderRadius: 4, color: "#fff", cursor: "pointer", fontWeight: 700 }}
                    >
                      Use Suggestion
                    </button>
                  </>
                )}
              </div>
            )}
            <div style={{ display: "flex", gap: 10, background: C.bg1, border: `1px solid ${C.border}`, borderRadius: 24, padding: "8px 12px", alignItems: "flex-end", boxShadow: "0 2px 8px rgba(0,0,0,0.02)" }}>
              <textarea value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }} placeholder="Describe your business idea..." rows={1} style={{ flex: 1, background: "transparent", border: "none", color: C.text0, fontSize: 13, lineHeight: 1.5, fontFamily: "Poppins, sans-serif", maxHeight: 90, overflowY: "auto", padding: "4px 0" }} />
              <button onClick={send} disabled={loading || !input.trim()} style={{ width: 34, height: 34, borderRadius: 17, background: loading || !input.trim() ? C.border : "#042B3D", border: "none", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, transition: "background 0.2s" }}>
                <Send size={15} strokeWidth={2} color="#fff" style={{ marginLeft: -2 }} />
              </button>
            </div>
            <div style={{ textAlign: "center", marginTop: 8, fontSize: 10, color: C.text3 }}>Press Enter to send · Shift+Enter for new line</div>
          </div>
        </div>
      )}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
        *{box-sizing:border-box;} ::-webkit-scrollbar{width:4px;height:4px;} ::-webkit-scrollbar-track{background:${C.bg0};} ::-webkit-scrollbar-thumb{background:${C.navB};border-radius:2px;}
        input::placeholder,textarea::placeholder{color:${C.text3};} textarea{resize:none;} input:focus,textarea:focus{outline:none;}
        tr:hover td{background:rgba(255,255,255,0.015)!important;}
      `}</style>

      {/* MAIN AREA STAYS LEFT */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden", minWidth: 0 }}>
        {/* Top bar (Consolidated) */}
        <div style={{ 
          display: "grid", 
          gridTemplateColumns: "250px 1fr 320px", 
          alignItems: "center", 
          padding: "0 16px", 
          height: 54, 
          background: "#FFF", 
          borderBottom: `1px solid ${C.border}`, 
          flexShrink: 0,
          boxShadow: "0 1px 3px rgba(0,0,0,0.02)"
        }}>
          {/* Left: Branding */}
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ 
              width: 32, height: 32, borderRadius: 8, 
              background: `linear-gradient(135deg, ${C.teal}, ${C.tealL})`, 
              display: "flex", alignItems: "center", justifyContent: "center", color: "#fff",
              boxShadow: `0 2px 4px ${C.teal}33`
            }}>
              <Activity size={18} strokeWidth={2.5} />
            </div>
            <div>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.text0, lineHeight: 1 }}>Fina AI</div>
              <div style={{ fontSize: 9, color: C.text3, textTransform: "uppercase", letterSpacing: "0.05em", marginTop: 2 }}>Strategist · {SHEETS.length} Sheets</div>
            </div>
          </div>

          {/* Center: Live Stats */}
          <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 24 }}>
            {[
              ["Revenue", revY1, C.teal],
              ["OPEX", opexY1, C.redL],
              ["EBITDA", revY1 - opexY1, revY1 > opexY1 ? C.greenL : C.redL]
            ].map(([label, val, color]) => (
              <div key={label} style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
                <span style={{ fontSize: 9, fontWeight: 700, color: C.text3, textTransform: "uppercase", letterSpacing: "0.05em" }}>{label}</span>
                <span style={{ fontSize: 13, fontWeight: 700, color: color, fontFamily: "monospace" }}>{fmtINR(val, true)}</span>
              </div>
            ))}
            <div style={{ width: 1, height: 20, background: C.border, margin: "0 8px" }} />
            <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 10, color: C.greenL, fontWeight: 600 }}>
              <div style={{ width: 6, height: 6, background: C.greenL, borderRadius: "50%", boxShadow: `0 0 4px ${C.greenL}` }} />
              Live
            </div>
          </div>

          {/* Right: Actions */}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
            {sidebarOpen && (
              <>
                <button
                  onClick={handleDownloadExcel}
                  disabled={downloading}
                  title="Download Excel Model"
                  style={{
                    display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", fontSize: 11,
                    background: C.bg1, border: `1px solid ${C.border}`, borderRadius: 8, color: C.text1,
                    cursor: downloading ? "not-allowed" : "pointer", fontWeight: 600, transition: "all 0.2s"
                  }}
                >
                  <Download size={14} strokeWidth={2} />
                  {downloading ? "..." : "Export"}
                </button>
                
                <button 
                  onClick={() => setShowImportModal(true)} 
                  title="Import Strategy Plan"
                  style={{ 
                    display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", fontSize: 11, 
                    background: `linear-gradient(135deg, ${C.goldL}, ${C.gold})`, border: "none", borderRadius: 8, 
                    color: "#fff", cursor: "pointer", fontWeight: 600, boxShadow: "0 2px 4px rgba(0,0,0,0.1)"
                  }}
                >
                  <Sparkles size={14} strokeWidth={2} />
                  Import
                </button>
              </>
            )}

            <button 
              onClick={() => setChatOpen(v => !v)} 
              style={{ 
                width: 38, height: 38, display: "flex", alignItems: "center", justifyContent: "center",
                background: chatOpen ? `${C.blue}10` : C.bg1, 
                border: `1px solid ${chatOpen ? C.blue : C.border}`, 
                borderRadius: 8, color: chatOpen ? C.blue : C.text2, cursor: "pointer", transition: "all 0.2s"
              }}
              title={chatOpen ? "Hide Chat" : "Open Chat"}
            >
              <MessageSquare size={18} strokeWidth={2} />
            </button>

            <button 
              onClick={() => setSidebarOpen(v => !v)} 
              style={{ 
                width: 38, height: 38, display: "flex", alignItems: "center", justifyContent: "center",
                background: sidebarOpen ? `${C.teal}10` : C.bg1, 
                border: `1px solid ${sidebarOpen ? C.teal : C.border}`, 
                borderRadius: 8, color: sidebarOpen ? C.teal : C.text2, cursor: "pointer", transition: "all 0.2s"
              }}
              title={sidebarOpen ? "Hide Sheets" : "Show Sheets"}
            >
              {sidebarOpen ? <ChevronRight size={18} strokeWidth={2.5} /> : <Layout size={18} strokeWidth={2} />}
            </button>
          </div>
        </div>

        {/* Sheet content */}
        <div style={{ flex: 1, overflow: "auto" }}>{renderSheet()}</div>

        {/* Import Modal */}
        {showImportModal && (
          <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 9999, backdropFilter: "blur(4px)" }}>
            <div style={{ width: 600, maxWidth: "90%", background: C.bg1, borderRadius: 16, boxShadow: "0 10px 30px rgba(0,0,0,0.1)", display: "flex", flexDirection: "column", overflow: "hidden" }}>
              <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ width: 36, height: 36, borderRadius: 10, background: "rgba(214,158,46,0.1)", display: "flex", alignItems: "center", justifyContent: "center", color: C.goldL }}><Sparkles size={20} strokeWidth={2} /></div>
                  <div>
                    <h3 style={{ margin: 0, fontSize: 16, color: C.text0 }}>Generate from Strategic Execution Plan</h3>
                    <div style={{ fontSize: 12, color: C.text2, marginTop: 2 }}>Paste your Business Validator strategy to auto-build the model.</div>
                  </div>
                </div>
                <button onClick={() => setShowImportModal(false)} style={{ background: "transparent", border: "none", cursor: "pointer", fontSize: 20, color: C.text3 }}>&times;</button>
              </div>
              <div style={{ padding: 24 }}>
                <textarea
                  value={importText}
                  onChange={e => setImportText(e.target.value)}
                  placeholder="Paste the Strategic Execution Plan text here..."
                  style={{ width: "100%", height: 300, padding: 16, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 13, fontFamily: "monospace", resize: "none", outline: "none", background: C.bg0, color: C.text0 }}
                />
              </div>
              <div style={{ padding: "16px 24px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "flex-end", gap: 12, background: C.bg0 }}>
                <button onClick={() => setShowImportModal(false)} style={{ padding: "10px 20px", borderRadius: 8, border: `1px solid ${C.border}`, background: C.bg1, color: C.text1, cursor: "pointer", fontWeight: 600, fontSize: 13 }}>Cancel</button>
                <button onClick={handleImportPlan} disabled={importLoading || !importText.trim()} style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 24px", borderRadius: 8, border: "none", background: "linear-gradient(135deg,#ECC94B,#D69E2E)", color: "#fff", cursor: importLoading ? "not-allowed" : "pointer", fontWeight: 600, fontSize: 13, opacity: importLoading ? 0.7 : 1 }}>
                  {importLoading ? <Activity size={16} className="animate-spin" /> : <Sparkles size={16} />}
                  {importLoading ? "Generating Framework..." : "Generate Financial Model"}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Status bar combined with Hint */}
        <div style={{ 
          display: "flex", 
          padding: "6px 16px", 
          background: C.bg0, 
          borderTop: `1px solid ${C.border}`, 
          flexShrink: 0, 
          alignItems: "center",
          justifyContent: "space-between"
        }}>
          <div style={{ fontSize: 10, color: C.text3, display: "flex", alignItems: "center", gap: 4 }}>
            <Activity size={10} />
            <span>Ready for analysis</span>
          </div>
          <div style={{ fontSize: 10, color: C.text3, fontWeight: 500 }}>
            💡 Tip: Click <span style={{ color: C.blue, fontWeight: 700 }}>blue cells</span> to edit · Use chat to automate
          </div>
        </div>
      </div>

      {sidebarOpen && <Sidebar />}
    </div>
  );
}
