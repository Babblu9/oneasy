'use client';

import { useState, useRef, useEffect } from "react";
import { dataActionToPatches } from "@/lib/excelCellMap.js";
import { modelStateToPatches } from "@/lib/modelStateToPatches.js";
import ScenarioDashboard from "@/components/template/ScenarioDashboard";

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

// ─── INITIAL DATA ─────────────────────────────────────────────────────────────
const INIT = {
  basics: {
    legalName: "OnEasy Financial Services Pvt Ltd", tradeName: "OnEasy",
    address: "6-1-19 and 6-1-19/A, old no 108/3, Flat no 206, Musheerabad, Hyderabad - 500020, Telangana",
    email: "Karimsavenue@gmail.com", contact: "+91-99999-00000", promoters: "4",
    startDateP1: "2026-04-01", startDateP2: "2027-04-01",
    description: "Digital healthcare platform connecting patients to doctors for consultations, diagnostics, pharmacy and health packages across Hyderabad.",
    pitchDeck: "", burningDesire: "Make quality healthcare accessible and affordable for every Indian household.",
  },
  revP1: [
    {
      id: "1", header: "Doctor Consultations", items: [
        { id: "1a", sub: "Platform Access Fee (OPD)", qty: 28000, price: 30, gY1: 0.01, gY2: 0.82, gY3: 0.70, gY4: 0.55, gY5: 0.45 },
        { id: "1b", sub: "Telemedicine Consultation", qty: 3000, price: 200, gY1: 0.05, gY2: 0.90, gY3: 0.75, gY4: 0.55, gY5: 0.40 },
        { id: "1c", sub: "Specialist Referral Fee", qty: 800, price: 500, gY1: 0.03, gY2: 0.70, gY3: 0.60, gY4: 0.45, gY5: 0.35 },
        { id: "1d", sub: "Home Doctor Visit", qty: 200, price: 800, gY1: 0.02, gY2: 0.65, gY3: 0.55, gY4: 0.40, gY5: 0.30 },
        { id: "1e", sub: "Health Report Generation", qty: 5000, price: 50, gY1: 0.05, gY2: 0.80, gY3: 0.65, gY4: 0.50, gY5: 0.35 },
      ]
    },
    {
      id: "2", header: "Advertisement & Partnerships", items: [
        { id: "2a", sub: "Pharma Brand Ads", qty: 8, price: 50000, gY1: 0.05, gY2: 0.60, gY3: 0.55, gY4: 0.40, gY5: 0.30 },
        { id: "2b", sub: "Clinic / Lab Listing Fee", qty: 200, price: 5000, gY1: 0.08, gY2: 0.70, gY3: 0.60, gY4: 0.45, gY5: 0.35 },
        { id: "2c", sub: "Insurance Referral Fee", qty: 300, price: 1500, gY1: 0.05, gY2: 0.65, gY3: 0.55, gY4: 0.40, gY5: 0.30 },
        { id: "2d", sub: "Sponsored Health Posts", qty: 40, price: 2500, gY1: 0.03, gY2: 0.50, gY3: 0.45, gY4: 0.35, gY5: 0.25 },
        { id: "2e", sub: "Doctor Profile Boost", qty: 150, price: 1000, gY1: 0.05, gY2: 0.60, gY3: 0.50, gY4: 0.35, gY5: 0.25 },
      ]
    },
    {
      id: "3", header: "Pharmacy & Lab Diagnostics", items: [
        { id: "3a", sub: "Online Pharmacy Margin", qty: 5000, price: 180, gY1: 0.05, gY2: 0.75, gY3: 0.65, gY4: 0.50, gY5: 0.35 },
        { id: "3b", sub: "Lab Test Booking Fee", qty: 4000, price: 300, gY1: 0.06, gY2: 0.80, gY3: 0.70, gY4: 0.52, gY5: 0.38 },
        { id: "3c", sub: "Home Sample Collection", qty: 800, price: 150, gY1: 0.04, gY2: 0.65, gY3: 0.55, gY4: 0.40, gY5: 0.30 },
        { id: "3d", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "3e", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
      ]
    },
    {
      id: "4", header: "Health Membership Plans", items: [
        { id: "4a", sub: "Basic Plan (Rs.299/mo)", qty: 10000, price: 299, gY1: 0.10, gY2: 1.20, gY3: 0.90, gY4: 0.60, gY5: 0.40 },
        { id: "4b", sub: "Premium Plan (Rs.999/mo)", qty: 2000, price: 999, gY1: 0.08, gY2: 1.10, gY3: 0.85, gY4: 0.55, gY5: 0.38 },
        { id: "4c", sub: "Corporate Wellness Pack", qty: 15, price: 50000, gY1: 0.05, gY2: 0.80, gY3: 0.65, gY4: 0.50, gY5: 0.35 },
        { id: "4d", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "4e", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
      ]
    },
    {
      id: "5", header: "Data & AI Services (B2B)", items: [
        { id: "5a", sub: "Health Analytics Reports", qty: 5, price: 200000, gY1: 0, gY2: 0.50, gY3: 1.00, gY4: 0.60, gY5: 0.40 },
        { id: "5b", sub: "Hospital Dashboard Licence", qty: 20, price: 15000, gY1: 0, gY2: 0.80, gY3: 1.20, gY4: 0.70, gY5: 0.45 },
        { id: "5c", sub: "AI Diagnosis API", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0.50, gY4: 0.80, gY5: 0.60 },
        { id: "5d", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "5e", sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
      ]
    },
  ],
  revP2: [
    {
      id: "1", header: "Cakes", items: [
        { id: "1a", sub: "Cup Cakes", qtyDay: 5, price: 20, gY1: 0.01, gY2: 0.50, gY3: 0.40, gY4: 0.30, gY5: 0.20 },
        { id: "1b", sub: "Customised Cakes", qtyDay: 4, price: 20, gY1: 0.01, gY2: 0.50, gY3: 0.40, gY4: 0.30, gY5: 0.20 },
        { id: "1c", sub: "Standard Cake", qtyDay: 3, price: 20, gY1: 0.01, gY2: 0.50, gY3: 0.40, gY4: 0.30, gY5: 0.20 },
        { id: "1d", sub: "Cake Slices", qtyDay: 2, price: 20, gY1: 0.01, gY2: 0.50, gY3: 0.40, gY4: 0.30, gY5: 0.20 },
        { id: "1e", sub: "", qtyDay: 1, price: 20, gY1: 0.01, gY2: 0.50, gY3: 0.40, gY4: 0.30, gY5: 0.20 },
      ]
    },
    {
      id: "2", header: "Confectionary", items: [
        { id: "2a", sub: "Burgers", qtyDay: 5, price: 20, gY1: 0.01, gY2: 0.50, gY3: 0.40, gY4: 0.30, gY5: 0.20 },
        { id: "2b", sub: "Pizza", qtyDay: 4, price: 20, gY1: 0.01, gY2: 0.50, gY3: 0.40, gY4: 0.30, gY5: 0.20 },
        { id: "2c", sub: "Puffs", qtyDay: 3, price: 20, gY1: 0.01, gY2: 0.50, gY3: 0.40, gY4: 0.30, gY5: 0.20 },
        { id: "2d", sub: "Other 1", qtyDay: 2, price: 20, gY1: 0.01, gY2: 0.50, gY3: 0.40, gY4: 0.30, gY5: 0.20 },
        { id: "2e", sub: "", qtyDay: 1, price: 20, gY1: 0.01, gY2: 0.50, gY3: 0.40, gY4: 0.30, gY5: 0.20 },
      ]
    },
  ],
  opexP1: [
    {
      id: "1", header: "Technology & Product Development", items: [
        { id: "1a", sub: "App Development (outsourced)", qty: 1, cost: 250000, gY1: -0.20, gY2: 0.35, gY3: 2.50, gY4: 0.10, gY5: 0.10 },
        { id: "1b", sub: "Platform Maintenance (SaaS)", qty: 1, cost: 120000, gY1: -0.20, gY2: 0.35, gY3: 2.50, gY4: 0.10, gY5: 0.10 },
        { id: "1c", sub: "Cloud Hosting (AWS/GCP)", qty: 1, cost: 80000, gY1: 0.08, gY2: 0.40, gY3: 1.00, gY4: 0.30, gY5: 0.20 },
        { id: "1d", sub: "Cybersecurity & Compliance", qty: 1, cost: 40000, gY1: 0.05, gY2: 0.10, gY3: 0.15, gY4: 0.10, gY5: 0.10 },
        { id: "1e", sub: "SMS / Notification API", qty: 1, cost: 25000, gY1: 0.10, gY2: 0.80, gY3: 0.70, gY4: 0.50, gY5: 0.35 },
        { id: "1f", sub: "Third-party Integrations", qty: 1, cost: 15000, gY1: 0.05, gY2: 0.20, gY3: 0.30, gY4: 0.15, gY5: 0.10 },
        { id: "1g", sub: "QA & Testing", qty: 1, cost: 30000, gY1: -0.30, gY2: 0.20, gY3: 0.50, gY4: 0.10, gY5: 0.10 },
      ]
    },
    {
      id: "2", header: "Legal, Compliance & Professional Charges", items: [
        { id: "2a", sub: "CA / Legal Retainer", qty: 1, cost: 50000, gY1: -0.08, gY2: 0.02, gY3: 0.15, gY4: 0.10, gY5: 0.10 },
        { id: "2b", sub: "NABH / Telemedicine Licence", qty: 1, cost: 20000, gY1: 0, gY2: 0, gY3: 0.05, gY4: 0.05, gY5: 0.05 },
        { id: "2c", sub: "Data Privacy (IT Act)", qty: 1, cost: 15000, gY1: 0, gY2: 0.05, gY3: 0.08, gY4: 0.08, gY5: 0.08 },
        { id: "2d", sub: "Insurance (D&O, Cyber)", qty: 1, cost: 10000, gY1: 0.05, gY2: 0.10, gY3: 0.10, gY4: 0.10, gY5: 0.10 },
        { id: "2e", sub: "ROC / MCA Compliance", qty: 1, cost: 8000, gY1: 0, gY2: 0.05, gY3: 0.05, gY4: 0.05, gY5: 0.05 },
        { id: "2f", sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "2g", sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
      ]
    },
    {
      id: "3", header: "Utilities & Office", items: [
        { id: "3a", sub: "Office Rent (Hyderabad)", qty: 1, cost: 75000, gY1: -0.08, gY2: 0.10, gY3: 0.35, gY4: 0.10, gY5: 0.10 },
        { id: "3b", sub: "Office Maintenance", qty: 1, cost: 30000, gY1: 0, gY2: 0.05, gY3: 0.20, gY4: 0.08, gY5: 0.08 },
        { id: "3c", sub: "Power & Electricity", qty: 1, cost: 15000, gY1: 0, gY2: 0.08, gY3: 0.20, gY4: 0.10, gY5: 0.10 },
        { id: "3d", sub: "Broadband & Communication", qty: 1, cost: 8000, gY1: 0, gY2: 0.05, gY3: 0.10, gY4: 0.05, gY5: 0.05 },
        { id: "3e", sub: "Travel & Conveyance", qty: 1, cost: 20000, gY1: -0.10, gY2: 0.15, gY3: 0.30, gY4: 0.15, gY5: 0.10 },
        { id: "3f", sub: "Office Supplies", qty: 1, cost: 10000, gY1: -0.05, gY2: 0.05, gY3: 0.15, gY4: 0.08, gY5: 0.08 },
        { id: "3g", sub: "Miscellaneous", qty: 1, cost: 25000, gY1: -0.10, gY2: 0.05, gY3: 0.20, gY4: 0.10, gY5: 0.10 },
        { id: "3h", sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
        { id: "3i", sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 },
      ]
    },
    {
      id: "4", header: "People (Salaries & HR)", items: [
        { id: "4a", sub: "Tech & Engineering Team", qty: 4, cost: 80000, gY1: 0, gY2: 0.25, gY3: 0.20, gY4: 0.15, gY5: 0.12 },
        { id: "4b", sub: "Customer Support Team", qty: 3, cost: 25000, gY1: 0, gY2: 0.30, gY3: 0.25, gY4: 0.15, gY5: 0.12 },
        { id: "4c", sub: "Sales & BD Team", qty: 2, cost: 45000, gY1: 0, gY2: 0.50, gY3: 0.40, gY4: 0.25, gY5: 0.15 },
        { id: "4d", sub: "Operations & Admin", qty: 2, cost: 35000, gY1: 0, gY2: 0.20, gY3: 0.25, gY4: 0.12, gY5: 0.10 },
        { id: "4e", sub: "Field Agents (Part-time)", qty: 5, cost: 15000, gY1: 0.20, gY2: 0.60, gY3: 0.50, gY4: 0.30, gY5: 0.20 },
        { id: "4f", sub: "Founders (Stipend Y1)", qty: 4, cost: 50000, gY1: 0.50, gY2: 0.25, gY3: 0.25, gY4: 0.20, gY5: 0.20 },
      ]
    },
    {
      id: "5", header: "Marketing & Sales", items: [
        { id: "5a", sub: "Digital Marketing (Google/Meta)", qty: 1, cost: 150000, gY1: -0.10, gY2: 0.50, gY3: 0.80, gY4: 0.40, gY5: 0.25 },
        { id: "5b", sub: "Content & SEO", qty: 1, cost: 30000, gY1: 0, gY2: 0.30, gY3: 0.40, gY4: 0.20, gY5: 0.15 },
        { id: "5c", sub: "Influencer / Doctor Referrals", qty: 1, cost: 40000, gY1: 0.10, gY2: 0.60, gY3: 0.70, gY4: 0.35, gY5: 0.20 },
        { id: "5d", sub: "Events & Partnerships", qty: 1, cost: 20000, gY1: 0, gY2: 0.40, gY3: 0.50, gY4: 0.25, gY5: 0.15 },
        { id: "5e", sub: "Branding & Design", qty: 1, cost: 15000, gY1: -0.20, gY2: 0.10, gY3: 0.15, gY4: 0.08, gY5: 0.08 },
      ]
    },
  ],
  capex: [
    {
      sno: 1, category: "Office Equipment", items: [
        { name: "Laptops (x6)", total: 480000, y1: 480000, y2: 0, y3: 120000, y4: 0, y5: 0 },
        { name: "Mobiles (x4)", total: 80000, y1: 80000, y2: 0, y3: 0, y4: 0, y5: 0 },
        { name: "Monitors (x6)", total: 90000, y1: 90000, y2: 0, y3: 0, y4: 0, y5: 0 },
        { name: "Printer", total: 25000, y1: 25000, y2: 0, y3: 0, y4: 0, y5: 0 },
        { name: "Server (Local)", total: 150000, y1: 150000, y2: 0, y3: 0, y4: 0, y5: 0 },
      ]
    },
    {
      sno: 2, category: "Technology Infrastructure", items: [
        { name: "App Build (Initial)", total: 1500000, y1: 1500000, y2: 0, y3: 500000, y4: 0, y5: 0 },
        { name: "Website & CMS", total: 200000, y1: 200000, y2: 0, y3: 0, y4: 0, y5: 0 },
      ]
    },
    {
      sno: 3, category: "Interior & Setup", items: [
        { name: "Office Furnishing", total: 300000, y1: 300000, y2: 0, y3: 0, y4: 0, y5: 0 },
        { name: "Security & CCTV", total: 80000, y1: 80000, y2: 0, y3: 0, y4: 0, y5: 0 },
      ]
    },
    {
      sno: 4, category: "Branding & Legal Setup", items: [
        { name: "Trademark & IP", total: 50000, y1: 50000, y2: 0, y3: 0, y4: 0, y5: 0 },
      ]
    },
  ],
  totalProjectCost: { total: 3455000, promoterContrib: 691000, termLoan: 2000000, wcLoan: 764000 },
  loan1: { amount: 2000000, duration: 60, rate: 12, startDate: "2026-10-01" },
  loan2: { amount: 5000000, duration: 72, rate: 11.5, startDate: "2027-04-01" },
  fixedAssets: [
    { name: "Office Equipment", rate: 0.15, opening: 825000, addAbove: 0, addBelow: 0 },
    { name: "Technology Infrastructure", rate: 0.25, opening: 1700000, addAbove: 0, addBelow: 500000 },
    { name: "Interior & Furnishing", rate: 0.10, opening: 380000, addAbove: 0, addBelow: 0 },
    { name: "Leasehold Improvements", rate: 0.10, opening: 0, addAbove: 0, addBelow: 0 },
    { name: "Other Assets", rate: 0.15, opening: 50000, addAbove: 0, addBelow: 0 },
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
      <SheetHeader title="1. Basic Information — OnEasy Financial Model" sub="Company Details & Project Basics" />
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
    <div style={{ overflowX: "auto" }}>
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
    </div>
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
    <div style={{ overflowX: "auto" }}>
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
    </div>
  );
}

// ─── SHEET: CAPEX ─────────────────────────────────────────────────────────────
function Capex({ data, setData, onFocus }) {
  const updItem = (ci, ii, k, v) => setData(p => p.map((cat, ci2) => ci2 !== ci ? cat : { ...cat, items: cat.items.map((it, ii2) => ii2 !== ii ? it : { ...it, [k]: v }) }));
  const updCat = (ci, k, v) => setData(p => p.map((cat, ci2) => ci2 !== ci ? cat : { ...cat, [k]: v }));
  const yLabels = ["Year 1 (31/03/2027)", "Year 2 (31/03/2028)", "Year 3 (31/03/2029)", "Year 4 (31/03/2030)", "Year 5 (31/03/2031)"];
  return (
    <div style={{ overflowX: "auto" }}>
      <SheetHeader title="A.III CAPEX — Capital Expenditure" sub="Cost for each period ending — define by equipment category" />
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
    </div>
  );
}

// ─── SHEET: B.I SALES (monthly view) ─────────────────────────────────────────
function SalesSheet({ groups, phase }) {
  const computed = calcRevYearly(groups, false);
  return (
    <div style={{ overflowX: "auto" }}>
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
    </div>
  );
}

// ─── SHEET: TOTAL PROJECT COST ────────────────────────────────────────────────
function TotalProjectCost({ data, setData, onFocus }) {
  const { total, promoterContrib, termLoan, wcLoan } = data;
  return (
    <div>
      <SheetHeader title="2. Statement of Total Project Cost" sub="OnEasy Financial Model" />
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
function PL({ revP1, opexP1 }) {
  const revByYear = YEARS.map((_, yi) => calcRevYearly(revP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const opexByYear = YEARS.map((_, yi) => calcOpexYearly(opexP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const grossProfit = YEARS.map((_, yi) => revByYear[yi] - opexByYear[yi] * 0.3);
  const ebitda = YEARS.map((_, yi) => revByYear[yi] - opexByYear[yi]);
  const deprn = YEARS.map(() => 75000 * 12);
  const ebit = YEARS.map((_, yi) => ebitda[yi] - deprn[yi]);
  const interest = YEARS.map((_, yi) => [0, 3000000 * 12, 5775000 * 12, 0, 0][yi] || 0);
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
    <div style={{ overflowX: "auto" }}>
      <SheetHeader title="4. Profit & Loss Statement" sub="OnEasy — Projected Annual P&L" readOnly />
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
    </div>
  );
}

// ─── SHEET: BALANCE SHEET ────────────────────────────────────────────────────
function BalanceSheet({ revP1, opexP1 }) {
  const revByYear = YEARS.map((_, yi) => calcRevYearly(revP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const opexByYear = YEARS.map((_, yi) => calcOpexYearly(opexP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const pat = YEARS.map((_, yi) => {
    const ebitda = revByYear[yi] - opexByYear[yi];
    const ebit = ebitda - 75000 * 12;
    const pbt = ebit;
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
    <div style={{ overflowX: "auto" }}>
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
    </div>
  );
}

// ─── SHEET: RATIOS ────────────────────────────────────────────────────────────
function Ratios({ revP1, opexP1 }) {
  const revByYear = YEARS.map((_, yi) => calcRevYearly(revP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const opexByYear = YEARS.map((_, yi) => calcOpexYearly(opexP1).reduce((s, g) => s + g.yearlyTotals[yi], 0));
  const pat = YEARS.map((_, yi) => { const e = revByYear[yi] - opexByYear[yi] - 75000 * 12; return e - Math.max(0, e * 0.25); });
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
    <div style={{ overflowX: "auto" }}>
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
    </div>
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
    <div style={{ overflowX: "auto" }}>
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
    </div>
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
    <div style={{ overflowX: "auto" }}>
      <SheetHeader title="FA Schedule — Fixed Assets Schedule" sub="OnEasy — Schedule - 01" />
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
    </div>
  );
}

// ─── SHEET: 3B CAPITAL COSTING ────────────────────────────────────────────────
function CapitalCosting({ data, setData, onFocus }) {
  const updItem = (ci, ii, k, v) => setData(p => p.map((cat, ci2) => ci2 !== ci ? cat : { ...cat, items: cat.items.map((it, ii2) => ii2 !== ii ? it : { ...it, [k]: v }) }));
  const yLabels = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"];
  return (
    <div style={{ overflowX: "auto" }}>
      <SheetHeader title="3b. Costing — Capital Expenditure" sub="Cost for each period ending" />
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
    </div>
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
const AI_SYSTEM = `You are the OnEasy Financial Strategist, a professional AI assistant specialized in building investor-grade financial models.

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
  const endRef = useRef(null);
  useEffect(() => { endRef.current?.scrollIntoView({ behavior: "smooth" }); }, [msgs]);

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
      const bizName = String(d?.basics?.tradeName || d?.basics?.legalName || "OnEasy").replace(/[^a-zA-Z0-9]/g, "_").slice(0, 30);
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
    const qty = Math.max(0, Number(action.units) || 0);
    const price = Math.max(0, Number(action.price ?? action.value) || 0);

    console.log("[DEBUG] applyRevenueAction:", { streamName, productName, qty, price });

    const next = { ...prev, revP1: prev.revP1.map(g => ({ ...g, items: g.items.map(it => ({ ...it })) })) };

    // Try to find existing group or empty group
    let targetGroup = next.revP1.find(g => String(g.header || "").toLowerCase() === streamName.toLowerCase());
    if (!targetGroup) targetGroup = next.revP1.find(g => !String(g.header || "").trim());
    if (!targetGroup && next.revP1.length > 0) targetGroup = next.revP1[0];

    // If still no group, create new one
    if (!targetGroup) {
      targetGroup = { id: String(next.revP1.length + 1), header: streamName, items: Array(5).fill(null).map((_, i) => ({ id: `${next.revP1.length + 1}${String.fromCharCode(97 + i)}`, sub: "", qty: 0, price: 0, gY1: 0.01, gY2: 0.82, gY3: 0.70, gY4: 0.55, gY5: 0.45 })) };
      next.revP1.push(targetGroup);
    }

    if (!targetGroup.header) targetGroup.header = streamName;

    // Find empty item slot
    let targetItem = targetGroup.items.find(it => !String(it.sub || "").trim());
    if (!targetItem) targetItem = targetGroup.items[0];

    if (targetItem) {
      targetItem.sub = productName || subName;
      targetItem.qty = qty;
      targetItem.price = price;
    }
    console.log("[DEBUG] Updated revP1:", JSON.stringify(next.revP1).substring(0, 200));
    return next;
  };

  const applyOpexAction = (prev, action) => {
    const cat = String(action.category || "Expense").trim();
    const sub = String(action.subCategory || action.label || "General").trim();
    const units = Math.max(0, Number(action.units) || 1);
    const cost = Math.max(0, Number(action.price ?? action.value) || 0);

    console.log("[DEBUG] applyOpexAction:", { cat, sub, units, cost });

    const next = { ...prev, opexP1: prev.opexP1.map(g => ({ ...g, items: g.items.map(it => ({ ...it })) })) };

    let targetGroup = next.opexP1.find(g => String(g.header || "").toLowerCase() === cat.toLowerCase());
    if (!targetGroup) targetGroup = next.opexP1.find(g => !String(g.header || "").trim());
    if (!targetGroup && next.opexP1.length > 0) targetGroup = next.opexP1[0];

    if (!targetGroup) {
      targetGroup = { id: String(next.opexP1.length + 1), header: cat, items: Array(5).fill(null).map((_, i) => ({ id: `${next.opexP1.length + 1}${String.fromCharCode(97 + i)}`, sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 })) };
      next.opexP1.push(targetGroup);
    }

    if (!targetGroup.header) targetGroup.header = cat;

    let targetItem = targetGroup.items.find(it => !String(it.sub || "").trim());
    if (!targetItem) targetItem = targetGroup.items[0];

    if (targetItem) {
      targetItem.sub = sub;
      targetItem.qty = units;
      targetItem.cost = cost;
    }
    console.log("[DEBUG] Updated opexP1:", JSON.stringify(next.opexP1).substring(0, 200));
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
    const preset = MODEL_DB[modelKey] || MODEL_DB.consulting;
    const list = (Array.isArray(products) && products.length ? products : preset.products).slice(0, 20);
    const next = INIT.revP1.map(g => ({
      ...g,
      header: "",
      items: g.items.map(it => ({ ...it, sub: "", qty: 0, price: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 })),
    }));
    const headers = ["Primary Revenue", "Secondary Revenue", "Other Revenue", "", ""];
    headers.forEach((h, i) => { if (next[i]) next[i].header = h; });
    list.forEach((name, idx) => {
      const gi = Math.floor(idx / 5);
      const ii = idx % 5;
      if (!next[gi] || !next[gi].items[ii]) return;
      next[gi].items[ii].sub = String(name).trim();
      next[gi].items[ii].qty = 1;
      next[gi].items[ii].price = 1000;
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
    const next = buildRevenueFromStreams(streams);
    streams.forEach((streamName, si) => {
      if (!next[si]) return;
      const subs = subMap[String(streamName || "").toLowerCase()] || [];
      subs.slice(0, 5).forEach((sub, ii) => {
        if (!next[si].items[ii]) return;
        next[si].items[ii].sub = String(sub).trim();
        next[si].items[ii].qty = 1;
        next[si].items[ii].price = 1000;
      });
    });
    return next;
  };

  const buildCleanOpexState = (modelKey, costs = []) => {
    const preset = MODEL_DB[modelKey] || MODEL_DB.consulting;
    const list = (Array.isArray(costs) && costs.length ? costs : preset.opex).slice(0, 20);
    const next = INIT.opexP1.map(g => ({
      ...g,
      header: "",
      items: g.items.map(it => ({ ...it, sub: "", qty: 0, cost: 0, gY1: 0, gY2: 0, gY3: 0, gY4: 0, gY5: 0 })),
    }));
    const headers = ["Operating Expense", "People Cost", "Admin Cost", "", ""];
    headers.forEach((h, i) => { if (next[i]) next[i].header = h; });
    list.forEach((name, idx) => {
      const gi = Math.floor(idx / 7);
      const ii = idx % 7;
      if (!next[gi] || !next[gi].items[ii]) return;
      next[gi].items[ii].sub = String(name).trim();
      next[gi].items[ii].qty = 0;
      next[gi].items[ii].cost = 0;
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

  const SHEETS = [
    { id: "1. Basics", label: "1. Basics" },
    { id: "A. Data Needed", label: "A. Data Needed" },
    { id: "A.I Revenue Streams - P1", label: "A.I Rev. P1" },
    { id: "A.I Revenue Streams - P2", label: "A.I Rev. P2" },
    { id: "A.IIOPEX", label: "A.II OPEX" },
    { id: "A.III CAPEX", label: "A.III CAPEX" },
    { id: "B.I Sales - P1", label: "B.I Sales P1" },
    { id: "B.I Sales - P2", label: "B.I Sales P2" },
    { id: "B.II - OPEX", label: "B.II OPEX" },
    { id: "2.Total Project Cost", label: "2. Project Cost" },
    { id: "B.II OPEX - P1", label: "B.II OPEX P1" },
    { id: "4. P&L", label: "4. P&L" },
    { id: "3b. Costing - (Capital Exp)", label: "3b. Costing" },
    { id: "5. Balance sheet", label: "5. Balance Sheet" },
    { id: "6. Ratios", label: "6. Ratios" },
    { id: "DSCR", label: "DSCR" },
    { id: "Repayment schedule", label: "Repayment" },
    { id: "FA Schedule", label: "FA Schedule" },
    { id: "Phase 1", label: "Phase 1 Loan" },
    { id: "Phase 2", label: "Phase 2 Loan" },
    { id: "Costing", label: "Costing" },
    { id: "Scenarios", label: "🎯 Scenarios" },
  ];

  const send = async () => {
    const text = input.trim();
    if (!text || loading) return;
    setInput(""); setLoading(true);
    const nextMsgs = [...msgs, { role: "user", text }];
    setMsgs(nextMsgs);
    if (/start fresh|new data|reset/i.test(text)) {
      setBusinessModel("consulting");
      setD(buildEmptyState());
      setSheet("1. Basics");
      console.log("[AGENT] Action: Reset flow (Static)");
      setMsgs(p => [...p, { role: "assistant", text: "Of course! Let's start with a clean slate. Tell me about your business vision, and we'll build the model together from scratch." }]);
      setLoading(false);
      return;
    }
    try {
      console.log("[AGENT] Action: Model refinement (LLM)");
      const res = await fetch("/api/chat-simple", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          system: AI_SYSTEM,
          messages: nextMsgs,
          prompt: `Current data snapshot:\n${JSON.stringify(d)}\n\nLatest user request:\n${text}`
        })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data?.error || "LLM request failed");
      const raw = data?.text || "{}";
      console.log("[DEBUG] Raw LLM response:", raw.substring(0, 500));
      const tags = extractDataTags(raw);
      console.log("[DEBUG] Extracted tags:", tags);
      if (tags.length) {
        for (const tag of tags) {
          console.log("[DEBUG] Applying action:", tag);
          // eslint-disable-next-line no-await-in-loop
          await applyDataAction(tag);
        }
      }
      const suggestions = extractSuggestions(raw);
      if (suggestions && Array.isArray(suggestions) && suggestions.length > 0) {
        setLlmSuggestions(suggestions);
      }
      let parsed;
      try { parsed = JSON.parse(raw.replace(/```json|```/g, "").trim()); } catch { parsed = { message: raw, changes: null }; }
      if (parsed.changes) {
        setD(prev => ({
          basics: parsed.changes.basics ? { ...prev.basics, ...parsed.changes.basics } : prev.basics,
          revP1: parsed.changes.revP1 || prev.revP1,
          revP2: parsed.changes.revP2 || prev.revP2,
          opexP1: parsed.changes.opexP1 || prev.opexP1,
          capex: parsed.changes.capex || prev.capex,
          totalProjectCost: parsed.changes.totalProjectCost ? { ...prev.totalProjectCost, ...parsed.changes.totalProjectCost } : prev.totalProjectCost,
          loan1: parsed.changes.loan1 ? { ...prev.loan1, ...parsed.changes.loan1 } : prev.loan1,
          loan2: parsed.changes.loan2 ? { ...prev.loan2, ...parsed.changes.loan2 } : prev.loan2,
          fixedAssets: parsed.changes.fixedAssets || prev.fixedAssets,
        }));
      }
      setMsgs(p => [...p, { role: "assistant", text: cleanForDisplay(parsed.message || raw || "Done!") }]);
    } catch (e) {
      setMsgs(p => [...p, { role: "assistant", text: `I'm sorry, I hit a small snag while processing that: ${(e && e.message) ? e.message : "something went wrong on my end."} Should we try again, or perhaps look at another part of the model?` }]);
    }
    setLoading(false);
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
      case "4. P&L": return <PL revP1={d.revP1} opexP1={d.opexP1} />;
      case "3b. Costing - (Capital Exp)": return <CapitalCosting data={d.capex} setData={v => setD(p => ({ ...p, capex: typeof v === "function" ? v(p.capex) : v }))} onFocus={handleCellFocus} />;
      case "5. Balance sheet": return <BalanceSheet revP1={d.revP1} opexP1={d.opexP1} />;
      case "6. Ratios": return <Ratios revP1={d.revP1} opexP1={d.opexP1} />;
      case "DSCR": return <DSCR revP1={d.revP1} opexP1={d.opexP1} />;
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
    <div style={{ display: "flex", height: "100vh", background: C.bg0, fontFamily: "'Inter', sans-serif", overflow: "hidden" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        *{box-sizing:border-box;} ::-webkit-scrollbar{width:4px;height:4px;} ::-webkit-scrollbar-track{background:${C.bg0};} ::-webkit-scrollbar-thumb{background:${C.navB};border-radius:2px;}
        input::placeholder,textarea::placeholder{color:${C.text3};} textarea{resize:none;} input:focus,textarea:focus{outline:none;}
        tr:hover td{background:rgba(255,255,255,0.015)!important;}
      `}</style>

      {/* MAIN AREA */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden", minWidth: 0 }}>
        {/* Top bar */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 14px", height: 42, background: C.nav, borderBottom: `1px solid ${C.border}`, flexShrink: 0 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 26, height: 26, borderRadius: 6, background: "linear-gradient(135deg,#2A9E9E,#2EA870)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13 }}>🏥</div>
            <div>
              <div style={{ fontSize: 13, fontWeight: 700, color: C.text0, lineHeight: 1 }}>OnEasy Financial Model</div>
              <div style={{ fontSize: 9, color: C.text2, textTransform: "uppercase", letterSpacing: "0.07em" }}>OnEasy · Hyderabad · {SHEETS.length} Sheets</div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
            <button
              onClick={handleDownloadExcel}
              disabled={downloading}
              style={{
                padding: "4px 12px",
                fontSize: 11,
                background: "linear-gradient(135deg,#0f9f52,#18c463)",
                border: "1px solid rgba(24,196,99,0.45)",
                borderRadius: 5,
                color: "#eaffef",
                cursor: downloading ? "not-allowed" : "pointer",
                opacity: downloading ? 0.7 : 1,
                fontWeight: 700
              }}
            >
              {downloading ? "Downloading..." : "Download Excel"}
            </button>
            <div style={{ fontSize: 10, padding: "2px 8px", background: "rgba(46,168,112,0.1)", border: "1px solid rgba(46,168,112,0.3)", borderRadius: 4, color: C.greenL }}>● Live Calculation</div>
            <button onClick={() => setChatOpen(v => !v)} style={{ padding: "4px 10px", fontSize: 11, background: C.navB, border: `1px solid ${C.border}`, borderRadius: 5, color: C.text1, cursor: "pointer" }}>{chatOpen ? "Hide ▶" : "◀ Chat"}</button>
          </div>
        </div>

        {/* Sheet tabs — scrollable */}
        <div style={{ display: "flex", gap: 1, padding: "5px 10px 0", background: C.nav, borderBottom: `1px solid ${C.border}`, overflowX: "auto", flexShrink: 0, scrollbarWidth: "none" }}>
          {SHEETS.map(s => (
            <button key={s.id} onClick={() => setSheet(s.id)} style={{ padding: "4px 11px", fontSize: 10, fontWeight: 600, background: sheet === s.id ? C.bg0 : "transparent", color: sheet === s.id ? C.gold : C.text2, border: "none", borderRadius: "3px 3px 0 0", cursor: "pointer", borderBottom: sheet === s.id ? `2px solid ${C.gold}` : "2px solid transparent", whiteSpace: "nowrap", flexShrink: 0 }}>
              {s.label}
            </button>
          ))}
        </div>

        {/* Sheet content */}
        <div style={{ flex: 1, overflow: "auto" }}>{renderSheet()}</div>

        {/* Status bar */}
        <div style={{ display: "flex", gap: 18, padding: "3px 14px", background: C.nav, borderTop: `1px solid ${C.border}`, flexShrink: 0, alignItems: "center" }}>
          {[["Rev Y1", fmtINR(revY1, true), C.tealL], ["OPEX Y1", fmtINR(opexY1, true), C.redL], ["EBITDA Y1", fmtINR(revY1 - opexY1, true), revY1 > opexY1 ? C.greenL : C.redL]].map(([k, v, col]) => (
            <div key={k} style={{ display: "flex", gap: 6 }}>
              <span style={{ fontSize: 10, color: C.text2 }}>{k}:</span>
              <span style={{ fontSize: 10, color: col, fontFamily: "monospace", fontWeight: 700 }}>{v}</span>
            </div>
          ))}
          <div style={{ marginLeft: "auto", fontSize: 10, color: C.text3 }}>💡 Click <span style={{ color: C.inputBlue }}>blue</span> cells to edit · Use chat for bulk changes</div>
        </div>
      </div>

      {/* CHAT PANEL */}
      {chatOpen && (
        <div style={{ width: 340, flexShrink: 0, display: "flex", flexDirection: "column", background: C.bg1, borderLeft: `1px solid ${C.border}` }}>
          <div style={{ padding: "10px 14px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ width: 26, height: 26, borderRadius: 7, background: "linear-gradient(135deg,#1a4db5,#3b78d4)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13 }}>🤖</div>
            <div>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.text0 }}>AI Model Editor</div>
              <div style={{ fontSize: 9, color: C.teal }}>Edit any sheet with natural language</div>
            </div>
          </div>
          <div style={{ flex: 1, overflowY: "auto", padding: 10 }}>
            {msgs.map((m, i) => (
              <div key={i} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start", marginBottom: 8 }}>
                <div style={{ maxWidth: "90%", padding: "8px 11px", borderRadius: m.role === "user" ? "10px 10px 2px 10px" : "2px 10px 10px 10px", background: m.role === "user" ? C.navB : "#0D1E38", border: `1px solid ${m.role === "user" ? C.borderLight : C.border}`, fontSize: 11, lineHeight: 1.6, color: m.role === "user" ? C.text0 : C.text1, whiteSpace: "pre-wrap" }}>
                  {m.text.split(/(\*\*.*?\*\*)/).map((p, j) => p.startsWith("**") && p.endsWith("**") ? <strong key={j} style={{ color: C.text0 }}>{p.slice(2, -2)}</strong> : p)}
                </div>
              </div>
            ))}
            {loading && (
              <div style={{ display: "flex", justifyContent: "flex-start", marginBottom: 8 }}>
                <div style={{ padding: "8px 13px", borderRadius: "2px 10px 10px 10px", background: "#0D1E38", border: `1px solid ${C.border}`, display: "flex", gap: 4 }}>
                  {[0, 1, 2].map(i => <div key={i} style={{ width: 5, height: 5, borderRadius: "50%", background: C.teal, animation: `p 1.2s ${i * 0.2}s infinite` }} />)}
                  <style>{`@keyframes p{0%,80%,100%{opacity:.3;transform:scale(.8)}40%{opacity:1;transform:scale(1)}}`}</style>
                </div>
              </div>
            )}
            <div ref={endRef} />
          </div>
          <div style={{ padding: "4px 8px 6px", display: "flex", flexWrap: "wrap", gap: 4 }}>
            <div style={{ display: "flex", width: "100%", gap: 6, marginBottom: 4 }}>
              <select
                value={selectedSuggestion}
                onChange={(e) => {
                  setSelectedSuggestion(e.target.value);
                  if (e.target.value) setInput(e.target.value);
                }}
                style={{ flex: 1, background: C.bg0, color: C.text0, border: `1px solid ${C.teal}`, borderRadius: 8, padding: "6px 8px", fontSize: 11 }}
              >
                <option value="">💡 AI Suggestions...</option>
                {llmSuggestions.map((s, i) => <option key={i} value={s}>{s}</option>)}
              </select>
            </div>
            <div style={{ display: "flex", width: "100%", background: C.bg0, border: `1px solid ${C.borderLight}`, borderRadius: 8, padding: 4 }}>
              <input
                type="text"
                value={input}
                onChange={e => setInput(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter") send(); }}
                placeholder="Ask AI to change data e.g., 'Increase price of consultations to 500'"
                style={{
                  flex: 1, background: "transparent", border: "none", color: C.text0, padding: "8px 12px", fontSize: 12
                }}
              />
              <button
                onClick={send}
                disabled={loading}
                style={{
                  background: loading ? C.nav : "linear-gradient(135deg,#3b78d4,#1a4db5)",
                  border: "none", borderRadius: 6, color: "white", padding: "0 14px", fontWeight: 600, fontSize: 13, cursor: loading ? "not-allowed" : "pointer",
                  display: "flex", alignItems: "center", justifyContent: "center"
                }}
              >
                {loading ? "..." : "↑"}
              </button>
            </div>
          </div>
          <div style={{ padding: "8px 10px", borderTop: `1px solid ${C.border}` }}>
            {cellSuggestion && (
              <div style={{ marginBottom: 10, padding: 10, background: "#0D1E38", border: `1px solid ${C.teal}`, borderRadius: 8 }}>
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
            <div style={{ display: "flex", gap: 6, background: C.bg0, border: `1px solid ${C.border}`, borderRadius: 8, padding: "6px 8px 6px 10px", alignItems: "flex-end" }}>
              <textarea value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }} placeholder="Edit model or ask a question..." rows={2} style={{ flex: 1, background: "transparent", border: "none", color: C.text0, fontSize: 11, lineHeight: 1.5, fontFamily: "Inter,sans-serif", maxHeight: 70, overflowY: "auto" }} />
              <button onClick={send} disabled={loading || !input.trim()} style={{ width: 28, height: 28, borderRadius: 7, background: loading || !input.trim() ? C.border : "linear-gradient(135deg,#1a4db5,#3b78d4)", border: "none", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none"><path d="M22 2L11 13M22 2L15 22L11 13M22 2L2 9L11 13" stroke="white" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" /></svg>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
