/**
 * ScenarioDashboard.jsx
 * =====================
 * Scenario Modeling: Baseline / Optimistic / Pessimistic
 *
 * Runs the financial engine with 3 different growth rate profiles
 * and displays a side-by-side 5-year P&L comparison.
 *
 * Also wires real-time recalculation — any change to revP1/opexP1
 * in the parent propagates here and all three scenarios update instantly.
 */
'use client';

import { useMemo, useState } from 'react';

// ─── Color Theme (matches main component) ────────────────────────────────────
const C = {
    bg0: '#05091A', bg1: '#0A1020', bg2: '#0E1628', bg3: '#121E34',
    nav: '#0C1426', navB: '#162040',
    border: '#1A2B48', borderLight: '#243550',
    gold: '#C4972A', goldL: '#E8C96B', goldD: '#8A6A1A',
    teal: '#2A9E9E', tealL: '#3DBDBD',
    green: '#2EA870', greenL: '#3DCA87',
    red: '#C84040', redL: '#E85555',
    blue: '#3B78D4', blueL: '#5B9CF6',
    purple: '#7B52D4', purpleL: '#A67EF7',
    text0: '#E0EAF8', text1: '#8AAAC8', text2: '#4A6888', text3: '#2A3D58',
    sectionBg: '#0C1830', totalBg: '#091422', headerBg: '#101E38',
};

// ─── Scenario Configs ─────────────────────────────────────────────────────────
const SCENARIO_CONFIGS = {
    pessimistic: {
        label: 'Pessimistic',
        icon: '🔻',
        color: C.redL,
        accent: 'rgba(200,64,64,0.15)',
        border: 'rgba(200,64,64,0.4)',
        // Growth multiplier applied to base year-on-year rates
        growthMultiplier: 0.5,
        opexMultiplier: 1.2,   // OPEX grows 20% more
        description: 'Conservative growth, higher costs',
    },
    baseline: {
        label: 'Baseline',
        icon: '📊',
        color: C.gold,
        accent: 'rgba(196,151,42,0.12)',
        border: 'rgba(196,151,42,0.4)',
        growthMultiplier: 1.0,
        opexMultiplier: 1.0,
        description: 'Expected growth trajectory',
    },
    optimistic: {
        label: 'Optimistic',
        icon: '🚀',
        color: C.greenL,
        accent: 'rgba(46,168,112,0.12)',
        border: 'rgba(46,168,112,0.4)',
        growthMultiplier: 1.6,
        opexMultiplier: 0.85,  // OPEX grows 15% less (operational leverage)
        description: 'Strong growth, controlled costs',
    },
};

const YEARS = ['2026-27', '2027-28', '2028-29', '2029-30', '2030-31'];

// ─── Core Calculation Functions (pure, client-side) ──────────────────────────

function calcRevYearly(groups) {
    return groups.map(g => ({
        ...g,
        yearlyTotals: YEARS.map((_, yi) => {
            return g.items.reduce((sum, item) => {
                const base = (item.qty || 0) * (item.price || 0) * 12;
                if (!base) return sum;
                let v = base;
                for (let i = 0; i < yi; i++) v *= (1 + (item[`gY${i + 1}`] || 0));
                return sum + v;
            }, 0);
        }),
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
                for (let i = 0; i < yi; i++) v *= (1 + (item[`gY${i + 1}`] || 0));
                return sum + v;
            }, 0);
        }),
    }));
}

function applyScenarioMultiplier(groups, multiplier, isRev) {
    return groups.map(g => ({
        ...g,
        items: g.items.map(item => ({
            ...item,
            gY1: (item.gY1 || 0) * multiplier,
            gY2: (item.gY2 || 0) * multiplier,
            gY3: (item.gY3 || 0) * multiplier,
            gY4: (item.gY4 || 0) * multiplier,
            gY5: (item.gY5 || 0) * multiplier,
        })),
    }));
}

function calcScenario(revP1, opexP1, loan1, loan2, scenarioConfig) {
    const adjRev = applyScenarioMultiplier(revP1, scenarioConfig.growthMultiplier, true);
    const adjOpex = applyScenarioMultiplier(opexP1, scenarioConfig.opexMultiplier, false);

    const revByYear = YEARS.map((_, yi) =>
        calcRevYearly(adjRev).reduce((s, g) => s + g.yearlyTotals[yi], 0)
    );
    const opexByYear = YEARS.map((_, yi) =>
        calcOpexYearly(adjOpex).reduce((s, g) => s + g.yearlyTotals[yi], 0)
    );
    const ebitda = YEARS.map((_, yi) => revByYear[yi] - opexByYear[yi]);
    const ebitdaMargin = YEARS.map((_, yi) => revByYear[yi] > 0 ? ebitda[yi] / revByYear[yi] : 0);
    const deprn = YEARS.map(() => 75000 * 12);
    const ebit = YEARS.map((_, yi) => ebitda[yi] - deprn[yi]);
    
    // Dynamic interest calculation (matching main logic)
    const interest = YEARS.map((_, yi) => {
        if (yi === 1) return (Number(loan1?.amount) || 0) * (Number(loan1?.rate) || 0) / 100;
        if (yi === 2) return (Number(loan2?.amount) || 0) * (Number(loan2?.rate) || 0) / 100;
        return 0;
    });

    const pbt = YEARS.map((_, yi) => ebit[yi] - interest[yi]);
    const tax = YEARS.map((_, yi) => Math.max(0, pbt[yi] * 0.25));
    const pat = YEARS.map((_, yi) => pbt[yi] - tax[yi]);

    return { revByYear, opexByYear, ebitda, ebitdaMargin, pat };
}

// ─── Format Helpers ───────────────────────────────────────────────────────────
const fmtCr = (v) => {
    if (!v && v !== 0) return '—';
    const abs = Math.abs(v);
    const neg = v < 0;
    let s;
    if (abs >= 10000000) s = `₹${(abs / 10000000).toFixed(2)} Cr`;
    else if (abs >= 100000) s = `₹${(abs / 100000).toFixed(2)} L`;
    else s = `₹${abs.toLocaleString('en-IN', { maximumFractionDigits: 0 })}`;
    return neg ? `(${s})` : s;
};
const fmtPct = v => v == null ? '—' : `${(v * 100).toFixed(1)}%`;

// ─── Scenario Card ────────────────────────────────────────────────────────────
function ScenarioCard({ label, icon, color, accent, border, description, data, years }) {
    const { revByYear, opexByYear, ebitda, ebitdaMargin, pat } = data;
    const y1Rev = revByYear[0] || 0;
    const y5Rev = revByYear[4] || 0;
    const cagr = y1Rev > 0 && y5Rev > 0 ? Math.pow(y5Rev / y1Rev, 0.25) - 1 : 0;
    const profitableFrom = pat.findIndex(p => p > 0);

    return (
        <div style={{
            flex: 1, minWidth: 320,
            background: accent,
            border: `1px solid ${border}`,
            borderRadius: 12, overflow: 'hidden',
        }}>
            {/* Header */}
            <div style={{ padding: '14px 18px', borderBottom: `1px solid ${border}` }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                    <span style={{ fontSize: 20 }}>{icon}</span>
                    <div>
                        <div style={{ fontSize: 15, fontWeight: 700, color }}>{label}</div>
                        <div style={{ fontSize: 11, color: C.text2 }}>{description}</div>
                    </div>
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10, marginTop: 10 }}>
                    {[
                        ['Y1 Revenue', fmtCr(y1Rev)],
                        ['Y5 Revenue', fmtCr(y5Rev)],
                        ['CAGR', fmtPct(cagr)],
                        ['Y1 EBITDA', fmtCr(ebitda[0])],
                        ['Y1 Margin', fmtPct(ebitdaMargin[0])],
                        ['Breakeven', profitableFrom >= 0 ? `Y${profitableFrom + 1}` : 'Y5+'],
                    ].map(([k, v]) => (
                        <div key={k} style={{ padding: '6px 10px', background: 'rgba(0,0,0,0.25)', borderRadius: 6 }}>
                            <div style={{ fontSize: 10, color: C.text2, marginBottom: 2 }}>{k}</div>
                            <div style={{ fontSize: 13, fontWeight: 700, color, fontFamily: 'monospace' }}>{v}</div>
                        </div>
                    ))}
                </div>
            </div>

            {/* Year-by-year table */}
            <div style={{ overflowX: 'auto' }}>
                <table style={{ borderCollapse: 'collapse', width: '100%' }}>
                    <thead>
                        <tr style={{ background: 'rgba(0,0,0,0.3)' }}>
                            <th style={{ padding: '8px 12px', fontSize: 11, fontWeight: 700, color: C.text2, textAlign: 'left', whiteSpace: 'nowrap' }}>Metric</th>
                            {years.map(y => (
                                <th key={y} style={{ padding: '8px 10px', fontSize: 11, fontWeight: 700, color, textAlign: 'right', whiteSpace: 'nowrap' }}>{y}</th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {[
                            { label: 'Revenue', vals: revByYear, isBold: false },
                            { label: 'OPEX', vals: opexByYear.map(v => -v), isNeg: true },
                            { label: 'EBITDA', vals: ebitda, isBold: true, isTotal: true },
                            { label: 'EBITDA %', vals: ebitdaMargin, isPct: true },
                            { label: 'Net Profit', vals: pat, isBold: true },
                        ].map((row, i) => (
                            <tr key={i} style={{ background: row.isTotal ? 'rgba(0,0,0,0.3)' : i % 2 === 0 ? 'rgba(255,255,255,0.02)' : 'transparent' }}>
                                <td style={{ padding: '6px 12px', fontSize: 12, color: row.isTotal ? color : C.text1, fontWeight: row.isBold ? 700 : 400, whiteSpace: 'nowrap', borderBottom: `1px solid ${C.border}` }}>
                                    {row.label}
                                </td>
                                {row.vals.map((v, yi) => {
                                    const col = row.isPct
                                        ? (v >= 0 ? C.greenL : C.redL)
                                        : row.isTotal
                                            ? (v >= 0 ? color : C.redL)
                                            : (v < 0 ? C.redL : C.text1);
                                    return (
                                        <td key={yi} style={{ padding: '6px 10px', fontSize: 12, fontFamily: 'monospace', color: col, textAlign: 'right', fontWeight: row.isBold ? 700 : 400, borderBottom: `1px solid ${C.border}` }}>
                                            {row.isPct ? fmtPct(v) : fmtCr(v)}
                                        </td>
                                    );
                                })}
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </div>
    );
}

// ─── Revenue Waterfall Bar ─────────────────────────────────────────────────────
function WaterfallChart({ scenarios, years }) {
    const maxRev = Math.max(
        ...Object.values(scenarios).flatMap(s => s.revByYear)
    );

    return (
        <div style={{ padding: '20px 16px' }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: C.text0, marginBottom: 16 }}>Revenue Projection Comparison</div>
            <div style={{ display: 'flex', gap: 16, alignItems: 'flex-end', height: 200 }}>
                {years.map((year, yi) => (
                    <div key={year} style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4 }}>
                        <div style={{ display: 'flex', gap: 3, alignItems: 'flex-end', height: 160, width: '100%' }}>
                            {['pessimistic', 'baseline', 'optimistic'].map(key => {
                                const cfg = SCENARIO_CONFIGS[key];
                                const val = scenarios[key]?.revByYear[yi] || 0;
                                const heightPct = maxRev > 0 ? (val / maxRev) * 100 : 0;
                                return (
                                    <div key={key} title={`${cfg.label}: ${fmtCr(val)}`} style={{
                                        flex: 1, height: `${Math.max(2, heightPct)}%`,
                                        background: cfg.color,
                                        borderRadius: '3px 3px 0 0',
                                        opacity: 0.8,
                                        transition: 'height 0.4s ease',
                                        cursor: 'default',
                                        minHeight: 2,
                                    }} />
                                );
                            })}
                        </div>
                        <div style={{ fontSize: 10, color: C.text2, textAlign: 'center' }}>{year}</div>
                    </div>
                ))}
            </div>
            <div style={{ display: 'flex', gap: 16, justifyContent: 'center', marginTop: 12 }}>
                {Object.entries(SCENARIO_CONFIGS).map(([key, cfg]) => (
                    <div key={key} style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                        <div style={{ width: 10, height: 10, borderRadius: 2, background: cfg.color }} />
                        <span style={{ fontSize: 11, color: C.text1 }}>{cfg.label}</span>
                    </div>
                ))}
            </div>
        </div>
    );
}

// ─── Main ScenarioDashboard ────────────────────────────────────────────────────
export default function ScenarioDashboard({ revP1, opexP1, loan1, loan2 }) {
    const [activeScenario, setActiveScenario] = useState('all');

    // Real-time: recalculates whenever revP1 or opexP1 changes (no button needed)
    const scenarios = useMemo(() => {
        return Object.fromEntries(
            Object.entries(SCENARIO_CONFIGS).map(([key, cfg]) => [
                key,
                calcScenario(revP1, opexP1, loan1, loan2, cfg),
            ])
        );
    }, [revP1, opexP1, loan1, loan2]);

    const hasData = useMemo(() => {
        const totalRev = scenarios.baseline?.revByYear[0] || 0;
        return totalRev > 0;
    }, [scenarios]);

    return (
        <div style={{ background: C.bg0, minHeight: '100%' }}>
            {/* Sheet Header */}
            <div style={{ padding: '10px 16px', background: C.nav, borderBottom: `1px solid ${C.border}`, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <div>
                    <div style={{ fontSize: 14, fontWeight: 700, color: C.text0 }}>Scenario Analysis — 5-Year Projections</div>
                    <div style={{ fontSize: 11, color: C.text2, marginTop: 2 }}>
                        Pessimistic · Baseline · Optimistic — real-time, based on your Revenue & OPEX inputs
                    </div>
                </div>
                <div style={{ display: 'flex', gap: 6 }}>
                    {['all', 'pessimistic', 'baseline', 'optimistic'].map(key => {
                        const cfg = SCENARIO_CONFIGS[key];
                        const isAll = key === 'all';
                        return (
                            <button
                                key={key}
                                onClick={() => setActiveScenario(key)}
                                style={{
                                    padding: '4px 12px', fontSize: 11,
                                    background: activeScenario === key ? (isAll ? C.navB : cfg?.accent) : 'transparent',
                                    border: `1px solid ${activeScenario === key ? (isAll ? C.teal : cfg?.border) : C.border}`,
                                    borderRadius: 6,
                                    color: activeScenario === key ? (isAll ? C.teal : cfg?.color) : C.text2,
                                    cursor: 'pointer',
                                    transition: 'all 0.15s',
                                }}
                            >
                                {isAll ? 'All Scenarios' : `${cfg.icon} ${cfg.label}`}
                            </button>
                        );
                    })}
                </div>
            </div>

            {!hasData && (
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: 400, gap: 12, color: C.text2 }}>
                    <div style={{ fontSize: 40 }}>📊</div>
                    <div style={{ fontSize: 14, color: C.text1 }}>No data yet — enter revenue & OPEX first</div>
                    <div style={{ fontSize: 12 }}>Go to Revenue Streams (A.I) and OPEX (A.II) to add your numbers</div>
                </div>
            )}

            {hasData && (
                <>
                    {/* Waterfall chart */}
                    <div style={{ borderBottom: `1px solid ${C.border}`, background: C.bg1 }}>
                        <WaterfallChart scenarios={scenarios} years={YEARS} />
                    </div>

                    {/* Summary comparison table */}
                    <div style={{ overflowX: 'auto', padding: '16px', borderBottom: `1px solid ${C.border}` }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: C.text0, marginBottom: 12 }}>5-Year Net Profit Comparison (PAT)</div>
                        <table style={{ borderCollapse: 'collapse', width: '100%', minWidth: 600 }}>
                            <thead>
                                <tr style={{ background: C.headerBg }}>
                                    <th style={{ padding: '8px 14px', fontSize: 12, color: C.gold, fontWeight: 700, textAlign: 'left', whiteSpace: 'nowrap' }}>Scenario</th>
                                    {YEARS.map(y => (
                                        <th key={y} style={{ padding: '8px 12px', fontSize: 12, color: C.gold, fontWeight: 700, textAlign: 'right', whiteSpace: 'nowrap' }}>{y}</th>
                                    ))}
                                    <th style={{ padding: '8px 12px', fontSize: 12, color: C.gold, fontWeight: 700, textAlign: 'right' }}>5Y Total</th>
                                </tr>
                            </thead>
                            <tbody>
                                {Object.entries(SCENARIO_CONFIGS).map(([key, cfg], i) => {
                                    const s = scenarios[key];
                                    const total = s.pat.reduce((a, b) => a + b, 0);
                                    return (
                                        <tr key={key} style={{ background: i % 2 === 0 ? C.bg1 : C.bg0 }}>
                                            <td style={{ padding: '8px 14px', fontSize: 13, color: cfg.color, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>
                                                {cfg.icon} {cfg.label}
                                            </td>
                                            {s.pat.map((v, yi) => (
                                                <td key={yi} style={{ padding: '8px 12px', fontSize: 13, fontFamily: 'monospace', color: v >= 0 ? C.greenL : C.redL, textAlign: 'right', fontWeight: 600, borderBottom: `1px solid ${C.border}` }}>
                                                    {fmtCr(v)}
                                                </td>
                                            ))}
                                            <td style={{ padding: '8px 12px', fontSize: 13, fontFamily: 'monospace', color: total >= 0 ? cfg.color : C.redL, textAlign: 'right', fontWeight: 700, borderBottom: `1px solid ${C.border}`, borderLeft: `2px solid ${C.borderLight}` }}>
                                                {fmtCr(total)}
                                            </td>
                                        </tr>
                                    );
                                })}
                            </tbody>
                        </table>
                    </div>

                    {/* Scenario cards */}
                    <div style={{ padding: '16px', display: 'flex', gap: 16, flexWrap: 'wrap', overflowX: 'auto' }}>
                        {Object.entries(SCENARIO_CONFIGS)
                            .filter(([key]) => activeScenario === 'all' || activeScenario === key)
                            .map(([key, cfg]) => (
                                <ScenarioCard
                                    key={key}
                                    {...cfg}
                                    data={scenarios[key]}
                                    years={YEARS}
                                />
                            ))}
                    </div>
                </>
            )}
        </div>
    );
}
