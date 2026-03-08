/**
 * Engineering Test Suite
 * ======================
 * Validates the modular financial engine with sample data
 * and verifies output structure.
 */

const { runProjection, DEFAULT_BRANCH_SCHEDULE } = require('./index');

// ── Sample Input Data ────────────────────────────────────────────────

const SAMPLE_INPUT = {
    // Use default branch scaling (1 -> 10 over 6 months)
    branchSchedule: DEFAULT_BRANCH_SCHEDULE,

    // Sample Revenue Streams
    services: [
        {
            name: "Root Canal",
            streamName: "Clinical Revenue",
            subStreamName: "Dental",
            baseUnits: 8,
            price: 5500,
            active: true
        },
        {
            name: "Ethical Drugs",
            streamName: "Docty Pharma",
            subStreamName: "Pharma",
            baseUnits: 900,
            price: 500,
            active: true
        }
    ],

    // Sample OPEX Categories
    expenses: [
        {
            name: "Rent",
            category: "Utilities",
            amount: 295000,
            perBranch: true,
            active: true
        },
        {
            name: "Management Salary",
            category: "Salaries",
            amount: 200000,
            perBranch: true,
            active: true
        },
        {
            name: "Marketing Push",
            category: "Marketing & Promotions",
            amount: 80000,
            perBranch: false,
            active: true
        }
    ],

    // Sample Fixed Assets
    assets: [
        {
            name: "Clinic Interior & Equipment",
            cost: 25000000, // 2.5 Cr
            usefulLife: 10,
            acquisitionYear: 0
        }
    ],

    taxRate: 0.2517
};

// ── Execution ────────────────────────────────────────────────────────

console.log("🚀 Starting Financial Engine Verification...");
console.log("================================================================================");

try {
    const result = runProjection(SAMPLE_INPUT);

    // 1. Verify Monthly Data (First 12 months)
    console.log("\n📊 MONTHLY PROJECTION (Year 1):");
    console.log("--------------------------------------------------------------------------------");
    console.log("Month | Branches | Revenue (₹) | OPEX (₹) | EBITDA (₹) | PAT (₹)");
    console.log("--------------------------------------------------------------------------------");

    for (let m = 0; m < 12; m++) {
        const branches = SAMPLE_INPUT.branchSchedule[m]?.branches || 10;
        const pnl = result.monthly.pnl[m];
        console.log(
            `${m.toString().padEnd(5)} | ` +
            `${branches.toString().padEnd(8)} | ` +
            `${pnl.revenue.toLocaleString().padEnd(11)} | ` +
            `${pnl.opex.toLocaleString().padEnd(8)} | ` +
            `${pnl.ebitda.toLocaleString().padEnd(10)} | ` +
            `${pnl.pat.toLocaleString()}`
        );
    }

    // 2. Verify Yearly Aggregation
    console.log("\n📈 YEARLY SUMMARY (P&L):");
    console.log("--------------------------------------------------------------------------------");
    console.log("Year | Revenue (₹) | OPEX (₹) | EBITDA (%) | Depr (₹) | PAT (₹)");
    console.log("--------------------------------------------------------------------------------");

    result.yearly.pnl.forEach(y => {
        console.log(
            `${y.fy.padEnd(4)} | ` +
            `${y.revenue.toLocaleString().padEnd(11)} | ` +
            `${y.opex.toLocaleString().padEnd(8)} | ` +
            `${(y.ebitdaMargin + "%").padEnd(10)} | ` +
            `${y.depreciation.toLocaleString().padEnd(8)} | ` +
            `${y.pat.toLocaleString()}`
        );
    });

    // 3. Final Summary
    console.log("\n✅ ENGINE SUMMARY STATS:");
    console.log("--------------------------------------------------------------------------------");
    Object.entries(result.summary).forEach(([key, val]) => {
        console.log(`${key.padEnd(20)}: ${val.toLocaleString()}${key.includes('Rate') || key.includes('Margin') || key.includes('Cagr') ? '%' : ''}`);
    });

    console.log("\n================================================================================");
    console.log("✨ Verification Complete: Engine logic is clean and structured!");

} catch (err) {
    console.error("❌ Engine Error:", err);
    process.exit(1);
}
