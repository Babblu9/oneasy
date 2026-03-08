const xlsx = require('xlsx');

function listTabs() {
    const workbook = xlsx.readFile('/Users/vinaygogula/Downloads/srinivas/2026/OnEasy/finance_model/Docty Healthcare - Business Plan.xlsx');
    console.log("Tabs in Workbook:");
    workbook.SheetNames.forEach(name => console.log(`- ${name}`));
}

listTabs();
