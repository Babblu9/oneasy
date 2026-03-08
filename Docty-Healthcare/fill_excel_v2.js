const ExcelJS = require('exceljs');
const path = require('path');

async function fillExcel(branches, product, units, price, outputPath) {
    const templatePath = path.join(process.cwd(), 'Docty Healthcare - Business Plan.xlsx');
    const workbook = new ExcelJS.Workbook();

    console.log('Reading workbook...');
    await workbook.xlsx.readFile(templatePath);

    // Convert to number for internal use but we will use formulas in the sheet
    const branchCount = parseInt(branches);

    // 1. MASTER INPUT SHEET: A.I Revenue Streams - P1
    const sheetInputs = workbook.getWorksheet('A.I Revenue Streams - P1');
    if (sheetInputs) {
        // H7 is our master branch count cell
        sheetInputs.getCell('H7').value = branchCount;

        // Row 10: Primary Product (Fixed values as these are the inputs)
        sheetInputs.getCell('F10').value = String(product);
        sheetInputs.getCell('H10').value = parseFloat(units);
        sheetInputs.getCell('J10').value = parseFloat(price);

        // Rows 11-204: Clear other product inputs
        for (let i = 11; i <= 204; i++) {
            sheetInputs.getCell(`F${i}`).value = '';
            sheetInputs.getCell(`H${i}`).value = '';
            sheetInputs.getCell(`J${i}`).value = '';
        }
    }

    // 2. BRANCH SHEET: Link schedule to Master Input
    const sheetBranch = workbook.getWorksheet('Branch');
    if (sheetBranch) {
        // Formulate a cross-sheet reference
        const formula = "'A.I Revenue Streams - P1'!$H$7";
        for (let i = 4; i <= 75; i++) {
            sheetBranch.getCell(`E${i}`).value = { formula: formula };
        }
    }

    // 3. OPEX SHEET: Link multiplier to Master Input
    const sheetOpex = workbook.getWorksheet('A.IIOPEX');
    if (sheetOpex) {
        // I8 controls scaling for Staff and Rent
        sheetOpex.getCell('I8').value = { formula: "'A.I Revenue Streams - P1'!$H$7" };
    }

    // 4. SALES CALC SHEETS: Link hardcoded multipliers to Master Input
    const salesSheets = ['B.I Sales - P1', 'B.I Sales - P1 (2)', 'B.I Sales - P2'];
    salesSheets.forEach(name => {
        const ws = workbook.getWorksheet(name);
        if (ws) {
            // K2 is often used as the Month 1 base multiplier
            ws.getCell('K2').value = { formula: "'A.I Revenue Streams - P1'!$H$7" };
        }
    });

    // 5. Clear P2 Inputs
    const sheetInputs2 = workbook.getWorksheet('A.I Revenue Streams - P2');
    if (sheetInputs2) {
        for (let i = 10; i <= 200; i++) {
            sheetInputs2.getCell(`F${i}`).value = '';
            sheetInputs2.getCell(`H${i}`).value = '';
            sheetInputs2.getCell(`J${i}`).value = '';
        }
    }

    console.log('Writing workbook...');
    await workbook.xlsx.writeFile(outputPath);
    console.log(`Successfully saved to ${outputPath}`);
}

const args = process.argv.slice(2);
if (args.length < 5) {
    console.log('Usage: node fill_excel_v2.js <branches> <product> <units> <price> <outputPath>');
    process.exit(1);
}

fillExcel(args[0], args[1], args[2], args[3], args[4])
    .catch(err => {
        console.error(err);
        process.exit(1);
    });
