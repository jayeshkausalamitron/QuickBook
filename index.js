const QuickBooks = require('node-quickbooks');
const fs = require('fs');
const XLSX = require('xlsx');

// Configure QuickBooks API
const qbo = new QuickBooks({
    consumerKey: process.env.CONSUMERKEY,
    consumerSecret: process.env.CONSUMERSECRET,
    token: process.env.TOKEN,
    tokenSecret: process.env.TOKENSECRET,
    realmId: process.env.REALMID,
    useSandbox: process.env.USESANDBOX // Set to true if you are using the sandbox environment
});

// Define the report name
const reportName = 'BalanceSheet'; // Change this to the name of the report you want to pull

// Pull a report from QuickBooks
qbo.report(reportName, (err, report) => {
    if (err) {
        console.error('Error pulling report:', err);
        return;
    }

    // Log the report data
    console.log(report);

    // Convert JSON data to worksheet
    const ws = XLSX.utils.json_to_sheet(report);

    // Create a workbook and add the worksheet
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    // Define the Excel file path
    const excelFilePath = 'output.xlsx';

    // Write the workbook to a file
    XLSX.writeFile(wb, excelFilePath, { bookType: 'xlsx', type: 'file' });

    // Log success message
    console.log(`Excel file "${excelFilePath}" created successfully.`);
});
