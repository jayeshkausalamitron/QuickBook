const QuickBooks = require('node-quickbooks');
const fs = require('fs');
const XLSX = require('xlsx');
const cron = require('node-cron');

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

// Define the cron schedule (e.g., every day at 8 AM)
const cronSchedule = '0 8 * * *';

// Schedule the report pulling and Excel file generation
cron.schedule(cronSchedule, async () => {
    try {
        // Pull a report from QuickBooks
        const report = await new Promise((resolve, reject) => {
            qbo.report(reportName, (err, report) => {
                if (err) {
                    reject(err);
                } else {
                    resolve(report);
                }
            });
        });

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
    } catch (err) {
        console.error('Error pulling report:', err);
    }
});
