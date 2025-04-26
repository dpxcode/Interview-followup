const XLSX = require('xlsx');

// Create a new workbook
const workbook = XLSX.utils.book_new();

// Define the headers
const headers = ['email', 'name', 'company', 'lastSentDate', 'followUpCount', 'responseStatus', 'responseDate', 'notes'];

// Create a worksheet with the headers
const worksheet = XLSX.utils.aoa_to_sheet([headers]);

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, worksheet, 'Follow-up Tracking');

// Write the workbook to a file
XLSX.writeFile(workbook, 'followup_tracking.xlsx');

console.log('Follow-up tracking Excel file created successfully with headers:', headers.join(', ')); 