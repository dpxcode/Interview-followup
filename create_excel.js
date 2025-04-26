const XLSX = require('xlsx');

// Create a new workbook
const workbook = XLSX.utils.book_new();

// Define the headers
const headers = ['email', 'name', 'company', 'lastSentDate'];

// Create a worksheet with the headers
const worksheet = XLSX.utils.aoa_to_sheet([headers]);

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, worksheet, 'HR Contacts');

// Write the workbook to a file
XLSX.writeFile(workbook, 'hr_contacts.xlsx');

console.log('Excel file created successfully with headers:', headers.join(', ')); 