const XLSX = require('xlsx');
const chalk = require('chalk');

function processExcelSheets(mainSheetPath, otherSheetPaths) {
  try {
    // Read main workbook
    const mainWorkbook = XLSX.readFile(mainSheetPath);
    const mainSheet = mainWorkbook.Sheets[mainWorkbook.SheetNames[0]];
    const mainData = XLSX.utils.sheet_to_json(mainSheet);

    // Collect all phone numbers from other sheets
    const numbersToRemove = new Set();
    
    otherSheetPaths.forEach(path => {
      const workbook = XLSX.readFile(path);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet);
      
      // Assuming phone numbers are in a column named 'Phone' or 'PhoneNumber'
      data.forEach(row => {
        const phoneNumber = row.Phone || row.PhoneNumber;
        if (phoneNumber) {
          // Normalize phone number (remove spaces, dashes, etc.)
          const normalized = phoneNumber.toString().replace(/\D/g, '');
          numbersToRemove.add(normalized);
        }
      });
    });

    // Filter out matching numbers from main sheet
    const filteredData = mainData.filter(row => {
      const phoneNumber = row.Phone || row.PhoneNumber;
      if (!phoneNumber) return true;
      
      const normalized = phoneNumber.toString().replace(/\D/g, '');
      return !numbersToRemove.has(normalized);
    });

    // Create new workbook with filtered data
    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(filteredData);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Filtered Data');

    // Save the result
    const outputPath = 'filtered_main_sheet.xlsx';
    XLSX.writeFile(newWorkbook, outputPath);

    console.log(chalk.green(`✓ Process completed successfully!`));
    console.log(chalk.blue(`→ Removed ${mainData.length - filteredData.length} phone numbers`));
    console.log(chalk.blue(`→ Output saved to: ${outputPath}`));

  } catch (error) {
    console.error(chalk.red('Error processing Excel sheets:'), error.message);
  }
}

// Example usage
const sheets = {
  main: 'main_sheet.xlsx',
  others: [
    'bm_sheet.xlsx',
    'dmm_sheet.xlsx',
    'ho_sheet.xlsx',
    'hoopla_sheet.xlsx'
  ]
};

processExcelSheets(sheets.main, sheets.others);