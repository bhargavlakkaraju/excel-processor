const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

function processExcelSheets(mainSheetPath, otherSheetPaths) {
  return new Promise((resolve, reject) => {
    try {
      // Create downloads directory if it doesn't exist
      const downloadsDir = 'downloads';
      if (!fs.existsSync(downloadsDir)) {
        fs.mkdirSync(downloadsDir);
      }

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
        
        data.forEach(row => {
          // Look for phone numbers in any column of reference sheets
          Object.values(row).forEach(value => {
            if (value) {
              // Convert to string and normalize
              const normalized = value.toString().replace(/\D/g, '');
              // Only add if it's a 10-digit number
              if (normalized.length === 10) {
                numbersToRemove.add(normalized);
              }
            }
          });
        });
      });

      // Filter out rows with:
      // 1. Empty mobile numbers
      // 2. Mobile numbers that match reference sheets
      const filteredData = mainData.filter(row => {
        const mobileNumber = row['Mobile Number (10 Digit)'];
        
        // Remove row if mobile number is empty, undefined, or null
        if (!mobileNumber && mobileNumber !== 0) {
          return false;
        }
        
        const normalized = mobileNumber.toString().replace(/\D/g, '');
        // Remove row if number matches reference sheets
        return !numbersToRemove.has(normalized);
      });

      // Create new workbook with filtered data
      const newWorkbook = XLSX.utils.book_new();
      const newSheet = XLSX.utils.json_to_sheet(filteredData);
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Filtered Data');

      // Save the result
      const outputPath = path.join(downloadsDir, `filtered_main_sheet_${Date.now()}.xlsx`);
      XLSX.writeFile(newWorkbook, outputPath);

      resolve({
        outputPath,
        totalRows: mainData.length,
        remainingRows: filteredData.length,
        removedRows: mainData.length - filteredData.length
      });
    } catch (error) {
      reject(error);
    }
  });
}

module.exports = { processExcelSheets };