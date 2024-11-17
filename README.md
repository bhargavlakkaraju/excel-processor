# Excel Phone Number Processor

This tool processes Excel sheets to remove phone numbers from a main sheet that appear in other reference sheets.

## Usage

1. Place your Excel files in the project directory:
   - main_sheet.xlsx (the sheet to be filtered)
   - bm_sheet.xlsx
   - dmm_sheet.xlsx
   - ho_sheet.xlsx
   - hoopla_sheet.xlsx

2. Run the tool:
   ```bash
   npm start
   ```

3. The tool will create a new file called `filtered_main_sheet.xlsx` with the filtered results.

## Requirements

- The Excel sheets should have a column named either 'Phone' or 'PhoneNumber'
- Phone numbers can be in any format - the tool will normalize them by removing spaces, dashes, etc.

## Output

The tool will:
- Create a new Excel file with filtered data
- Show the number of entries removed
- Display the output file location