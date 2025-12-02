XLSX Table Summarizer and Combiner (ExcelJS).

- Reads 2 Excel files and stores them as arrays
- Summarizes items by simplifiedName and Ціна (sums Сума з Залишки)
- Combines the summarized tables side by side in a single XLSX file
- Adds 1 empty row above and 1 empty row on the left side between tables

Usage:
npm install exceljs
node main.js file1.xlsx file2.xlsx output.xlsx

With environment overrides:
PRICE_DECIMALS=2 BALANCE_TOLERANCE=1 node main.js file1.xlsx file2.xlsx output.xlsx