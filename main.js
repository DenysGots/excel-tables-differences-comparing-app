/**
 * XLSX Table Summarizer and Combiner (ExcelJS).
 *
 * - Reads 2 Excel files and stores them as arrays
 * - Summarizes items by simplifiedName and Ціна (sums Сума з Залишки)
 * - Combines the summarized tables side by side in a single XLSX file
 * - Adds 1 empty row above and 1 empty row on the left side between tables
 *
 * Usage:
 *   npm install exceljs
 *   node main.js file1.xlsx file2.xlsx output.xlsx
 *
 * With environment overrides:
 *   PRICE_DECIMALS=2 BALANCE_TOLERANCE=1 node main.js file1.xlsx file2.xlsx output.xlsx
 */

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const Fuse = require('fuse.js')

// Diff configuration (overridable via env)
const PRICE_DECIMALS = Number(process.env.PRICE_DECIMALS || 2);
const BALANCE_TOLERANCE = Number(process.env.BALANCE_TOLERANCE || 10.00);
const NAME_MATCH_THRESHOLD = Number(process.env.NAME_MATCH_THRESHOLD || 0.25);

function roundPriceValue(value) {
  const numeric = Number.parseFloat(String(value ?? '').toString().replace(',', '.'));
  if (!Number.isFinite(numeric)) return 0;
  const factor = 10 ** PRICE_DECIMALS;
  return Math.round(numeric * factor) / factor;
}

let fuse;

/**
 * Converts Excel value to appropriate value for writing
 * @param {any} value - The Excel cell value
 * @returns {any} - The converted value
 */
function convertExcelValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

  // Handle ExcelJS Date objects
  if (value instanceof Date) {
    return value;
  }

  // Handle ExcelJS RichText
  if (value && typeof value === 'object' && value.richText) {
    return value.richText.map(part => part.text).join('');
  }

  // Handle formulas (return the calculated result)
  if (value && typeof value === 'object' && value.formula) {
    return value.result;
  }

  return value;
}

/**
 * Validates input file exists and is readable
 * @param {string} filePath - Path to the file
 * @throws {Error} If file doesn't exist or isn't readable
 */
function validateInputFile(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Input file does not exist: ${filePath}`);
  }

  try {
    fs.accessSync(filePath, fs.constants.R_OK);
  } catch (err) {
    throw new Error(`Cannot read input file: ${filePath}`);
  }
}

/**
 * Reads an XLSX file and returns the data as a 2D array
 * @param {string} filePath - Path to the XLSX file
 * @returns {Promise<Array<Array<any>>>} - 2D array of cell values
 */
async function readXlsxAsArray(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = workbook.worksheets[0]; // Get first worksheet
  const data = [];

  worksheet.eachRow((row, rowNumber) => {
    const rowData = [];
    row.eachCell((cell, colNumber) => {
      rowData[colNumber - 1] = convertExcelValue(cell.value);
    });
    data.push(rowData);
  });

  return data;
}

/**
 * Writes a 2D array to an XLSX file
 * @param {Array<Array<any>>} data - 2D array of cell values
 * @param {string} filePath - Path to save the XLSX file
 */
async function writeArrayToXlsx(data, filePath) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Combined');

  data.forEach((row, rowIndex) => {
    row.forEach((cellValue, colIndex) => {
      if (cellValue !== undefined && cellValue !== '') {
        worksheet.getCell(rowIndex + 1, colIndex + 1).value = cellValue;
      }
    });
  });

  await workbook.xlsx.writeFile(filePath);
}

function normalizeItem(item) {
  const rules = [
    [/і/gi, "i"],
    [/с/gi, "c"],
    [/н/gi, "h"],
    [/е/gi, "e"],
    [/э/gi, "e"],
    [/у/gi, "y"],
    [/р/gi, "p"],
    [/а/gi, "A"],
    [/х/gi, "X"],
    [/к/gi, "K"],
    [/в/gi, "B"]
  ];

  return rules.reduce((text, [pattern, replacement]) => {
    return text.replace(pattern, replacement);
  }, item);
}

/**
 * Creates a simplified name by removing special characters and normalizing
 * @param {string} name - Original name
 * @returns {string} - Simplified name
 */
function createSimplifiedName(name) {
  if (!name || typeof name !== 'string') return '';
  const normalizedName = name.split(' ').sort().join(' ').replace(/[xX\/\\'"?._=|*:;~!@#№$%&()+`,\s\t\r–-]/gi, '').toLowerCase();

  return normalizeItem(normalizedName);
}

function createSimplifiedNameWithSpaces(name) {
  if (!name || typeof name !== 'string') return '';
  return name.trim().replace(/[xX\/\\'"?._=|*:;~!@#№$%&()+`,\t\r–-]/gi, '');
}

function buildNameVocabulary(allNames) {
  const uniqueEntries = [];
  const seenSimplified = new Set();

  allNames.forEach(name => {
    if (!name || typeof name !== 'string') return;
    const simplified = createSimplifiedName(name);
    if (!simplified) return;
    if (seenSimplified.has(simplified)) return;
    uniqueEntries.push({
      original: name,
      simplified,
      simplifiedWithSpaces: createSimplifiedNameWithSpaces(name),
    });
    seenSimplified.add(simplified);
  });

  const fuseInstance = new Fuse(uniqueEntries, {
    includeScore: true,
    threshold: 0.3,
    keys: ['simplified', 'simplifiedWithSpaces', 'original'],
  });

  const canonicalMap = new Map();

  uniqueEntries.forEach(entry => {
    if (canonicalMap.has(entry.simplified)) return;
    canonicalMap.set(entry.simplified, entry.simplified);
    const matches = fuseInstance.search(entry.simplified);
    matches.forEach(result => {
      if (!result.item) return;
      if (result.item.simplified === entry.simplified) return;
      if (result.score !== undefined && result.score <= NAME_MATCH_THRESHOLD) {
        canonicalMap.set(result.item.simplified, entry.simplified);
      }
    });
  });

  const resolveName = (name) => {
    if (!name || typeof name !== 'string') return '';
    const simplified = createSimplifiedName(name);
    if (!simplified) return '';
    if (canonicalMap.has(simplified)) {
      return canonicalMap.get(simplified);
    }
    const [bestMatch] = fuseInstance.search(simplified, { limit: 1 });
    if (bestMatch && bestMatch.score !== undefined && bestMatch.score <= NAME_MATCH_THRESHOLD) {
      const canonical = canonicalMap.get(bestMatch.item.simplified) || bestMatch.item.simplified;
      canonicalMap.set(simplified, canonical);
      return canonical;
    }
    canonicalMap.set(simplified, simplified);
    return simplified;
  };

  return { resolveName, fuseInstance };
}

/**
 * Converts a 2D array to an array of objects with headers
 * @param {Array<Array<any>>} data - 2D array data
 * @returns {Array<Object>} - Array of objects with header keys
 */
function arrayToObjects(data) {
  if (data.length === 0) return [];

  const headers = data[0];
  const objects = [];

  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = data[i][index];
    });
    objects.push(obj);
  }

  return objects;
}

/**
 * Normalize monetary fields and recalc balance
 * Ensures fields are floats with PRICE_DECIMALS precision; empty -> 0
 */
function normalizeAndRecalc(rows) {
  const IN_FIELD = 'Сума з Надійшло';
  const OUT_FIELD = 'Сума з Вибуло';
  const BAL_FIELD = 'Сума з Залишки';

  return rows.map(row => {
    const incoming = Number.parseFloat(String(row[IN_FIELD] ?? '').toString().replace(',', '.'));
    const outgoing = Number.parseFloat(String(row[OUT_FIELD] ?? '').toString().replace(',', '.'));
    const inVal = Number.isFinite(incoming) ? incoming : 0;
    const outVal = Number.isFinite(outgoing) ? outgoing : 0;
    const balVal = Number((inVal - outVal).toFixed(PRICE_DECIMALS));
    return {
      ...row,
      [IN_FIELD]: Number(inVal.toFixed(PRICE_DECIMALS)),
      [OUT_FIELD]: Number(outVal.toFixed(PRICE_DECIMALS)),
      [BAL_FIELD]: balVal,
    };
  });
}

/**
 * Converts an array of objects back to 2D array
 * @param {Array<Object>} objects - Array of objects
 * @param {Array<string>} headers - Headers in order
 * @returns {Array<Array<any>>} - 2D array with headers as first row
 */
function objectsToArray(objects, headers) {
  const result = [headers];

  objects.forEach(obj => {
    const row = headers.map(header => {
      // Normalize numeric fields to 0
      if (header === 'Ціна' || header === 'Сума з Надійшло' || header === 'Сума з Вибуло' || header === 'Сума з Залишки') {
        const num = Number.parseFloat(String(obj[header] ?? '').toString().replace(',', '.'));
        const safe = Number.isFinite(num) ? num : 0;
        return Number(safe.toFixed(PRICE_DECIMALS));
      }
      return obj[header] || '';
    });
    result.push(row);
  });

  return result;
}

/**
 * Summarizes table data by simplifiedName and Ціна, summing Сума з Залишки
 * @param {Array<Object>} tableData - Array of objects representing table rows
 * @returns {Array<Object>} - Summarized table data
 */
function summarizeTable(tableData, resolveName) {
  const summaryMap = new Map();

  tableData.forEach(row => {
    const simplifiedName = resolveName(row['Найменування']);
    const price = parseFloat(row['Ціна']) || 0;
    const sumWithBalance = parseFloat(row['Сума з Залишки']) || 0;
    const source = (row['Джерело надходження'] || '').toString().trim();

    // Skip rows with invalid data
    if (!simplifiedName || isNaN(price)) return;

    // Round price to BALANCE_TOLERANCE for grouping
    const roundedPrice = roundPriceValue(price);
    const key = `${simplifiedName}_${roundedPrice.toFixed(PRICE_DECIMALS)}`;

    if (summaryMap.has(key)) {
      const existing = summaryMap.get(key);
      existing['Сума з Залишки'] = (parseFloat(existing['Сума з Залишки']) || 0) + sumWithBalance;
      // Concatenate sources preserving duplicates and order
      if (source && existing['Джерело надходження']) {
        existing['Джерело надходження'] = `${existing['Джерело надходження']}, ${source}`;
      } else if (source) {
        existing['Джерело надходження'] = source;
      }
    } else {
      // Create a copy of the row and update the sum
      const newRow = { ...row };
      newRow['Сума з Залишки'] = sumWithBalance;
      if (source) {
        newRow['Джерело надходження'] = source;
      }
      summaryMap.set(key, newRow);
    }
  });

  return Array.from(summaryMap.values());
}

/**
 * Filters out rows where "Сума з Залишки" is empty or zero
 * @param {Array<Object>} rows - Array of summarized rows
 * @returns {Array<Object>} - Filtered rows
 */
function filterRowsWithNonEmptyBalance(rows) {
  return rows.filter(row => {
    const value = row['Сума з Залишки'];
    if (value === null || value === undefined) return false;
    if (value === '') return false;
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return false;
    return numeric !== 0;
  });
}

/**
 * Builds a composite key for diffing: simplifiedName + rounded price + rounded balance
 * @param {Object} row
 * @returns {string}
 */
function buildKeyNameAndRoundedPrice(row, resolveName) {
  const simplifiedName = resolveName(row['Найменування']);
  const price = Number(row['Ціна']) || 0;
  const balance = Number(row['Сума з Залишки']) || 0;
  const roundedPrice = roundPriceValue(price);
  const roundedBalance = Math.round(balance / BALANCE_TOLERANCE) * BALANCE_TOLERANCE;
  return `${simplifiedName}__${roundedPrice.toFixed(PRICE_DECIMALS)}__${roundedBalance.toFixed(PRICE_DECIMALS)}`;
}

/**
 * Filters out rows that are present in both arrays using exact key match
 * @param {Array<Object>} a
 * @param {Array<Object>} b
 * @returns {{onlyA: Array<Object>, onlyB: Array<Object>, matchedCount: number}}
 */
function diffTables(a, b, resolveName) {
  // Build map for B: key -> array of rows
  const mapB = new Map();
  b.forEach(row => {
    const key = buildKeyNameAndRoundedPrice(row, resolveName);
    const arr = mapB.get(key) || [];
    arr.push(row);
    mapB.set(key, arr);
  });

  let matchedCount = 0;
  const onlyA = [];

  a.forEach(rowA => {
    const key = buildKeyNameAndRoundedPrice(rowA, resolveName);
    const candidates = mapB.get(key);
    if (!candidates || candidates.length === 0) {
      onlyA.push(rowA);
      return;
    }
    // Exact match found, remove it
    candidates.splice(0, 1);
    matchedCount += 1;
  });

  // Remaining rows in mapB are onlyB
  const onlyB = [];
  for (const arr of mapB.values()) {
    for (const row of arr) {
      onlyB.push(row);
    }
  }

  // TODO: narrow the result onlyA and onlyB using mapA and mapB and fuse library
  return { onlyA, onlyB, matchedCount };
}

/**
 * Writes combined data and a summary sheet to XLSX
 * @param {Array<Array<any>>} combinedData
 * @param {Array<Array<any>>} summaryRows - 2D array with summary rows
 * @param {string} filePath
 */
async function writeCombinedAndSummary(combinedData, summaryRows, filePath) {
  const workbook = new ExcelJS.Workbook();
  const wsCombined = workbook.addWorksheet('Combined');
  combinedData.forEach((row, rowIndex) => {
    row.forEach((cellValue, colIndex) => {
      if (cellValue !== undefined && cellValue !== '') {
        wsCombined.getCell(rowIndex + 1, colIndex + 1).value = cellValue;
        }
      });
    });

  const wsSummary = workbook.addWorksheet('Summary');
  summaryRows.forEach((row, rowIndex) => {
    row.forEach((cellValue, colIndex) => {
      wsSummary.getCell(rowIndex + 1, colIndex + 1).value = cellValue;
    });
  });

  await workbook.xlsx.writeFile(filePath);
}

/**
 * Combines two 2D arrays side by side with gaps
 * @param {Array<Array<any>>} table1 - First table data
 * @param {Array<Array<any>>} table2 - Second table data
 * @returns {Array<Array<any>>} - Combined table data
 */
function combineTablesSideBySide(table1, table2, file1Name, file2Name) {
  const maxRows = Math.max(table1.length, table2.length);
  const table1Cols = table1.length > 0 ? table1[0].length : 0;
  const table2Cols = table2.length > 0 ? table2[0].length : 0;

  const combined = [];

  // Title row directly above tables
  const titleRow = new Array(table1Cols + 1 + table2Cols).fill('');
  if (file1Name) titleRow[1] = file1Name; // shifted one cell to the right
  if (file2Name) titleRow[table1Cols + 2] = file2Name; // shifted one cell to the right
  combined.push(titleRow);

  for (let row = 0; row < maxRows; row++) {
    const combinedRow = [];

    if (row < table1.length) {
      combinedRow.push(...table1[row]);
    } else {
      combinedRow.push(...new Array(table1Cols).fill(''));
    }

    combinedRow.push('');

    if (row < table2.length) {
      combinedRow.push(...table2[row]);
    } else {
      combinedRow.push(...new Array(table2Cols).fill(''));
    }

    combined.push(combinedRow);
  }

  return combined;
}

async function main() {
  try {
    if (process.argv.length < 5) {
      console.error("Usage: node main.js <file1.xlsx> <file2.xlsx> <output.xlsx>");
      process.exit(1);
    }

    const file1Path = path.resolve(process.argv[2]);
    const file2Path = path.resolve(process.argv[3]);
    const outputPath = path.resolve(process.argv[4]);

    // Validate input files
    validateInputFile(file1Path);
    validateInputFile(file2Path);

    console.log(`Reading first Excel file: ${file1Path}`);
    console.log(`Reading second Excel file: ${file2Path}`);
    console.log(`Output will be saved to: ${outputPath}`);

    // Read both files as arrays
    const table1Raw = await readXlsxAsArray(file1Path);
    const table2Raw = await readXlsxAsArray(file2Path);

    console.log(`Raw Table 1 dimensions: ${table1Raw.length} rows x ${table1Raw.length > 0 ? table1Raw[0].length : 0} columns`);
    console.log(`Raw Table 2 dimensions: ${table2Raw.length} rows x ${table2Raw.length > 0 ? table2Raw[0].length : 0} columns`);

    // Convert to objects for easier processing
    let table1Objects = arrayToObjects(table1Raw);
    let table2Objects = arrayToObjects(table2Raw);

    // Normalize monetary fields and recalc balances upfront
    table1Objects = normalizeAndRecalc(table1Objects);
    table2Objects = normalizeAndRecalc(table2Objects);

    console.log(`Table 1 data rows: ${table1Objects.length}`);
    console.log(`Table 2 data rows: ${table2Objects.length}`);

    const allNames = [...table1Objects, ...table2Objects].map(row => row['Найменування']).filter(Boolean);
    const { resolveName, fuseInstance } = buildNameVocabulary(allNames);
    fuse = fuseInstance;

    // Summarize both tables
    const summarizedTable1 = summarizeTable(table1Objects, resolveName);
    const summarizedTable2 = summarizeTable(table2Objects, resolveName);

    // Filter out rows with empty/zero "Сума з Залишки"
    const filteredTable1 = filterRowsWithNonEmptyBalance(summarizedTable1);
    const filteredTable2 = filterRowsWithNonEmptyBalance(summarizedTable2);

    console.log(`Summarized Table 1 rows: ${summarizedTable1.length} (reduced from ${table1Objects.length})`);
    console.log(`Summarized Table 2 rows: ${summarizedTable2.length} (reduced from ${table2Objects.length})`);
    console.log(`Filtered Table 1 rows (non-empty balance): ${filteredTable1.length}`);
    console.log(`Filtered Table 2 rows (non-empty balance): ${filteredTable2.length}`);

    // Diff: leave only items that are different in each table
    const { onlyA: onlyTable1, onlyB: onlyTable2 } = diffTables(filteredTable1, filteredTable2, resolveName);
    console.log(`Only in Table 1 after diff: ${onlyTable1.length}`);
    console.log(`Only in Table 2 after diff: ${onlyTable2.length}`);

    // Show some examples of summarization
    if (summarizedTable1.length < table1Objects.length) {
      console.log(`Table 1: ${table1Objects.length - summarizedTable1.length} duplicate items were summarized`);
    }
    if (summarizedTable2.length < table2Objects.length) {
      console.log(`Table 2: ${table2Objects.length - summarizedTable2.length} duplicate items were summarized`);
    }

    // Convert back to 2D arrays with per-table headers and source as rightmost if present
    const sourceField = 'Джерело надходження';
    let headers1 = Array.isArray(table1Raw[0]) ? [...table1Raw[0]] : [];
    let headers2 = Array.isArray(table2Raw[0]) ? [...table2Raw[0]] : [];

    const table1HasSource = onlyTable1.some(r => r[sourceField] !== undefined && r[sourceField] !== null && r[sourceField] !== '');
    const table2HasSource = onlyTable2.some(r => r[sourceField] !== undefined && r[sourceField] !== null && r[sourceField] !== '');

    if (table1HasSource && !headers1.includes(sourceField)) headers1.push(sourceField);
    if (table2HasSource && !headers2.includes(sourceField)) headers2.push(sourceField);

    const table1 = objectsToArray(onlyTable1, headers1);
    const table2 = objectsToArray(onlyTable2, headers2);

    console.log(`Final Table 1 dimensions: ${table1.length} rows x ${table1.length > 0 ? table1[0].length : 0} columns`);
    console.log(`Final Table 2 dimensions: ${table2.length} rows x ${table2.length > 0 ? table2[0].length : 0} columns`);

    // Combine tables side by side
    const combinedData = combineTablesSideBySide(table1, table2, path.basename(file1Path), path.basename(file2Path));

    console.log(`Combined table dimensions: ${combinedData.length} rows x ${combinedData.length > 0 ? combinedData[0].length : 0} columns`);

    // Prepare summary sheet
    const summaryRows = [
      ['Setting', 'Value'],
      ['PRICE_DECIMALS', PRICE_DECIMALS],
      ['BALANCE_TOLERANCE', BALANCE_TOLERANCE],
      [],
      ['Metric', 'Table 1', 'Table 2'],
      ['Raw rows', table1Raw.length - 1, table2Raw.length - 1],
      ['Summarized rows', summarizedTable1.length, summarizedTable2.length],
      ['Filtered non-empty rows', filteredTable1.length, filteredTable2.length],
      ['Only in after diff', onlyTable1.length, onlyTable2.length],
    ];

    // Write combined data and summary to new XLSX file
    await writeCombinedAndSummary(combinedData, summaryRows, outputPath);

    console.log(`Successfully combined tables and summary saved to: ${outputPath}`);

  } catch (err) {
    console.error(`Error: ${err.message}`);
    if (err.stack) {
      console.error(err.stack);
    }
    process.exit(1);
  }
}

main().catch((err) => {
  console.error(`Unexpected error: ${err.message}`);
  if (err.stack) {
    console.error(err.stack);
  }
  process.exit(1);
});
