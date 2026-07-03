const xlsx = require('xlsx');
const fs = require('fs');
const { execFileSync } = require('child_process');

// Month name → index map (0-based, fiscal order)
const MONTH_NAME_TO_INDEX = {
    January: 0, February: 1, March: 2, April: 3, May: 4, June: 5,
    July: 6, August: 7, September: 8, October: 9, November: 10, December: 11
};

const MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/**
 * Converts a full month name + financial year string into a compact folder label.
 * Example: ("April", "2025-26") → "Apr 25"
 *          ("January", "2025-26") → "Jan 26"
 * @param {string} monthName  - Full month name, e.g. "April"
 * @param {string} financialYear - FY string, e.g. "2025-26"
 * @returns {string} Abbreviated label like "Apr 25"
 */
function formatFolderMonth(monthName, financialYear) {
    const abbr = MONTH_ABBR[MONTH_NAME_TO_INDEX[monthName]] || monthName.substring(0, 3);
    // Months Jan–Mar belong to the second calendar year of the FY (the "26" in "2025-26")
    const isSecondHalf = ['January', 'February', 'March'].includes(monthName);
    const parts = financialYear.split('-');
    const yearPart = isSecondHalf ? parts[1] : String(parseInt(parts[0], 10)).slice(-2);
    return `${abbr} ${yearPart}`;
}

/**
 * Detects the .dimmer-holder loader element on the page and waits until it disappears stably.
 * 
 * @param {import('playwright').Page} pageInstance - The Playwright Page instance.
 * @param {number} quietPeriodMs - The continuous time in milliseconds that the dimmer must remain hidden (default: 300).
 * @param {number} checkIntervalMs - The checking interval in milliseconds (default: 200).
 */
async function waitForDimmer(pageInstance, quietPeriodMs = 300, checkIntervalMs = 200) {
  
  // Helper function to check if the dimmer is currently visible
  const isDimmerVisible = async () => {
    return await pageInstance.evaluate(() => {
      const dimmer = document.querySelector('.dimmer-holder');
      if (!dimmer) return false;
      const style = window.getComputedStyle(dimmer);
      return (
        style.display !== 'none' &&
        style.visibility !== 'hidden' &&
        parseFloat(style.opacity) > 0
      );
    });
  };

  // 1. Initial wait for dimmer to disappear
  await pageInstance.waitForFunction(() => {
    const dimmer = document.querySelector('.dimmer-holder');
    if (!dimmer) return true;
    const style = window.getComputedStyle(dimmer);
    return (
      style.display === 'none' || 
      style.visibility === 'hidden' || 
      parseFloat(style.opacity) === 0
    );
  }, { timeout: 30000 }).catch(() => {});

  // 2. Verify that it stays gone (quiet period)
  let elapsedMs = 0;
  while (elapsedMs < quietPeriodMs) {
    await pageInstance.waitForTimeout(checkIntervalMs);
    elapsedMs += checkIntervalMs;

    if (await isDimmerVisible()) {    
      // Wait for it to disappear again
      await pageInstance.waitForFunction(() => {
        const dimmer = document.querySelector('.dimmer-holder');
        if (!dimmer) return true;
        const style = window.getComputedStyle(dimmer);
        return (
          style.display === 'none' || 
          style.visibility === 'hidden' || 
          parseFloat(style.opacity) === 0
        );
      }, { timeout: 30000 });

      // Reset elapsed time to restart the confirmation window
      elapsedMs = 0;
    }
  }
}

// --- DATE FORMATTING UTILITIES ---

const DATE_MONTHS_MAP = {
    jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2,
    apr: 3, april: 3, may: 4, jun: 5, june: 5, jul: 6, july: 6,
    aug: 7, august: 7, sep: 8, september: 8, sept: 8,
    oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11
};

/**
 * Parse any date-like value and return { date: JSDate, format } or null.
 * Handles: dd/mm/yyyy, dd-mm-yyyy, mm/yyyy, MMM-YY, yyyy-mm-dd, Excel serial numbers
 */
function parseToJSDate(val) {
    if (val === null || val === undefined || val === '') return null;
    if (typeof val === 'number') {
        // Only treat as date if within reasonable Excel date range (year 2000–2099)
        if (val >= 36526 && val <= 73050) {
            const excelEpoch = new Date(Date.UTC(1899, 11, 30));
            const jsDate = new Date(excelEpoch.getTime() + val * 24 * 60 * 60 * 1000);
            return { date: jsDate, format: 'dd-mm-yyyy' };
        }
        return null;
    }

    const str = String(val).trim();
    if (!str) return null;
    let match;

    // dd/mm/yyyy or dd-mm-yyyy
    match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (match) {
        const d = parseInt(match[1], 10), m = parseInt(match[2], 10), y = parseInt(match[3], 10);
        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
            return { date: new Date(Date.UTC(y, m - 1, d)), format: 'dd-mm-yyyy' };
        }
    }

    // yyyy-mm-dd or yyyy/mm/dd
    match = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (match) {
        const y = parseInt(match[1], 10), m = parseInt(match[2], 10), d = parseInt(match[3], 10);
        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
            return { date: new Date(Date.UTC(y, m - 1, d)), format: 'dd-mm-yyyy' };
        }
    }

    // MM/YYYY or MM-YYYY (period/month)
    match = str.match(/^(\d{1,2})[\/\-](\d{4})$/);
    if (match) {
        const m = parseInt(match[1], 10), y = parseInt(match[2], 10);
        if (m >= 1 && m <= 12) {
            return { date: new Date(Date.UTC(y, m - 1, 1)), format: 'mmm-yy' };
        }
    }

    // YYYY/MM or YYYY-MM
    match = str.match(/^(\d{4})[\/\-](\d{1,2})$/);
    if (match) {
        const y = parseInt(match[1], 10), m = parseInt(match[2], 10);
        if (m >= 1 && m <= 12) {
            return { date: new Date(Date.UTC(y, m - 1, 1)), format: 'mmm-yy' };
        }
    }

    // Alpha month + year: Apr-25, April-2025, Apr 25
    match = str.match(/^([A-Za-z]+)[\s\-]*(\d{2,4})$/);
    if (match) {
        const ms = match[1].toLowerCase(), ys = match[2];
        if (DATE_MONTHS_MAP[ms] !== undefined) {
            let y = parseInt(ys, 10);
            if (ys.length === 2) y += 2000;
            return { date: new Date(Date.UTC(y, DATE_MONTHS_MAP[ms], 1)), format: 'mmm-yy' };
        }
    }

    // Year + alpha month: 2025-Apr, 2025 Apr
    match = str.match(/^(\d{2,4})[\s\-]*([A-Za-z]+)$/);
    if (match) {
        const ys = match[1], ms = match[2].toLowerCase();
        if (DATE_MONTHS_MAP[ms] !== undefined) {
            let y = parseInt(ys, 10);
            if (ys.length === 2) y += 2000;
            return { date: new Date(Date.UTC(y, DATE_MONTHS_MAP[ms], 1)), format: 'mmm-yy' };
        }
    }

    return null;
}

/**
 * Determine the Excel date format for a column based on its header text.
 * Returns 'm/d/yyyy' for date columns, 'mmm-yy' for period/month columns, or null.
 */
function getDateFormat(headerText) {
    if (!headerText) return null;
    const h = String(headerText).trim().toLowerCase();
    if (h.includes('period') || h.includes('month')) return 'mmm-yy';
    if (h.includes('invoice date') || h.includes('note date') || h.includes('original date') || h.includes('revised date') || h.includes('date')) return 'dd-mm-yyyy';
    return null;
}

/**
 * Post-process a worksheet to convert all date column values to proper Excel dates.
 * Call this after xlsx.utils.aoa_to_sheet() to ensure proper date formatting.
 * Modifies cells in the sheet directly.
 */
function applyDateFormattingToSheet(sheet) {
    if (!sheet || !sheet['!ref']) return;

    const range = xlsx.utils.decode_range(sheet['!ref']);
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: true });

    // Find header row: first row with 4+ non-empty cells containing recognizable headers
    let headerIdx = -1;
    for (let r = 0; r < Math.min(10, rows.length); r++) {
        const row = rows[r];
        if (row &&
            row.filter(c => c !== null && c !== undefined && String(c).trim() !== '').length > 3 &&
            row.some(cell => cell && (
                getDateFormat(cell) !== null ||
                String(cell).toLowerCase().includes('status') ||
                String(cell).toLowerCase().includes('gstin') ||
                String(cell).toLowerCase().includes('remarks')
            ))
        ) {
            headerIdx = r;
            break;
        }
    }
    if (headerIdx === -1) return;

    // Apply autofilter starting from the header row
    sheet['!autofilter'] = { ref: xlsx.utils.encode_range({ s: { r: headerIdx, c: range.s.c }, e: range.e }) };

    const headers = rows[headerIdx];
    const dateColumns = [];
    headers.forEach((h, idx) => {
        const fmt = getDateFormat(h);
        if (fmt) dateColumns.push({ colIdx: idx, defaultFormat: fmt });
    });
    if (dateColumns.length === 0) return;

    // Convert each date cell to a proper Excel date
    for (const { colIdx, defaultFormat } of dateColumns) {
        for (let rowIdx = headerIdx + 1; rowIdx <= range.e.r; rowIdx++) {
            const cellAddr = xlsx.utils.encode_cell({ r: rowIdx, c: colIdx });
            const cell = sheet[cellAddr];
            if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue;

            const dateInfo = parseToJSDate(cell.v);
            if (dateInfo) {
                // For numbers already in the cell, use the column's default format
                // For parsed string values, use the format determined by the parser
                const fmt = (typeof cell.v === 'number') ? defaultFormat : dateInfo.format;
                sheet[cellAddr] = { v: dateInfo.date, t: 'd', z: fmt };
            }
        }
    }
}

/**
 * Fix date formatting in an Excel file in-place.
 * Reads the file, applies date formatting to all sheets, and writes back.
 */
function fixDatesInExcelFile(filePath) {
    try {
        if (!filePath || !filePath.endsWith('.xlsx')) return;
        const workbook = xlsx.readFile(filePath);
        for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];

            // Fix sheet range first
            let maxRow = 0;
            const cellKeys = Object.keys(sheet).filter(k => /^[A-Z]+\d+$/.test(k));
            for (const key of cellKeys) {
                const rowNum = parseInt(key.replace(/^[A-Z]+/, ''), 10);
                if (rowNum > maxRow) maxRow = rowNum;
            }
            if (maxRow > 0) {
                let colEnd = 'O';
                if (sheet['!ref']) {
                    const rangeMatch = sheet['!ref'].match(/:([A-Z]+)\d+/);
                    if (rangeMatch) colEnd = rangeMatch[1];
                }
                sheet['!ref'] = `A1:${colEnd}${maxRow}`;
            }

            applyDateFormattingToSheet(sheet);
        }
        xlsx.writeFile(workbook, filePath, { cellDates: true });
    } catch (err) {
        console.error(`⚠️ Warning: Failed to fix dates in ${filePath}: ${err.message}`);
    }
}

/**
 * Securely extracts a ZIP file to a target directory using powershell with -LiteralPath parameter.
 */
function unzipFile(zipPath, destDir) {
    if (!fs.existsSync(destDir)) {
        fs.mkdirSync(destDir, { recursive: true });
    }
    // Using powershell.exe with -LiteralPath to avoid wildcard expansion of brackets [ ] in paths.
    execFileSync('powershell.exe', [
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-Command',
        `Expand-Archive -LiteralPath "${zipPath.replace(/"/g, '`"')}" -DestinationPath "${destDir.replace(/"/g, '`"')}" -Force`
    ]);
}

module.exports = {
  waitForDimmer,
  formatFolderMonth,
  applyDateFormattingToSheet,
  fixDatesInExcelFile,
  unzipFile
};
