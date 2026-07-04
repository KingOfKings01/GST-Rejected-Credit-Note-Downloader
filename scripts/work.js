const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const { fillImsForm } = require('./formHandler');
const { processRejectedInvoices } = require('./subWorkHandler');
const { processLargeCountRejectedInvoices } = require('./largeCountSubWork');
const { waitForDimmer } = require('./utils');

async function retry(fn, retries = 2, delay = 1000) {
    for (let i = 0; i < retries; i++) {
        try {
            return await fn();
        } catch (e) {
            if (i === retries - 1) throw e;
            await new Promise(r => setTimeout(r, delay));
        }
    }
}

// Localized, sequential array of the Indian Fiscal Year months structure
const FISCAL_MONTHS = [
    "April", "May", "June", "July", "August", "September",
    "October", "November", "December", "January", "February", "March"
];

/**
 * Helper function to safely decrement a financial year string.
 * Example: Transforms '2025-26' into '2024-25'
 * @param {string} finYearStr 
 * @returns {string}
 */
function decrementFinancialYear(finYearStr) {
    const parts = finYearStr.split('-');
    if (parts.length !== 2) return finYearStr; // Return unchanged if format is unexpected

    const startYear = parseInt(parts[0], 10);
    const endYear = parseInt(parts[1], 10);

    // Decrement both segments by 1 year
    const newStart = startYear - 1;
    const newEnd = endYear - 1;

    // Maintain the YY format for the second block (extracting last 2 digits)
    const newEndStr = String(newEnd).slice(-2);

    return `${newStart}-${newEndStr}`;
}

function getTargetPeriodsQueue(baseSearch) {
    const targetPeriodsQueue = [];
    
    let fyRange = [];
    if (baseSearch.financialYearFrom && baseSearch.financialYearTo) {
        const fromStart = parseInt(baseSearch.financialYearFrom.split('-')[0], 10);
        const toStart = parseInt(baseSearch.financialYearTo.split('-')[0], 10);
        if (fromStart > toStart) {
            throw new Error("Financial Year From cannot be later than Financial Year To.");
        }
        for (let yr = fromStart; yr <= toStart; yr++) {
            const endYrStr = String(yr + 1).slice(-2);
            fyRange.push(`${yr}-${endYrStr}`);
        }
    } else {
        fyRange.push(baseSearch.financialYear || '2025-26');
    }

    if (baseSearch.returnPeriodFrom && baseSearch.returnPeriodTo) {
        const fromMonthIdx = FISCAL_MONTHS.indexOf(baseSearch.returnPeriodFrom);
        const toMonthIdx = FISCAL_MONTHS.indexOf(baseSearch.returnPeriodTo);

        if (fromMonthIdx === -1 || toMonthIdx === -1) {
            throw new Error(`Invalid return period range: ${baseSearch.returnPeriodFrom} to ${baseSearch.returnPeriodTo}`);
        }
        if (fyRange.length === 1 && fromMonthIdx > toMonthIdx) {
            throw new Error(`From Month (${baseSearch.returnPeriodFrom}) cannot be after To Month (${baseSearch.returnPeriodTo}).`);
        }

        for (let i = 0; i < fyRange.length; i++) {
            const fy = fyRange[i];
            const startIdx = (i === 0) ? fromMonthIdx : 0;
            const endIdx = (i === fyRange.length - 1) ? toMonthIdx : 11;

            // Chronological order for months within each FY
            for (let idx = startIdx; idx <= endIdx; idx++) {
                targetPeriodsQueue.push({
                    financialYear: fy,
                    returnPeriod: FISCAL_MONTHS[idx],
                    returnType: baseSearch.returnType
                });
            }
        }
    } else {
        // Locate where the user-selected month sits in our tracking array
        const selectedMonthIdx = FISCAL_MONTHS.indexOf(baseSearch.returnPeriod);
        if (selectedMonthIdx === -1) {
            throw new Error(`The return period '${baseSearch.returnPeriod}' is invalid.`);
        }

        // Fallback legacy 3-month lookback
        for (let lookbackOffset = 0; lookbackOffset < 3; lookbackOffset++) {
            const calculatedIdx = (selectedMonthIdx - lookbackOffset + 12) % 12;
            const targetedMonth = FISCAL_MONTHS[calculatedIdx];
            let targetedYear = fyRange[fyRange.length - 1];

            if (calculatedIdx > selectedMonthIdx) {
                targetedYear = decrementFinancialYear(targetedYear);
            }

            targetPeriodsQueue.push({
                financialYear: targetedYear,
                returnPeriod: targetedMonth,
                returnType: baseSearch.returnType
            });
        }
    }
    return targetPeriodsQueue;
}
/**
 * Helper function to safely decrement a financial year string.
 * Example: Transforms '2025-26' into '2024-25'
 * @param {string} finYearStr 
 * @returns {string}
 */
function decrementFinancialYear(finYearStr) {
    const parts = finYearStr.split('-');
    if (parts.length !== 2) return finYearStr; // Return unchanged if format is unexpected

    const startYear = parseInt(parts[0], 10);
    const endYear = parseInt(parts[1], 10);

    // Decrement both segments by 1 year
    const newStart = startYear - 1;
    const newEnd = endYear - 1;

    // Maintain the YY format for the second block (extracting last 2 digits)
    const newEndStr = String(newEnd).slice(-2);

    return `${newStart}-${newEndStr}`;
}

async function doGstWork(page, selections, downloadFolder, client, excelFilePath) {
    console.log("Starting post-login workflow operations...");
    let currentStage = 'Popup Handling';

    try {
        // --- 1. POPUP HANDLER SECTION ---
        let popupButton = page.locator('button, a, input, [role="button"]', { hasText: /remind me later/i }).first();

        try {
            // The popup might take a few seconds to load on slower connections
            await popupButton.waitFor({ state: 'visible', timeout: 8000 });
            await popupButton.click({ force: true });
         
            await page.waitForTimeout(2000);
        } catch (popupErr) {
            try {
                const fallbackButton = page.locator('div, span', { hasText: /remind me later/i }).first();
                if (await fallbackButton.count() > 0) {
                    await fallbackButton.click({ force: true });
         
                    await page.waitForTimeout(2000);
                }
            } catch (fallbackErr) {
                // If no popup appears or fails to click, proceed
            }
        }

        // --- 2. BASELINE INPUT CONFIGURATION ---
    const baseSearch = selections || {
        financialYearFrom: '2025-26',
        financialYearTo: '2025-26',
        returnPeriod: 'December',
        returnType: 'GSTR-1/IFF'
    };

    let targetPeriodsQueue = getTargetPeriodsQueue(baseSearch);

    if (client.excelFailedMonths) {
        const failedMonths = client.excelFailedMonths.split(',').map(m => m.trim().toLowerCase());
        if (failedMonths.length > 0 && failedMonths[0] !== '') {
            console.log(`System: Found previously failed months for ${client.username}: ${client.excelFailedMonths}. Restricting execution to these months.`);
            targetPeriodsQueue = targetPeriodsQueue.filter(p => failedMonths.includes(p.returnPeriod.toLowerCase()));
        }
    }

    const monthlyResults = [];

        // --- 4. MULTI-MONTH RUNNER ENGINE ---
        for (const currentPeriod of targetPeriodsQueue) {
            let monthDownloaded = false;
            let monthError = null;
            let isZipPending = false;
            let zipReadyAt = 0;
            let pendingSection = '';
            const successfulDownloads = [];

            try {
                currentStage = 'Navigating to IMS Dashboard';
                const imsSelector = 'a[href="//return.gst.gov.in/imsweb/auth/imsDashboard"]';
                await retry(async () => {
                    await page.waitForSelector(imsSelector, { state: 'attached', timeout: 30000 });
                    await page.evaluate((selector) => {
                        const element = document.querySelector(selector);
                        if (element) element.click();
                    }, imsSelector);
                    await page.waitForLoadState('networkidle');
                    await waitForDimmer(page, 300);
                });

                currentStage = 'Selecting Outward Supplies View';
                await retry(async () => {
                    const viewButton = page.locator('button[data-ng-click*="outwardsupplies"]').first();
                    await viewButton.waitFor({ state: 'visible', timeout: 30000 });
                    await viewButton.click();
                    await page.waitForLoadState('networkidle');
                    await waitForDimmer(page, 300);
                });

                currentStage = `Filing IMS Form details for ${currentPeriod.returnPeriod}`;
                await retry(async () => {
                    await fillImsForm(page, currentPeriod);
                });

                // Wait for dimmer loader to disappear stably
                await waitForDimmer(page, 300);

                currentStage = 'Loading Summary Table';
                const tableRowsSelector = 'table.table-responsive tbody tr';
                await retry(async () => {
                    await page.waitForSelector(tableRowsSelector, { state: 'visible', timeout: 30000 });
                });

                const totalRowsToProcess = 6;

                // Poll until all row counts are loaded (none are empty string "")
                let allCellsPopulated = false;
                for (let poll = 0; poll < 20; poll++) {
                    let emptyFound = false;
                    for (let r = 1; r <= totalRowsToProcess; r++) {
                        const currentRow = page.locator(tableRowsSelector, { has: page.locator(`td.sno-col:text-is("${r}")`) });
                        if (await currentRow.count() > 0) {
                            const val = await currentRow.locator('td:nth-child(3)').innerText();
                            if (val.trim() === '') {
                                emptyFound = true;
                                break;
                            }
                        } else {
                            emptyFound = true;
                            break;
                        }
                    }
                    if (!emptyFound) {
                        allCellsPopulated = true;
                        break;
                    }
                    await page.waitForTimeout(500);
                }

                // Extra short stabilization delay
                await page.waitForTimeout(1000);

                for (let srNo = 1; srNo <= totalRowsToProcess; srNo++) {
                    currentStage = `Processing Summary Row ${srNo}`;
                    const currentRow = page.locator(tableRowsSelector, { has: page.locator(`td.sno-col:text-is("${srNo}")`) });
                    await currentRow.waitFor({ state: 'attached', timeout: 30000 });

                    // Read the heading text of the section
                    const headingText = (await currentRow.locator('td:nth-child(2)').innerText()).trim().replace(/\n/g, ' ');

                    // Read the record count from the third column
                    const countText = (await currentRow.locator('td:nth-child(3)').innerText()).trim();
                    if (countText === '0') {
                        continue;
                    }

                    const count = parseInt(countText, 10);

                    // Check if there is a clickable link inside the heading column
                    const recordLink = currentRow.locator('td:nth-child(2) a');
                    if (await recordLink.count() > 0) {
                        // Check if already downloaded
                        const { formatFolderMonth } = require('./utils');
                        const folderLabel = formatFolderMonth(currentPeriod.returnPeriod, currentPeriod.financialYear);
                        const targetDir = path.join(downloadFolder, client.clientName, client.stateName, folderLabel);
                        
                        let sectionCode = '';
                        const lowerHeading = headingText.toLowerCase();
                        if (lowerHeading.includes('b2b')) sectionCode = 'B2B';
                        else if (lowerHeading.includes('cdnr')) sectionCode = 'CDNR';
                        else if (lowerHeading.includes('cdnra')) sectionCode = 'CDNRA';
                        
                        let alreadyDownloaded = false;
                        if (fs.existsSync(targetDir) && sectionCode) {
                            const files = fs.readdirSync(targetDir);
                            alreadyDownloaded = files.some(f => f.toUpperCase().includes(sectionCode) && (f.endsWith('.xlsx') || f.endsWith('.csv')));
                        }
                        
                        if (alreadyDownloaded) {
                            console.log(`System: Skipping Row ${srNo} (${headingText}) - already downloaded in ${folderLabel}`);
                            successfulDownloads.push({
                                section: headingText,
                                records: countText
                            });
                            monthDownloaded = true;
                            continue;
                        }

                        await retry(async () => {
                            await recordLink.click();
                            await page.waitForLoadState('networkidle');
                        });

                        if (count > 500) {
                            currentStage = `Downloading large count rejected invoices for Row ${srNo} (${headingText})`;
                            const result = await processLargeCountRejectedInvoices(page, downloadFolder, client, currentPeriod.returnPeriod, currentPeriod.financialYear);
                            if (result && result.success) {
                                monthDownloaded = true;
                                successfulDownloads.push({
                                    section: headingText,
                                    records: String(result.rejectedCount)
                                });
                            } else if (result && result.pending) {
                                isZipPending = true;
                                zipReadyAt = result.readyAt;
                                pendingSection = headingText;

                                if (excelFilePath) {
                                    const fy = (selections && selections.financialYearFrom) || '2025-26';
                                    const rt = (selections && selections.returnType) || 'GSTR-1/IFF';
                                    const singlePending = `Zip Pending | ${currentPeriod.returnPeriod} | ${result.readyAt} | ${headingText} | ${fy} | ${rt}`;
                                    
                                    updateExcelClientStatus(excelFilePath, client.username, 'Zip Pending', 'STATUS');
                                    
                                    const currentExcelZipStatus = getExcelZipStatus(excelFilePath, client.username);
                                    const existingZips = currentExcelZipStatus ? currentExcelZipStatus.split(';;').map(s => s.trim()).filter(s => s.startsWith('Zip Pending')) : [];
                                    const filteredZips = existingZips.filter(zipStr => {
                                        const parts = zipStr.split('|').map(s => s.trim());
                                        if (parts.length >= 4) {
                                            const existingMonth = parts[1];
                                            const existingSection = parts[3];
                                            return !(existingMonth === currentPeriod.returnPeriod && existingSection === headingText);
                                        }
                                        return true;
                                    });
                                    filteredZips.push(singlePending);
                                    updateExcelClientStatus(excelFilePath, client.username, filteredZips.join(' ;; '), 'ZIP STATUS');
                                }
                            }

                            // Try to click back button or go back
                            const uiBackButton = page.locator('button:has-text("Back"), button[data-ng-click*="back"], button[ng-click*="back"]').first();
                            if (await uiBackButton.count() > 0 && await uiBackButton.isVisible()) {
                                await uiBackButton.click();
                            } else {
                                await page.goBack();
                            }

                            await page.waitForLoadState('networkidle');
                            await page.waitForSelector(tableRowsSelector, { state: 'visible', timeout: 30000 });

                            // Wait for dimmer loader to disappear stably
                            await waitForDimmer(page, 300);
                        } else {
                            // Check if we actually entered the drill-down detail page by waiting for the Filter button
                            const filterButton = page.locator('button#showbutton');
                            let enteredDetail = false;
                            try {
                                await filterButton.waitFor({ state: 'visible', timeout: 10000 });
                                enteredDetail = true;
                            } catch (e) {
                            }

                            if (enteredDetail) {
                                currentStage = `Downloading rejected invoices for Row ${srNo} (${headingText})`;
                                // subWorkHandler checks for records and downloads
                                const downloaded = await processRejectedInvoices(page, downloadFolder, client, currentPeriod.returnPeriod, currentPeriod.financialYear);
                                if (downloaded) {
                                    monthDownloaded = true;
                                    successfulDownloads.push({
                                        section: headingText,
                                        records: countText
                                    });
                                }

                                // Try to click the page's Back button first to preserve SPA navigation routing
                                const uiBackButton = page.locator('button:has-text("Back"), button[data-ng-click*="back"], button[ng-click*="back"]').first();
                                if (await uiBackButton.count() > 0 && await uiBackButton.isVisible()) {
                                    await uiBackButton.click();
                                } else {
                                    await page.goBack();
                                }

                                await page.waitForLoadState('networkidle');
                                await page.waitForSelector(tableRowsSelector, { state: 'visible', timeout: 30000 });

                                // Wait for dimmer loader to disappear stably
                                await waitForDimmer(page, 300);
                            } else {
                                // Ensure summary table is still visible and stable before next iteration
                                await page.waitForSelector(tableRowsSelector, { state: 'visible', timeout: 15000 }).catch(() => { });
                            }
                        }
                    }
                }

            } catch (err) {
                let errMsg = err.message || String(err);
                if (errMsg.toLowerCase().includes('timeout')) {
                    errMsg = `Session Timeout / Page Load Timeout during stage: "${currentStage}"`;
                } else {
                    errMsg = `Error during stage "${currentStage}": ${errMsg}`;
                }
                console.error(`Error processing period ${currentPeriod.returnPeriod}:`, errMsg);
                monthError = errMsg;
            } finally {
                try {
                    const dashboardHomeBtn = page.locator('a[href*="dashboard"], button[id*="home"]').first();
                    if (await dashboardHomeBtn.count() > 0) {
                        await dashboardHomeBtn.click();
                        await page.waitForLoadState('networkidle');
                    } else {
                        await page.goto('https://services.gst.gov.in/services/auth/fowelcome');
                        const browserZoom = (selections && selections.browserZoom) || '0.85';
                        await page.evaluate((zoom) => {
                            const scaleVal = parseFloat(zoom) || 0.85;
                            document.body.style.transform = `scale(${scaleVal})`;
                            document.body.style.transformOrigin = 'top left'; // Keeps content aligned to the top-left
                            document.body.style.width = `${(1 / scaleVal) * 100}%`; // Compenses for the scale down so layout doesn't break
                        }, browserZoom);

                        await page.waitForLoadState('networkidle');
                    }
                } catch (resetErr) {
                    try {
                        await page.goto('https://services.gst.gov.in/services/auth/fowelcome');
                        const browserZoom = (selections && selections.browserZoom) || '0.85';
                        await page.evaluate((zoom) => {
                            const scaleVal = parseFloat(zoom) || 0.85;
                            document.body.style.transform = `scale(${scaleVal})`;
                            document.body.style.transformOrigin = 'top left'; // Keeps content aligned to the top-left
                            document.body.style.width = `${(1 / scaleVal) * 100}%`; // Compenses for the scale down so layout doesn't break
                        }, browserZoom);

                        await page.waitForLoadState('networkidle');
                    } catch (e) {
                        // Suppressed recovery logging
                    }
                }

                monthlyResults.push({
                    month: currentPeriod.returnPeriod,
                    financialYear: currentPeriod.financialYear,
                    isMonthChecked: true,
                    isDownloaded: monthDownloaded,
                    isZipPending: isZipPending,
                    zipReadyAt: zipReadyAt,
                    pendingSection: pendingSection,
                    downloads: successfulDownloads,
                    error: monthError
                });
            }
        }

        return monthlyResults;

    } catch (error) {
        console.error("An error occurred during post-login work steps:", error.message || error);
        return [{
            month: selections ? selections.returnPeriod : 'December',
            isMonthChecked: false,
            isDownloaded: false,
            downloads: [],
            error: error.message || String(error)
        }];
    }
}

function getExcelZipStatus(filePath, username) {
    if (!filePath || !fs.existsSync(filePath)) return '';
    try {
        const workbook = xlsx.readFile(filePath);
        const sheetName = 'customer data and filing return';
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return '';

        const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        if (rows.length === 0) return '';

        let headerRowIndex = 0;
        let userIdColIndex = -1;
        let zipStatusColIndex = -1;

        for (let r = 0; r < rows.length; r++) {
            const row = rows[r];
            if (row && row.indexOf('USER ID') !== -1) {
                headerRowIndex = r;
                userIdColIndex = row.indexOf('USER ID');
                break;
            }
        }

        if (userIdColIndex === -1) return '';

        const headerRow = rows[headerRowIndex];
        zipStatusColIndex = headerRow.indexOf('ZIP STATUS');
        if (zipStatusColIndex === -1) return '';

        for (let r = headerRowIndex + 1; r < rows.length; r++) {
            const row = rows[r];
            if (row && String(row[userIdColIndex]).trim() === String(username).trim()) {
                return String(row[zipStatusColIndex] || '').trim();
            }
        }
    } catch (e) {
        console.error('Error reading zip status from Excel:', e);
    }
    return '';
}

function updateExcelClientStatus(filePath, username, newStatus, columnName = 'STATUS') {
    if (!filePath) return;
    let attempts = 3;
    while (attempts > 0) {
        try {
            const workbook = xlsx.readFile(filePath);
            const sheetName = 'customer data and filing return';
            const sheet = workbook.Sheets[sheetName];
            if (!sheet) return;

            const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
            if (rows.length === 0) return;

            let headerRowIndex = 0;
            let userIdColIndex = -1;
            let statusColIndex = -1;

            for (let r = 0; r < rows.length; r++) {
                const row = rows[r];
                if (row && row.indexOf('USER ID') !== -1) {
                    headerRowIndex = r;
                    userIdColIndex = row.indexOf('USER ID');
                    break;
                }
            }

            if (userIdColIndex === -1) {
                const firstRow = rows[0] || [];
                userIdColIndex = firstRow.indexOf('USER ID');
                if (userIdColIndex === -1) return;
            }

            const headerRow = rows[headerRowIndex];
            statusColIndex = headerRow.indexOf(columnName);
            if (statusColIndex === -1) {
                statusColIndex = headerRow.length;
                const headerCellRef = xlsx.utils.encode_cell({ r: headerRowIndex, c: statusColIndex });
                sheet[headerCellRef] = { t: 's', v: columnName };
            }

            let targetRowIndex = -1;
            for (let r = headerRowIndex + 1; r < rows.length; r++) {
                const row = rows[r];
                if (row && String(row[userIdColIndex]).trim() === String(username).trim()) {
                    targetRowIndex = r;
                    break;
                }
            }

             if (targetRowIndex !== -1) {
                const cellRef = xlsx.utils.encode_cell({ r: targetRowIndex, c: statusColIndex });
                sheet[cellRef] = { t: 's', v: newStatus };

                const range = xlsx.utils.decode_range(sheet['!ref']);
                if (targetRowIndex > range.e.r) range.e.r = targetRowIndex;
                if (statusColIndex > range.e.c) range.e.c = statusColIndex;
                sheet['!ref'] = xlsx.utils.encode_range(range);

                xlsx.writeFile(workbook, filePath);
                return; // Success!
            }
            break;
        } catch (err) {
            if (err.code === 'EBUSY') {
                const start = Date.now();
                while (Date.now() - start < 2000) {} // Wait 2s synchronously
                attempts--;
            } else {
                console.error(`[Excel Sync Error]: Failed to update status in Excel for ${username}:`, err.message || err);
                break;
            }
        }
    }
}

module.exports = {
    doGstWork,
    getTargetPeriodsQueue
};