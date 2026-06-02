const { fillImsForm } = require('./formHandler');
const { processRejectedInvoices } = require('./subWorkHandler');
const { processLargeCountRejectedInvoices } = require('./largeCountSubWork');

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

async function doGstWork(page, selections, downloadFolder, client) {
    console.log("Starting post-login workflow operations...");
    let currentStage = 'Popup Handling';

    try {
        // --- 1. POPUP HANDLER SECTION ---
        const popupButton = page.locator('button, a, input, [role="button"], div, span', { hasText: /remind me later/i }).first();

        try {
            // The popup might take a few seconds to load on slower connections
            await popupButton.waitFor({ state: 'visible', timeout: 8000 });
            await popupButton.click();
            await page.waitForTimeout(2000);
        } catch (popupErr) {
            // If no popup appears within the timeout, we proceed
        }

        // --- 2. BASELINE INPUT CONFIGURATION ---
        const baseSearch = selections || {
            financialYear: '2025-26',
            returnPeriod: 'December',
            returnType: 'GSTR-1/IFF'
        };

        // Locate where the user-selected month sits in our tracking array
        const selectedMonthIdx = FISCAL_MONTHS.indexOf(baseSearch.returnPeriod);
        if (selectedMonthIdx === -1) {
            throw new Error(`The return period '${baseSearch.returnPeriod}' is invalid.`);
        }

        // --- 3. BUILD LOOKBACK QUEUE USING THE MODULO (%) PROPERTY ---
        const targetPeriodsQueue = [];

        for (let lookbackOffset = 0; lookbackOffset < 3; lookbackOffset++) {
            // Apply infinite wrap-around modulo logic: (index - offset + 12) % 12
            // Adding 12 eliminates negative pointer outcomes when crossing into a previous fiscal year framework
            const calculatedIdx = (selectedMonthIdx - lookbackOffset + 12) % 12;
            const targetedMonth = FISCAL_MONTHS[calculatedIdx];

            let targetedYear = baseSearch.financialYear;

            // CRITICAL CHECK: If our lookback calculation crosses backward past April (index 0), 
            // the index becomes greater than the selected starting month index, signaling a year shift.
            if (calculatedIdx > selectedMonthIdx) {
                targetedYear = decrementFinancialYear(baseSearch.financialYear);
            }

            targetPeriodsQueue.push({
                financialYear: targetedYear,
                returnPeriod: targetedMonth,
                returnType: baseSearch.returnType
            });
        }


        const monthlyResults = [];

        // --- 4. MULTI-MONTH RUNNER ENGINE ---
        for (const currentPeriod of targetPeriodsQueue) {
            let monthDownloaded = false;
            let monthError = null;
            const successfulDownloads = [];

            try {
                currentStage = 'Navigating to IMS Dashboard';
                const imsSelector = 'a[href="//return.gst.gov.in/imsweb/auth/imsDashboard"]';
                await retry(async () => {
                    await page.waitForSelector(imsSelector, { state: 'attached', timeout: 15000 });
                    await page.evaluate((selector) => {
                        const element = document.querySelector(selector);
                        if (element) element.click();
                    }, imsSelector);
                    await page.waitForLoadState('networkidle');
                });

                currentStage = 'Selecting Outward Supplies View';
                await retry(async () => {
                    const viewButton = page.locator('button[data-ng-click*="outwardsupplies"]').first();
                    await viewButton.waitFor({ state: 'visible', timeout: 10000 });
                    await viewButton.click();
                    await page.waitForLoadState('networkidle');
                });

                currentStage = `Filing IMS Form details for ${currentPeriod.returnPeriod}`;
                await retry(async () => {
                    await fillImsForm(page, currentPeriod);
                });

                // Wait for summary loading spinner to hide completely
                const spinner = page.locator('.loading, .loading-backdrop, #loading, .spinner, .ajax-loader').first();
                for (let i = 0; i < 10; i++) {
                    if (await spinner.isVisible()) {
                        await spinner.waitFor({ state: 'hidden', timeout: 8000 }).catch(() => { });
                    }
                    await page.waitForTimeout(300);
                }
                await page.waitForTimeout(1000); // Increased stable rendering timeout

                currentStage = 'Loading Summary Table';
                const tableRowsSelector = 'table.table-responsive tbody tr';
                await retry(async () => {
                    await page.waitForSelector(tableRowsSelector, { state: 'visible', timeout: 10000 });
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
                    await currentRow.waitFor({ state: 'attached', timeout: 5000 });

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
                        await retry(async () => {
                            await recordLink.click();
                            await page.waitForLoadState('networkidle');
                        });

                        if (count > 500) {
                            currentStage = `Downloading large count rejected invoices for Row ${srNo} (${headingText})`;
                            const result = await processLargeCountRejectedInvoices(page, downloadFolder, client, currentPeriod.returnPeriod);
                            if (result && result.success) {
                                monthDownloaded = true;
                                successfulDownloads.push({
                                    section: headingText,
                                    records: String(result.rejectedCount)
                                });
                            }

                            // Try to click back button or go back
                            const uiBackButton = page.locator('button:has-text("Back"), button[data-ng-click*="back"], button[ng-click*="back"]').first();
                            if (await uiBackButton.count() > 0 && await uiBackButton.isVisible()) {
                                await uiBackButton.click();
                            } else {
                                await page.goBack();
                            }

                            await page.waitForLoadState('networkidle');
                            await page.waitForSelector(tableRowsSelector, { state: 'visible', timeout: 15000 });

                            // Wait for any loading spinner to disappear
                            const backSpinner = page.locator('.loading, .loading-backdrop, #loading, .spinner, .ajax-loader').first();
                            for (let i = 0; i < 10; i++) {
                                if (await backSpinner.isVisible()) {
                                    await backSpinner.waitFor({ state: 'hidden', timeout: 8000 }).catch(() => { });
                                }
                                await page.waitForTimeout(300);
                            }
                            await page.waitForTimeout(1500); // Stabilization pause
                        } else {
                            // Check if we actually entered the drill-down detail page by waiting for the Filter button
                            const filterButton = page.locator('button#showbutton');
                            let enteredDetail = false;
                            try {
                                await filterButton.waitFor({ state: 'visible', timeout: 4000 });
                                enteredDetail = true;
                            } catch (e) {
                            }

                            if (enteredDetail) {
                                currentStage = `Downloading rejected invoices for Row ${srNo} (${headingText})`;
                                // subWorkHandler checks for records and downloads
                                const downloaded = await processRejectedInvoices(page, downloadFolder, client, currentPeriod.returnPeriod);
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
                                await page.waitForSelector(tableRowsSelector, { state: 'visible', timeout: 15000 });

                                // Wait for any loading spinner to disappear after returning to the summary table
                                const backSpinner = page.locator('.loading, .loading-backdrop, #loading, .spinner, .ajax-loader').first();
                                for (let i = 0; i < 10; i++) {
                                    if (await backSpinner.isVisible()) {
                                        await backSpinner.waitFor({ state: 'hidden', timeout: 8000 }).catch(() => { });
                                    }
                                    await page.waitForTimeout(300);
                                }
                                await page.waitForTimeout(1500); // Stabilization pause
                            } else {
                                // Ensure summary table is still visible and stable before next iteration
                                await page.waitForSelector(tableRowsSelector, { state: 'visible', timeout: 5000 }).catch(() => { });
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
                    isMonthChecked: true,
                    isDownloaded: monthDownloaded,
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

module.exports = { doGstWork };