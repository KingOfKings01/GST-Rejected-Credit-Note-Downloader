const path = require('path');
const fs = require('fs'); // CRITICAL: Added File System module to build folders
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

/**
 * Handles the actions inside a specific invoice category drill-down view.
 * Opens the filter pane, selects 'Rejected', applies the filter, and downloads the Excel report if data exists.
 * * Saves file inside dynamic folder structure: downloadFolder/clientName/stateName/month/
 * * @param {import('playwright').Page} page - The active Playwright page instance
 * @param {string} downloadFolder - The root download directory from the GUI config panel
 * @param {object} client - The active client configuration containing metadata
 * @param {string} month - The currently evaluated targeted lookback month string name
 */
async function processRejectedInvoices(page, downloadFolder, client, month) {
    const { clientName, stateName } = client;

    try {
        // 1. Locate and click the "Filter" button
        const filterButton = page.locator('button#showbutton');
        await retry(async () => {
            await filterButton.waitFor({ state: 'visible', timeout: 10000 });
            await filterButton.click();
        });
        await page.waitForTimeout(300);

        // 2. Locate the precise "Status" dropdown form element
        const statusDropdown = page.locator('select[ng-model="status"]');
        await retry(async () => {
            await statusDropdown.waitFor({ state: 'visible', timeout: 10000 });
            await statusDropdown.click();
            await page.waitForTimeout(300);
            await statusDropdown.selectOption({ value: 'Rejected', label: 'Rejected' });
            await statusDropdown.dispatchEvent('change');
        });
        await page.waitForTimeout(300);

        // 3. Click the "Apply" button
        const applyButton = page.locator('button[ng-click*="suppfilterButton"]');
        await retry(async () => {
            await applyButton.waitFor({ state: 'visible', timeout: 10000 });
            await applyButton.click();
        });

        await page.waitForLoadState('networkidle');
        
        // Wait for dimmer loader to disappear stably
        await waitForDimmer(page, 300);

        // --- 4. CONDITIONAL DOWNLOAD CHECK ---
        const downloadButton = page.locator('button[data-ng-click*="exportDataOut"]').first();
        const noRecordAlert = page.locator('.alert-danger', { hasText: /No record found/i }).first();
        
        let isDataPresent = false;
        
        // Register dialog listener to handle any alerts (e.g. "No records found")
        const dialogHandler = async (dialog) => {
            await dialog.dismiss().catch(() => {});
        };
        page.on('dialog', dialogHandler);

        try {
            // Wait up to 10 seconds (20 * 500ms) for either "No record found" to show or download button to enable
            for (let attempt = 0; attempt < 20; attempt++) {
                if (await noRecordAlert.isVisible()) {
                    isDataPresent = false;
                    break;
                }
                if (await downloadButton.isVisible() && await downloadButton.isEnabled()) {
                    isDataPresent = true;
                    break;
                }
                await page.waitForTimeout(300);
            }

            if (isDataPresent) {
                // Set up a promise listener with a 15-second timeout to intercept the browser file stream download event
                const downloadPromise = page.waitForEvent('download', { timeout: 15000 }).catch((e) => {
                    return null;
                });

                await downloadButton.click();

                // Resolve the event promise stream to catch the file payload metadata
                const download = await downloadPromise;

                if (download) {
                    const suggestedFileName = download.suggestedFilename();

                    // =========================================================================
                    // 🛠️ DYNAMIC FOLDER STRUCTURE CREATION ENGINE
                    // =========================================================================
                    // Build absolute path to target sub-folder chain string: /Downloads/ClientName/State/Month
                    const targetNestedDirectory = path.join(downloadFolder, clientName, stateName, month);

                    // If any folders in this chain are missing, recursively build them automatically on the drive
                    if (!fs.existsSync(targetNestedDirectory)) {
                        fs.mkdirSync(targetNestedDirectory, { recursive: true });
                    }

                    // Append the actual file payload name to our newly configured directory
                    const finalSavePath = path.join(targetNestedDirectory, suggestedFileName);
                    // =========================================================================

                    // Commit file payload to your custom physical nested drive path layout
                    await download.saveAs(finalSavePath);
                    return true;
                }
            }
        } finally {
            // Clean up dialog handler to prevent memory leaks and unintended side effects on other pages
            page.off('dialog', dialogHandler);
        }

        await page.waitForTimeout(300);
        return false;

    } catch (error) {
        console.error("❌ An error occurred inside the sub-work routine:", error);
        throw error;
    }
}

module.exports = { processRejectedInvoices };