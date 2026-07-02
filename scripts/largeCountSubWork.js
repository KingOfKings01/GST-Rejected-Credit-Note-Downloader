const path = require('path');
const fs = require('fs');
const { formatFolderMonth, unzipFile, fixDatesInExcelFile } = require('./utils');

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
 * Skeleton handler for processing and downloading when record count is greater than 500.
 * @param {import('playwright').Page} page - The active Playwright page instance
 * @param {string} downloadFolder - The root download directory
 * @param {object} client - The active client configuration containing metadata
 * @param {string} month - The targeted lookback month string name
 * @returns {Promise<boolean>}
 */
async function processLargeCountRejectedInvoices(page, downloadFolder, client, month, financialYear) {
    const { clientName, stateName } = client;

    try {
        // 1. Locate and click the first "Download" link for large datasets (> 500 records)
        const firstDownloadLink = page.locator('a[data-ng-click*="dnldAdvSearchSumOut"], p a:has-text("Download")').first();
        await retry(async () => {
            await firstDownloadLink.waitFor({ state: 'visible', timeout: 15000 });
            await firstDownloadLink.click();
        });

        // 2. Locate the generated zip file link and the alert banner
        const secondDownloadLink = page.locator('span[data-ng-if="advSearchDnld"] a, a:has-text("Click here to")').first();
        const generationAlert = page.locator('.alert-success', { hasText: /request for file generation has been accepted|File generation is in progress/i }).first();

        // Wait up to 10 seconds for either the download link or the generation pending alert to appear
        let isAlertVisible = false;
        let isLinkVisible = false;
        
        for (let i = 0; i < 20; i++) {
            if (await generationAlert.isVisible()) {
                isAlertVisible = true;
                break;
            }
            if (await secondDownloadLink.isVisible()) {
                isLinkVisible = true;
                break;
            }
            await page.waitForTimeout(500);
        }

        if (isAlertVisible) {
            console.log(`⚠️ Zip file generation request accepted / in progress. Revisit in 20 minutes.`);
            return { success: false, pending: true, readyAt: Date.now() + 20 * 60 * 1000 };
        }

        // Wait up to 30 seconds for the download link to generate and become visible
        await retry(async () => {
            await secondDownloadLink.waitFor({ state: 'visible', timeout: 30000 });
        });

        // 3. Intercept download event and click the second download link
        const download = await retry(async () => {
            const downloadPromise = page.waitForEvent('download', { timeout: 30000 });
            await secondDownloadLink.click();
            return await downloadPromise;
        });

        if (download) {
            const suggestedFileName = download.suggestedFilename();
            const folderLabel = formatFolderMonth(month, financialYear);
            const targetNestedDirectory = path.join(downloadFolder, clientName, stateName, folderLabel);

            if (!fs.existsSync(targetNestedDirectory)) {
                fs.mkdirSync(targetNestedDirectory, { recursive: true });
            }

            const finalSavePath = path.join(targetNestedDirectory, suggestedFileName);
            await download.saveAs(finalSavePath);

            let totalRejectedCount = 0;
            // Automatically extract/unzip the file using PowerShell on Windows
            try {
                unzipFile(finalSavePath, targetNestedDirectory);

                // Read and process the extracted excel files to filter for "Rejected" records
                const xlsx = require('xlsx');
                const files = fs.readdirSync(targetNestedDirectory);
                
                for (const file of files) {
                    if (file.endsWith('.xlsx')) {
                        const filePath = path.join(targetNestedDirectory, file);
                        const workbook = xlsx.readFile(filePath);
                        const sheet = workbook.Sheets[workbook.SheetNames[0]];
                        
                        // Dynamically calculate the actual row count to fix the incorrect !ref range
                        let maxRow = 0;
                        const cellKeys = Object.keys(sheet).filter(k => /^[A-Z]+\d+$/.test(k));
                        for (const key of cellKeys) {
                            const rowNum = parseInt(key.replace(/^[A-Z]+/, ''), 10);
                            if (rowNum > maxRow) {
                                maxRow = rowNum;
                            }
                        }
                        if (maxRow > 0) {
                            let colEnd = 'O';
                            if (sheet['!ref']) {
                                const rangeMatch = sheet['!ref'].match(/:([A-Z]+)\d+/);
                                if (rangeMatch) colEnd = rangeMatch[1];
                            }
                            sheet['!ref'] = `A1:${colEnd}${maxRow}`;
                        }

                        const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

                        if (rows.length >= 5) {
                            const title = rows[0] ? rows[0][0] : '';
                            const section = rows[3] ? rows[3][0] : '';
                            const headers = rows[4] || [];

                            const statusIdx = headers.findIndex(h => h && String(h).trim().toLowerCase() === 'status');
                            const invTypeIdx = headers.findIndex(h => h && String(h).trim().toLowerCase() === 'invoice type');

                            const filteredRows = [];
                            for (let r = 5; r < rows.length; r++) {
                                const row = rows[r];
                                if (!row || row.length === 0) continue;
                                
                                const isStatusRejected = statusIdx !== -1 && String(row[statusIdx]).trim().toLowerCase() === 'rejected';
                                const isInvTypeRejected = invTypeIdx !== -1 && String(row[invTypeIdx]).trim().toLowerCase() === 'rejected';
                                
                                if (isStatusRejected || isInvTypeRejected) {
                                    filteredRows.push(row);
                                }
                            }

                            totalRejectedCount += filteredRows.length;

                            // Overwrite output XLSX file with filtered records, preserving headings
                            if (filteredRows.length > 0) {
                                const outputRows = [
                                    [title],
                                    [],
                                    [],
                                    [section],
                                    headers,
                                    [],
                                    ...filteredRows
                                ];
                                const newSheet = xlsx.utils.aoa_to_sheet(outputRows);
                                const newWb = xlsx.utils.book_new();
                                xlsx.utils.book_append_sheet(newWb, newSheet, workbook.SheetNames[0]);
                                xlsx.writeFile(newWb, filePath);
                                fixDatesInExcelFile(filePath);
                            } else {
                                // Clean up the .xlsx file if no rejected records exist
                                try {
                                    fs.unlinkSync(filePath);
                                } catch (e) {}
                            }
                        } else {
                            // Clean up invalid/empty files
                            try {
                                fs.unlinkSync(filePath);
                            } catch (e) {}
                        }
                    }
                }

                // Clean up the .zip file - Commented out for testing/verification purposes
                // fs.unlinkSync(finalSavePath);
            } catch (zipErr) {
                console.error("❌ Failed to process and extract ZIP archive:", zipErr.message || zipErr);
            }

            return { success: true, rejectedCount: totalRejectedCount };
        }

        return { success: false, rejectedCount: 0 };
    } catch (error) {
        console.error("❌ An error occurred inside large count sub-work routine:", error);
        return { success: false, rejectedCount: 0 };
    }
}

module.exports = { processLargeCountRejectedInvoices };
