const path = require('path');
const fs = require('fs');

/**
 * Skeleton handler for processing and downloading when record count is greater than 500.
 * @param {import('playwright').Page} page - The active Playwright page instance
 * @param {string} downloadFolder - The root download directory
 * @param {object} client - The active client configuration containing metadata
 * @param {string} month - The targeted lookback month string name
 * @returns {Promise<boolean>}
 */
async function processLargeCountRejectedInvoices(page, downloadFolder, client, month) {
    const { clientName, stateName } = client;

    try {
        // 1. Locate and click the first "Download" link for large datasets (> 500 records)
        const firstDownloadLink = page.locator('a[data-ng-click*="dnldAdvSearchSumOut"], p a:has-text("Download")').first();
        await firstDownloadLink.waitFor({ state: 'visible', timeout: 15000 });
        await firstDownloadLink.click();

        // 2. Locate the generated zip file link (which appears after click/generation finishes)
        const secondDownloadLink = page.locator('span[data-ng-if="advSearchDnld"] a, a:has-text("Click here to")').first();
        
        // Wait up to 30 seconds for the download link to generate and become visible
        await secondDownloadLink.waitFor({ state: 'visible', timeout: 30000 });

        // 3. Intercept download event and click the second download link
        const downloadPromise = page.waitForEvent('download', { timeout: 30000 });
        await secondDownloadLink.click();
        const download = await downloadPromise;

        if (download) {
            const suggestedFileName = download.suggestedFilename();
            const targetNestedDirectory = path.join(downloadFolder, clientName, stateName, month);

            if (!fs.existsSync(targetNestedDirectory)) {
                fs.mkdirSync(targetNestedDirectory, { recursive: true });
            }

            const finalSavePath = path.join(targetNestedDirectory, suggestedFileName);
            await download.saveAs(finalSavePath);

            let totalRejectedCount = 0;
            // Automatically extract/unzip the file using PowerShell on Windows
            try {
                const { execSync } = require('child_process');
                const cmd = `powershell -Command "Expand-Archive -Path \\"${finalSavePath}\\" -DestinationPath \\"${targetNestedDirectory}\\" -Force"`;
                execSync(cmd);

                // Read and process the extracted excel files to filter for "Rejected" records
                const xlsx = require('xlsx');
                const files = fs.readdirSync(targetNestedDirectory);
                
                for (const file of files) {
                    if (file.endsWith('.xlsx')) {
                        const filePath = path.join(targetNestedDirectory, file);
                        const workbook = xlsx.readFile(filePath);
                        const sheet = workbook.Sheets[workbook.SheetNames[0]];
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

                            // Format and write output CSV file only if there are rejected records
                            if (filteredRows.length > 0) {
                                const csvContent = [
                                    title,
                                    section,
                                    headers.join(','),
                                    ...filteredRows.map(row => row.map(val => {
                                        if (val === undefined || val === null) return '';
                                        const str = String(val);
                                        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                                            return `"${str.replace(/"/g, '""')}"`;
                                        }
                                        return str;
                                    }).join(','))
                                ].join('\n') + '\n';

                                const csvPath = filePath.replace(/\.xlsx$/, '.csv');
                                fs.writeFileSync(csvPath, csvContent, 'utf8');
                            }

                        }

                        // Clean up the .xlsx file
                        fs.unlinkSync(filePath);
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
