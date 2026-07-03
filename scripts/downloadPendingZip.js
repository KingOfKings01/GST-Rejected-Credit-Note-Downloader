const path = require('path');

const fs = require('fs');

const { fillImsForm } = require('./formHandler');

const { waitForDimmer, formatFolderMonth, unzipFile, fixDatesInExcelFile } = require('./utils');



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

 * Direct download handler for pre-generated zip files.

 * Navigates straight to the pending category details page and triggers the second download link.

 */

async function downloadPendingZip(page, selections, downloadFolder, client, pendingInfo) {

    const { clientName, stateName } = client;

    const month = pendingInfo.month;



    console.log(`Starting direct download workflow for pending zip: ${clientName} (${stateName}), month: ${month}`);



    try {

        // --- 1. POPUP HANDLER SECTION ---

        let popupButton = page.locator('button, a, input, [role="button"]', { hasText: /remind me later/i }).first();

        try {

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

            } catch (fallbackErr) {}

        }



        // --- 2. NAVIGATE TO IMS DASHBOARD ---

        const imsSelector = 'a[href="//return.gst.gov.in/imsweb/auth/imsDashboard"]';

        await retry(async () => {

            await page.waitForSelector(imsSelector, { state: 'attached', timeout: 30000 });

            await page.evaluate((selector) => {

                const element = document.querySelector(selector);

                if (element) element.click();

            }, imsSelector);

            await page.waitForLoadState('networkidle');

        });



        // --- 3. SELECT OUTWARD SUPPLIES VIEW ---

        await retry(async () => {

            const viewButton = page.locator('button[data-ng-click*="outwardsupplies"]').first();

            await viewButton.waitFor({ state: 'visible', timeout: 30000 });

            await viewButton.click();

            await page.waitForLoadState('networkidle');

        });



        // --- 4. FILL IMS FORM FOR PENDING MONTH ---

        const formPeriod = {

            financialYear: pendingInfo.financialYear || (selections && selections.financialYearFrom) || '2025-26',

            returnPeriod: month,

            returnType: pendingInfo.returnType || (selections && selections.returnType) || 'GSTR-1/IFF'

        };

        await retry(async () => {

            await fillImsForm(page, formPeriod);

        });



        // Wait for dimmer loader to disappear stably

        await waitForDimmer(page, 300);



        // --- 5. NAVIGATE INTO THE DETAILS ROW ---

        const tableRowsSelector = 'table.table-responsive tbody tr';

        await page.waitForSelector(tableRowsSelector, { state: 'visible', timeout: 30000 });



        // Locate row based on section name

        const pendingSectionText = pendingInfo.pendingSection;

        let rowLink = null;



        // Try to match exact heading text inside the rows

        const rowCount = await page.locator(tableRowsSelector).count();

        for (let r = 0; r < rowCount; r++) {

            const rowLoc = page.locator(tableRowsSelector).nth(r);

            const headingText = (await rowLoc.locator('td:nth-child(2)').innerText()).trim().replace(/\n/g, ' ');

           

            // Check if this row heading contains our target section name

            if (headingText.toLowerCase().includes(pendingSectionText.toLowerCase()) || pendingSectionText.toLowerCase().includes(headingText.toLowerCase())) {

                rowLink = rowLoc.locator('td:nth-child(2) a');

                break;

            }

        }



        if (!rowLink || (await rowLink.count()) === 0) {

            throw new Error(`Could not find summary row link for section: "${pendingSectionText}"`);

        }



        await retry(async () => {

            await rowLink.click();

            await page.waitForLoadState('networkidle');

        });

        // Wait for details page loader to disappear stably
        await waitForDimmer(page, 300);



        // --- 5.5 CLICK INITIAL DOWNLOAD TO INITIATE PORTAL CHECK ---
        const secondDownloadLink = page.locator('span[data-ng-if="advSearchDnld"] a, a:has-text("Click here to")').first();
        const firstDownloadLink = page.locator('a[data-ng-click*="dnldAdvSearchSumOut"], p a:has-text("Download")').first();

        const standardDownloadBtn = page.locator('button[data-ng-click*="exportDataOut"]').first();

        // 1. Check if the page has large-count download button or standard download button
        let isLargeCount = false;
        let isPreGeneratedVisible = false;

        try {
            // Wait up to 10 seconds to see if the second download link is visible
            await secondDownloadLink.waitFor({ state: 'visible', timeout: 10000 });
            isPreGeneratedVisible = true;
            isLargeCount = true;
        } catch (e) {
            // If second link is not visible, check if first download link is visible
            if (await firstDownloadLink.count() > 0 && await firstDownloadLink.isVisible()) {
                isLargeCount = true;
            }
        }

        if (isLargeCount) {
            if (!isPreGeneratedVisible) {
                await retry(async () => {
                    await firstDownloadLink.waitFor({ state: 'visible', timeout: 30000 });
                    await firstDownloadLink.click();
                });
                
                // Wait up to 10 seconds for either the download link or the generation pending alert to appear
                const generationAlert = page.locator('.alert-success', { hasText: /request for file generation has been accepted|File generation is in progress/i }).first();
                let isAlertVisible = false;
                let isLinkVisible = false;
                
                for (let i = 0; i < 20; i++) {
                    if (await secondDownloadLink.isVisible()) {
                        isLinkVisible = true;
                        isAlertVisible = false;
                        break;
                    }
                    if (await generationAlert.isVisible()) {
                        isAlertVisible = true;
                    }
                    await page.waitForTimeout(500);
                }

                if (isAlertVisible && !isLinkVisible) {
                    console.log(`⚠️ Zip file generation request still in progress. Resetting timer for another 20 minutes.`);
                    return { success: false, pending: true, readyAt: Date.now() + 20 * 60 * 1000 };
                }

                // Fallback wait just in case
                if (!isLinkVisible) {
                    await secondDownloadLink.waitFor({ state: 'visible', timeout: 30000 });
                }
            }

            const download = await retry(async () => {
                const downloadPromise = page.waitForEvent('download', { timeout: 45000 });
                await secondDownloadLink.click();
                return await downloadPromise;
            });

            if (download) {
                const suggestedFileName = download.suggestedFilename();
                const financialYear = pendingInfo.financialYear || (selections && selections.financialYearFrom) || '2025-26';
                const folderLabel = formatFolderMonth(month, financialYear);
                const targetNestedDirectory = path.join(downloadFolder, clientName, stateName, folderLabel);

                if (!fs.existsSync(targetNestedDirectory)) {
                    fs.mkdirSync(targetNestedDirectory, { recursive: true });
                }

                const finalSavePath = path.join(targetNestedDirectory, suggestedFileName);
                await download.saveAs(finalSavePath);

                let totalRejectedCount = 0;
                // Extract / unzip and filter
                try {
                    unzipFile(finalSavePath, targetNestedDirectory);
                    const xlsx = require('xlsx');
                    const files = fs.readdirSync(targetNestedDirectory);
                    for (const file of files) {
                        if (file.endsWith('.xlsx')) {
                            const filePath = path.join(targetNestedDirectory, file);
                            const workbook = xlsx.readFile(filePath);
                            const sheet = workbook.Sheets[workbook.SheetNames[0]];
                            
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
                                    try {
                                        fs.unlinkSync(filePath);
                                    } catch (e) {}
                                }
                            } else {
                                try {
                                    fs.unlinkSync(filePath);
                                } catch (e) {}
                            }
                        }
                    }
                } catch (zipErr) {
                    console.error("❌ Failed to process and extract ZIP archive:", zipErr.message || zipErr);
                }

                console.log(`✅ Direct download complete. Rejected count: ${totalRejectedCount}`);
                return { success: true, rejectedCount: totalRejectedCount };
            }
        } else {
            // Standard small-count flow fallback
            console.log(`System: This section does not require zip generation (count <= 500). Applying fallback direct Excel export...`);
            const filterButton = page.locator('button#showbutton');
            await retry(async () => {
                await filterButton.waitFor({ state: 'visible', timeout: 10000 });
                await filterButton.click();
            });
            await page.waitForTimeout(300);

            const statusDropdown = page.locator('select[ng-model="status"]');
            await retry(async () => {
                await statusDropdown.waitFor({ state: 'visible', timeout: 10000 });
                await statusDropdown.click();
                await page.waitForTimeout(300);
                await statusDropdown.selectOption({ value: 'Rejected', label: 'Rejected' });
                await statusDropdown.dispatchEvent('change');
            });
            await page.waitForTimeout(300);

            const applyButton = page.locator('button[ng-click*="suppfilterButton"]');
            await retry(async () => {
                await applyButton.waitFor({ state: 'visible', timeout: 10000 });
                await applyButton.click();
            });

            await page.waitForLoadState('networkidle');
            await waitForDimmer(page, 300);

            let isDataPresent = false;
            const noRecordAlert = page.locator('.alert-danger', { hasText: /No record found/i }).first();
            
            for (let attempt = 0; attempt < 20; attempt++) {
                if (await noRecordAlert.isVisible()) {
                    isDataPresent = false;
                    break;
                }
                if (await standardDownloadBtn.isVisible() && await standardDownloadBtn.isEnabled()) {
                    isDataPresent = true;
                    break;
                }
                await page.waitForTimeout(300);
            }

            if (isDataPresent) {
                const downloadPromise = page.waitForEvent('download', { timeout: 15000 }).catch((e) => {
                    return null;
                });
                await standardDownloadBtn.click();
                const download = await downloadPromise;
                if (download) {
                    const suggestedFileName = download.suggestedFilename();
                    const financialYear = pendingInfo.financialYear || (selections && selections.financialYearFrom) || '2025-26';
                    const folderLabel = formatFolderMonth(month, financialYear);
                    const targetNestedDirectory = path.join(downloadFolder, clientName, stateName, folderLabel);
                    if (!fs.existsSync(targetNestedDirectory)) {
                        fs.mkdirSync(targetNestedDirectory, { recursive: true });
                    }
                    const finalSavePath = path.join(targetNestedDirectory, suggestedFileName);
                    await download.saveAs(finalSavePath);
                    fixDatesInExcelFile(finalSavePath);

                    // Read excel to count rejected records
                    const xlsx = require('xlsx');
                    let totalRejectedCount = 0;
                    try {
                        const workbook = xlsx.readFile(finalSavePath);
                        const sheet = workbook.Sheets[workbook.SheetNames[0]];
                        const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
                        if (rows.length >= 5) {
                            totalRejectedCount = rows.length - 6; // Excluding headings
                        }
                    } catch (err) {}

                    console.log(`✅ Direct download complete (standard Excel export). Rejected count: ${totalRejectedCount}`);
                    return { success: true, rejectedCount: totalRejectedCount };
                }
            } else {
                console.log("No rejected records found for this section.");
                return { success: true, rejectedCount: 0 };
            }
        }

        return { success: false };

    } catch (err) {

        console.error("❌ Error during direct zip download exception:", err.message || err);

        return { success: false, error: err.message || String(err) };

    }

}



module.exports = { downloadPendingZip };