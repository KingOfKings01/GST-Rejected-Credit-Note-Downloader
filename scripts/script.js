const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright'); // Imported chromium here to manage its lifecycle
const xlsx = require('xlsx');
const { loginGST } = require('./login');
const { doGstWork } = require('./work');

function updateExcelClientStatus(filePath, username, newStatus) {
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
            statusColIndex = headerRow.indexOf('STATUS');
            if (statusColIndex === -1) {
                statusColIndex = headerRow.length;
                const headerCellRef = xlsx.utils.encode_cell({ r: headerRowIndex, c: statusColIndex });
                sheet[headerCellRef] = { t: 's', v: 'STATUS' };
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
    if (attempts === 0) {
        console.error(`❌ ERROR: Could not update Excel status for ${username} because the file is locked. Please close the Excel sheet before running!`);
    }
}

async function run() {
    // Check if configuration path is passed
    const configPath = process.argv[2];
    
    let config;
    try {
        config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    } catch (e) {
        console.error("Failed to parse configuration file:", e);
        process.exit(1);
    }

    const { clients, selections, downloadFolder, excelFilePath } = config;
    const allReportRows = [];

    const browserZoom = (selections && selections.browserZoom) || '0.85';

    // CRITICAL ENGINE OPTIMIZATION: Launch ONE browser instance for the entire job run
    let browser;
    try {
        browser = await chromium.launch({ 
            headless: false,
            channel: 'chrome',
            args: [
                '--start-maximized',
                // Natively forces page/UI scale factor (zoom) across all page navigations.
                `--force-device-scale-factor=${browserZoom}`
            ]
        });
    } catch (chromeErr) {
        try {
            browser = await chromium.launch({ 
                headless: false,
                channel: 'msedge',
                args: [
                    '--start-maximized',
                    `--force-device-scale-factor=${browserZoom}`
                ]
            });
        } catch (edgeErr) {
            browser = await chromium.launch({ 
                headless: false,
                args: [
                    '--start-maximized',
                    `--force-device-scale-factor=${browserZoom}`
                ]
            });
        }
    }

    for (let i = 0; i < clients.length; i++) {
        const client = clients[i];
        console.log(`CLIENT_PROGRESS:START:${i+1}:${clients.length}:${client.username}:${client.clientName || client.username}`);

        // Create a new context and page for EACH client to ensure absolute clean session isolation
        const context = await browser.newContext({
            viewport: null
        });
        const page = await context.newPage();

        try {
            // 1. Run the login process using the active page
            const browserZoom = (selections && selections.browserZoom) || '0.85';
            const isLoginSuccessful = await loginGST(page, client.username, client.password, browserZoom);

            // 2. If login failed or skipped, handle appropriately
            if (!isLoginSuccessful || isLoginSuccessful === 'skipped') {
                const isSkipped = isLoginSuccessful === 'skipped';
                const statusText = isSkipped ? 'Skipped by user' : 'Failed: Login unsuccessful';
                console.log(`CLIENT_PROGRESS:FAILED:${client.username}: ${statusText}`);
                updateExcelClientStatus(excelFilePath, client.username, statusText);
                allReportRows.push({
                    'Client Name': client.clientName || client.clientState || 'Unknown',
                    'State': client.stateName || 'N/A',
                    'GST Number': client.gstNo || 'N/A',
                    'Username': client.username,
                    [selections ? selections.returnPeriod : 'Selected Month']: isSkipped ? 'Skipped' : 'Login Failed',
                    'Error if Present': statusText,
                    'Timestamp': new Date().toLocaleString()
                });
                continue;
            }

            // Decode ExcelStatus for zip pending
            let clientPending = null;
            let hasReadyZip = false;
            
            if (client.excelStatus && client.excelStatus.startsWith('Zip Pending')) {
                const parts = client.excelStatus.split('|').map(s => s.trim());
                if (parts.length >= 4) {
                    const month = parts[1];
                    const readyAt = parseInt(parts[2], 10);
                    const pendingSection = parts[3];
                    const financialYear = parts[4] || (selections && selections.financialYear) || '2025-26';
                    const returnType = parts[5] || (selections && selections.returnType) || 'GSTR-1/IFF';
                    
                    clientPending = {
                        username: client.username,
                        month,
                        readyAt,
                        pendingSection,
                        financialYear,
                        returnType
                    };
                    
                    if (Date.now() >= readyAt) {
                        hasReadyZip = true;
                    }
                }
            }

            if (clientPending && !hasReadyZip) {
                const diffSeconds = Math.max(0, Math.floor((clientPending.readyAt - Date.now()) / 1000));
                const mins = Math.floor(diffSeconds / 60);
                const secs = diffSeconds % 60;
                const waitMsg = `Zip Pending: Revisit in ${mins}m ${secs}s`;
                console.log(`CLIENT_PROGRESS:FAILED:${client.username}: Zip generation in progress. ${mins}m ${secs}s remaining.`);
                allReportRows.push({
                    'Client Name': client.clientName || client.clientState || 'Unknown',
                    'State': client.stateName || 'N/A',
                    'GST Number': client.gstNo || 'N/A',
                    'Username': client.username,
                    [clientPending.month]: `Zip Pending (${mins}m remaining)`,
                    'Error if Present': 'Zip generation in progress',
                    'Timestamp': new Date().toLocaleString()
                });
                continue;
            }

            if (hasReadyZip) {
                // --- EXCEPTION FLOW: Direct zip download ---
                console.log(`System: Client has a pre-generated zip ready to download. Running isolated script...`);
                const { downloadPendingZip } = require('./downloadPendingZip');
                const directResult = await downloadPendingZip(page, selections, downloadFolder, client, clientPending);
                
                await page.waitForTimeout(1000);

                const isDirectDownloadSuccess = directResult && directResult.success;
                const directDownloadError = (directResult && directResult.error) || 'Failed to download zip';

                const clientReportRow = {
                    'Client Name': client.clientName || client.clientState || 'Unknown',
                    'State': client.stateName || 'N/A',
                    'GST Number': client.gstNo || 'N/A',
                    'Username': client.username,
                    [clientPending.month]: isDirectDownloadSuccess 
                        ? `1 Download: ${clientPending.pendingSection} (${directResult.rejectedCount} records)` 
                        : `Failed: ${directDownloadError}`,
                    'Error if Present': isDirectDownloadSuccess ? '' : directDownloadError,
                    'Timestamp': new Date().toLocaleString()
                };
                allReportRows.push(clientReportRow);

                const finalStatus = isDirectDownloadSuccess ? 'Success' : `Failed: ${directDownloadError}`;
                updateExcelClientStatus(excelFilePath, client.username, finalStatus);

                if (isDirectDownloadSuccess) {
                    console.log(`CLIENT_PROGRESS:SUCCESS:${client.username}`);
                } else {
                    console.log(`CLIENT_PROGRESS:FAILED:${client.username}: ${directDownloadError}`);
                }
            } else {
                // --- STANDARD FLOW ---
                const monthlyResults = await doGstWork(page, selections, downloadFolder, client);

                await page.waitForTimeout(1000);

                // 5. Update Status based on monthly results
                let hasAnyError = false;
                let isZipPending = false;
                let firstErrorMessage = '';
                
                const clientReportRow = {
                    'Client Name': client.clientName || client.clientState || 'Unknown',
                    'State': client.stateName || 'N/A',
                    'GST Number': client.gstNo || 'N/A',
                    'Username': client.username,
                };

                let pendingStatusString = '';

                monthlyResults.forEach(res => {
                    if (res.isZipPending) {
                        isZipPending = true;
                        clientReportRow[res.month] = `Zip Pending (Revisit in 20 mins)`;
                        console.log(`CLIENT_PROGRESS:ZIP_PENDING:${client.username}:${res.month}:${res.zipReadyAt}:${res.pendingSection}`);
                        
                        const fy = (selections && selections.financialYear) || '2025-26';
                        const rt = (selections && selections.returnType) || 'GSTR-1/IFF';
                        pendingStatusString = `Zip Pending | ${res.month} | ${res.zipReadyAt} | ${res.pendingSection} | ${fy} | ${rt}`;
                    } else if (res.error) {
                        hasAnyError = true;
                        if (!firstErrorMessage) firstErrorMessage = res.error;
                        clientReportRow[res.month] = `Failed: ${res.error}`;
                    } else if (res.downloads && res.downloads.length > 0) {
                        const downloadCount = res.downloads.length;
                        const details = res.downloads.map(d => `${d.section} (${d.records} records)`).join(', ');
                        clientReportRow[res.month] = `${downloadCount} Download${downloadCount > 1 ? 's' : ''}: ${details}`;
                    } else {
                        clientReportRow[res.month] = 'Checked (No Data)';
                    }
                });

                clientReportRow['Error if Present'] = firstErrorMessage;
                clientReportRow['Timestamp'] = new Date().toLocaleString();
                allReportRows.push(clientReportRow);

                const finalStatus = isZipPending ? pendingStatusString : (hasAnyError ? `Failed: ${firstErrorMessage}` : 'Success');
                updateExcelClientStatus(excelFilePath, client.username, finalStatus);
                
                if (isZipPending) {
                    console.log(`CLIENT_PROGRESS:FAILED:${client.username}: Zip Pending`);
                } else if (hasAnyError) {
                    console.log(`CLIENT_PROGRESS:FAILED:${client.username}: ${firstErrorMessage}`);
                } else {
                    console.log(`CLIENT_PROGRESS:SUCCESS:${client.username}`);
                }
            }

        } catch (err) {
            console.error(`CLIENT_PROGRESS:FAILED:${client.username}: Error occurred:`, err.message || err);
            const errStr = err.message || String(err);
            updateExcelClientStatus(excelFilePath, client.username, `Failed: ${errStr}`);
            allReportRows.push({
                'Client Name': client.clientName || client.clientState || 'Unknown',
                'State': client.stateName || 'N/A',
                'GST Number': client.gstNo || 'N/A',
                'Username': client.username,
                [selections ? selections.returnPeriod : 'Selected Month']: `Failed: ${errStr}`,
                'Error if Present': errStr,
                'Timestamp': new Date().toLocaleString()
            });
        } finally {
            try {
                await page.close();
            } catch (e) {}
            try {
                await context.close();
            } catch (e) {}
        }
    }

    // --- FINAL WORKFLOW TEARDOWN ---
    await browser.close();

    // Generate Final Excel Report
    try {
        if (allReportRows.length > 0) {
            const reportWb = xlsx.utils.book_new();
            const reportWs = xlsx.utils.json_to_sheet(allReportRows);
            xlsx.utils.book_append_sheet(reportWb, reportWs, 'Automation Report');

            // Format date for file name: YYYY-MM-DD_HH-MM-SS
            const now = new Date();
            const dateStr = now.getFullYear() + '-' + 
                            String(now.getMonth() + 1).padStart(2, '0') + '-' + 
                            String(now.getDate()).padStart(2, '0') + '_' + 
                            String(now.getHours()).padStart(2, '0') + '-' + 
                            String(now.getMinutes()).padStart(2, '0') + '-' + 
                            String(now.getSeconds()).padStart(2, '0');
            const reportFileName = `GST_Automation_Report_${dateStr}.xlsx`;
            const reportFilePath = path.join(downloadFolder, reportFileName);

            // Ensure downloadFolder exists
            if (!fs.existsSync(downloadFolder)) {
                fs.mkdirSync(downloadFolder, { recursive: true });
            }

            xlsx.writeFile(reportWb, reportFilePath);
        }

        console.log("All clients processed successfully!");
    } catch (reportErr) {
        console.error("Failed to generate Excel report:", reportErr);
    }
}

// Execute everything
run();