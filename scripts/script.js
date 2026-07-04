const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright-core'); // Imported chromium here to manage its lifecycle from playwright-core
const xlsx = require('xlsx');
const { loginGST } = require('./login');
const { doGstWork, getTargetPeriodsQueue } = require('./work');

function formatMonthYear(monthStr, financialYear) {
    if (!financialYear) return monthStr;
    const FISCAL_MONTHS = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"];
    const shortMonth = monthStr.substring(0, 3);
    const parts = financialYear.split('-');
    const startYear = parseInt(parts[0], 10);
    const endYearStr = parts[1];
    const monthIdx = FISCAL_MONTHS.indexOf(monthStr);
    
    if (monthIdx < 9) {
        return `${shortMonth} ${String(startYear).slice(-2)}`;
    } else {
        return `${shortMonth} ${endYearStr}`;
    }
}

function createBaseReportRow(client, selections, allReportRows) {
    return {
        'Sr. No.': allReportRows.length + 1,
        'Client Name': client.clientName || client.clientState || 'Unknown',
        'GST Number': client.gstNo || 'N/A',
        'Username': client.username,
        'Financial Year From': selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26',
        'Financial Year To': selections ? (selections.financialYearTo || selections.financialYear || '2025-26') : '2025-26',
        'Return Period From': selections ? (selections.returnPeriodFrom || selections.returnPeriod || 'Unknown') : 'Unknown',
        'Return Period To': selections ? (selections.returnPeriodTo || selections.returnPeriod || 'Unknown') : 'Unknown',
    };
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

    const targetPeriodsQueue = getTargetPeriodsQueue(selections || {});
    const dynamicMonthColumns = targetPeriodsQueue.map(p => formatMonthYear(p.returnPeriod, p.financialYear));
    const reportColumns = [
        'Sr. No.',
        'Client Name',
        'GST Number',
        'Username',
        'Financial Year From',
        'Financial Year To',
        'Return Period From',
        'Return Period To',
        ...dynamicMonthColumns,
        'Error (if present)',
        'Timestamp'
    ];

    const browserZoom = (selections && selections.browserZoom) || '0.85';
    const runHeadless = selections && selections.runHeadless === true;

    // CRITICAL ENGINE OPTIMIZATION: Launch ONE browser instance for the entire job run
    let browser;
    try {
        browser = await chromium.launch({ 
            headless: runHeadless,
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
                headless: runHeadless,
                channel: 'msedge',
                args: [
                    '--start-maximized',
                    `--force-device-scale-factor=${browserZoom}`
                ]
            });
        } catch (edgeErr) {
            browser = await chromium.launch({ 
                headless: runHeadless,
                args: [
                    '--start-maximized',
                    `--force-device-scale-factor=${browserZoom}`
                ]
            });
        }
    }

    const captchaResolvers = new Map(); // username -> resolveCallback

    process.stdin.on('data', (data) => {
        const line = data.toString().trim();
        const matchSolve = line.match(/^SOLVED_CAPTCHA:([^:]+):(.+)/);
        const matchRefresh = line.match(/^REFRESH_CAPTCHA:(.+)/);
        const matchSkip = line.match(/^SKIP_CLIENT:(.+)/);
        
        if (matchSolve) {
            const [_, username, captchaText] = matchSolve;
            const resolve = captchaResolvers.get(username);
            if (resolve) {
                resolve({ action: 'solve', text: captchaText });
                captchaResolvers.delete(username);
            }
        } else if (matchRefresh) {
            const [_, username] = matchRefresh;
            const resolve = captchaResolvers.get(username);
            if (resolve) {
                resolve({ action: 'refresh' });
                captchaResolvers.delete(username);
            }
        } else if (matchSkip) {
            const [_, username] = matchSkip;
            const resolve = captchaResolvers.get(username);
            if (resolve) {
                resolve({ action: 'skip' });
                captchaResolvers.delete(username);
            }
        }
    });

    // Concurrency limit setting
    const concurrencyLimit = parseInt(selections.maxBrowsers || 3, 10);
    console.log(`System: Running with ${concurrencyLimit} parallel browser instances.`);

    // Create a queue of clients with their indices
    const clientQueue = clients.map((c, i) => ({ client: c, index: i }));

    const runWorker = async () => {
        while (clientQueue.length > 0) {
            const item = clientQueue.shift();
            if (!item) break;
            const { client, index } = item;

            console.log(`CLIENT_PROGRESS:START:${index+1}:${clients.length}:${client.username}:${client.clientName || client.username}`);

            // Create a new context and page for EACH client to ensure absolute clean session isolation
            const context = await browser.newContext({
                viewport: null
            });
            const page = await context.newPage();

            try {
                // 1. Run the login process using the active page
                const isLoginSuccessful = await loginGST(page, client.username, client.password, browserZoom, captchaResolvers);

                // 2. If login failed or skipped, handle appropriately
                if (!isLoginSuccessful || isLoginSuccessful === 'skipped') {
                    const isSkipped = isLoginSuccessful === 'skipped';
                    const statusText = isSkipped ? 'Skipped by user' : 'Failed: Login unsuccessful';
                    console.log(`CLIENT_PROGRESS:FAILED:${client.username}: ${statusText}`);
                    updateExcelClientStatus(excelFilePath, client.username, statusText);
                    allReportRows.push({
                        'Sr. No.': index + 1,
                        'Client Name': client.clientName || client.clientState || 'Unknown',
                        'GST Number': client.gstNo || 'N/A',
                        'Username': client.username,
                        'Financial Year From': selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26',
                        'Financial Year To': selections ? (selections.financialYearTo || selections.financialYear || '2025-26') : '2025-26',
                        'Return Period From': selections ? (selections.returnPeriodFrom || selections.returnPeriod || 'Unknown') : 'Unknown',
                        'Return Period To': selections ? (selections.returnPeriodTo || selections.returnPeriod || 'Unknown') : 'Unknown',
                        [selections ? selections.returnPeriod : 'Selected Month']: isSkipped ? 'Skipped' : 'Login Failed',
                        'Error (if present)': statusText,
                        'Timestamp': new Date().toLocaleString()
                    });
                    continue;
                }

                // Decode ExcelStatus for zip pending
                let pendingZips = [];
                if (client.excelZipStatus) {
                    pendingZips = client.excelZipStatus.split(';;').map(s => s.trim()).filter(s => s.startsWith('Zip Pending'));
                }

                let readyZips = [];
                let remainingZips = [];
                
                for (let idx = 0; idx < pendingZips.length; idx++) {
                    const parts = pendingZips[idx].split('|').map(s => s.trim());
                    if (parts.length >= 4) {
                        const readyAt = parseInt(parts[2], 10);
                        if (Date.now() >= readyAt) {
                            readyZips.push({
                                index: idx,
                                month: parts[1],
                                readyAt: readyAt,
                                pendingSection: parts[3],
                                financialYear: parts[4] || (selections && selections.financialYearFrom) || '2025-26',
                                returnType: parts[5] || (selections && selections.returnType) || 'GSTR-1/IFF',
                                originalString: pendingZips[idx]
                            });
                        } else {
                            remainingZips.push(pendingZips[idx]);
                        }
                    } else {
                        remainingZips.push(pendingZips[idx]);
                    }
                }

                if (readyZips.length > 0) {
                    // --- EXCEPTION FLOW: Direct zip download loop ---
                    console.log(`System: Client has ${readyZips.length} pre-generated zip(s) ready to download. Running loop...`);
                    const { downloadPendingZip } = require('./downloadPendingZip');
                    
                    let downloadedCount = 0;
                    let lastError = null;
                    
                    for (const clientPending of readyZips) {
                        console.log(`Starting direct download for pending zip: month ${clientPending.month}, section: ${clientPending.pendingSection}`);
                        const directResult = await downloadPendingZip(page, selections, downloadFolder, client, clientPending);
                        
                        await page.waitForTimeout(1000);

                        const isDirectDownloadSuccess = directResult && directResult.success;
                        
                        if (isDirectDownloadSuccess) {
                            downloadedCount++;
                            console.log(`CLIENT_PROGRESS:ZIP_DOWNLOADED:${client.username}:${clientPending.month}:${clientPending.pendingSection}`);
                            
                            // Construct report row
                            const row = {
                                'Sr. No.': index + 1,
                                'Client Name': client.clientName || client.clientState || 'Unknown',
                                'GST Number': client.gstNo || 'N/A',
                                'Username': client.username,
                                'Financial Year From': selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26',
                                'Financial Year To': selections ? (selections.financialYearTo || selections.financialYear || '2025-26') : '2025-26',
                                'Return Period From': selections ? (selections.returnPeriodFrom || selections.returnPeriod || 'Unknown') : 'Unknown',
                                'Return Period To': selections ? (selections.returnPeriodTo || selections.returnPeriod || 'Unknown') : 'Unknown',
                            };
                            const monthCol = formatMonthYear(clientPending.month, clientPending.financialYear || (selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26'));
                            row[monthCol] = `1 Download: ${clientPending.pendingSection} (${directResult.rejectedCount} records)`;
                            row['Error (if present)'] = '';
                            row['Timestamp'] = new Date().toLocaleString();
                            allReportRows.push(row);
                        } else if (directResult && directResult.pending) {
                            const newPendingString = `Zip Pending | ${clientPending.month} | ${directResult.readyAt} | ${clientPending.pendingSection} | ${clientPending.financialYear} | ${clientPending.returnType}`;
                            remainingZips.push(newPendingString);
                            console.log(`CLIENT_PROGRESS:ZIP_PENDING:${client.username}:${clientPending.month}:${directResult.readyAt}:${clientPending.pendingSection}`);
                            
                            const row = {
                                'Sr. No.': index + 1,
                                'Client Name': client.clientName || client.clientState || 'Unknown',
                                'GST Number': client.gstNo || 'N/A',
                                'Username': client.username,
                                'Financial Year From': selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26',
                                'Financial Year To': selections ? (selections.financialYearTo || selections.financialYear || '2025-26') : '2025-26',
                                'Return Period From': selections ? (selections.returnPeriodFrom || selections.returnPeriod || 'Unknown') : 'Unknown',
                                'Return Period To': selections ? (selections.returnPeriodTo || selections.returnPeriod || 'Unknown') : 'Unknown',
                            };
                            const monthCol = formatMonthYear(clientPending.month, clientPending.financialYear || (selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26'));
                            row[monthCol] = `Zip Pending (Reset for another 20 mins)`;
                            row['Error (if present)'] = 'Zip file generation still in progress';
                            row['Timestamp'] = new Date().toLocaleString();
                            allReportRows.push(row);
                        } else {
                            lastError = (directResult && directResult.error) || 'Failed to download zip';
                            remainingZips.push(clientPending.originalString);
                            
                            const row = {
                                'Sr. No.': index + 1,
                                'Client Name': client.clientName || client.clientState || 'Unknown',
                                'GST Number': client.gstNo || 'N/A',
                                'Username': client.username,
                                'Financial Year From': selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26',
                                'Financial Year To': selections ? (selections.financialYearTo || selections.financialYear || '2025-26') : '2025-26',
                                'Return Period From': selections ? (selections.returnPeriodFrom || selections.returnPeriod || 'Unknown') : 'Unknown',
                                'Return Period To': selections ? (selections.returnPeriodTo || selections.returnPeriod || 'Unknown') : 'Unknown',
                            };
                            const monthCol = formatMonthYear(clientPending.month, clientPending.financialYear || (selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26'));
                            row[monthCol] = `Failed: ${lastError}`;
                            row['Error (if present)'] = lastError;
                            row['Timestamp'] = new Date().toLocaleString();
                            allReportRows.push(row);
                        }
                    }
                    
                    const newZipStatus = remainingZips.join(' ;; ');
                    updateExcelClientStatus(excelFilePath, client.username, newZipStatus, 'ZIP STATUS');
                    
                    if (remainingZips.length > 0) {
                        updateExcelClientStatus(excelFilePath, client.username, 'Zip Pending', 'STATUS');
                    } else if (downloadedCount === readyZips.length) {
                        updateExcelClientStatus(excelFilePath, client.username, 'Success', 'STATUS');
                        console.log(`CLIENT_PROGRESS:SUCCESS:${client.username}`);
                    } else {
                        updateExcelClientStatus(excelFilePath, client.username, `Failed: ${lastError || 'Unknown zip error'}`, 'STATUS');
                        console.log(`CLIENT_PROGRESS:FAILED:${client.username}: ${lastError || 'Unknown zip error'}`);
                    }
                } else {
                    // --- STANDARD FLOW ---
                    const monthlyResults = await doGstWork(page, selections, downloadFolder, client, excelFilePath);

                    await page.waitForTimeout(1000);

                    // 5. Update Status based on monthly results
                    const clientReportRow = {
                        'Sr. No.': index + 1,
                        'Client Name': client.clientName || client.clientState || 'Unknown',
                        'GST Number': client.gstNo || 'N/A',
                        'Username': client.username,
                        'Financial Year From': selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26',
                        'Financial Year To': selections ? (selections.financialYearTo || selections.financialYear || '2025-26') : '2025-26',
                        'Return Period From': selections ? (selections.returnPeriodFrom || selections.returnPeriod || 'Unknown') : 'Unknown',
                        'Return Period To': selections ? (selections.returnPeriodTo || selections.returnPeriod || 'Unknown') : 'Unknown',
                    };
                    
                    let hasAnyError = false;
                    let firstErrorMessage = '';
                    let isZipPending = false;
                    let pendingStatuses = [];

                    monthlyResults.forEach(res => {
                        const monthCol = formatMonthYear(res.month, res.financialYear || (selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26'));
                        let monthStatusParts = [];

                        if (res.isZipPending) {
                            isZipPending = true;
                            monthStatusParts.push(`Zip Pending (Revisit in 20 mins)`);
                            console.log(`CLIENT_PROGRESS:ZIP_PENDING:${client.username}:${res.month}:${res.zipReadyAt}:${res.pendingSection}`);
                            
                            const fy = (selections && selections.financialYearFrom) || '2025-26';
                            const rt = (selections && selections.returnType) || 'GSTR-1/IFF';
                            const singlePending = `Zip Pending | ${res.month} | ${res.zipReadyAt} | ${res.pendingSection} | ${fy} | ${rt}`;
                            pendingStatuses.push(singlePending);
                        }
                        
                        if (res.downloads && res.downloads.length > 0) {
                            const downloadCount = res.downloads.length;
                            const details = res.downloads.map(d => `${d.section} (${d.records} records)`).join(', ');
                            monthStatusParts.push(`${downloadCount} Download${downloadCount > 1 ? 's' : ''}: ${details}`);
                        }
                        
                        if (res.error) {
                            hasAnyError = true;
                            if (!firstErrorMessage) firstErrorMessage = res.error;
                            monthStatusParts.push(`Failed: ${res.error}`);
                        }
                        
                        if (monthStatusParts.length > 0) {
                            clientReportRow[monthCol] = monthStatusParts.join(' | ');
                        } else {
                            clientReportRow[monthCol] = 'Checked (No Data)';
                        }
                    });

                    clientReportRow['Error (if present)'] = firstErrorMessage;
                    clientReportRow['Timestamp'] = new Date().toLocaleString();
                    allReportRows.push(clientReportRow);

                    const finalStatus = isZipPending ? 'Zip Pending' : (hasAnyError ? `Failed: ${firstErrorMessage}` : 'Success');
                    updateExcelClientStatus(excelFilePath, client.username, finalStatus, 'STATUS');
                    
                    if (isZipPending) {
                        const existingZips = client.excelZipStatus ? client.excelZipStatus.split(';;').map(s => s.trim()).filter(s => s.startsWith('Zip Pending')) : [];
                        const allPendingZips = [...existingZips, ...pendingStatuses];
                        updateExcelClientStatus(excelFilePath, client.username, allPendingZips.join(' ;; '), 'ZIP STATUS');
                    } else {
                        updateExcelClientStatus(excelFilePath, client.username, '', 'ZIP STATUS');
                    }

                    const failedMonthsList = monthlyResults
                        .filter(res => res.error)
                        .map(res => res.month);
                    
                    if (failedMonthsList.length > 0) {
                        updateExcelClientStatus(excelFilePath, client.username, failedMonthsList.join(', '), 'FAILED MONTHS');
                    } else {
                        updateExcelClientStatus(excelFilePath, client.username, '', 'FAILED MONTHS');
                    }
                    
                    if (isZipPending) {
                        // Status remains zip_pending, do not overwrite to failed
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
                
                const row = {
                    'Sr. No.': index + 1,
                    'Client Name': client.clientName || client.clientState || 'Unknown',
                    'GST Number': client.gstNo || 'N/A',
                    'Username': client.username,
                    'Financial Year From': selections ? (selections.financialYearFrom || selections.financialYear || '2025-26') : '2025-26',
                    'Financial Year To': selections ? (selections.financialYearTo || selections.financialYear || '2025-26') : '2025-26',
                    'Return Period From': selections ? (selections.returnPeriodFrom || selections.returnPeriod || 'Unknown') : 'Unknown',
                    'Return Period To': selections ? (selections.returnPeriodTo || selections.returnPeriod || 'Unknown') : 'Unknown',
                };
                row['Error (if present)'] = errStr;
                row['Timestamp'] = new Date().toLocaleString();
                allReportRows.push(row);
            } finally {
                try {
                    await page.close();
                } catch (e) {}
                try {
                    await context.close();
                } catch (e) {}
            }
        }
    };

    // Launch all workers concurrently
    const workers = Array.from({ length: concurrencyLimit }, () => runWorker());
    await Promise.all(workers);

    // --- FINAL WORKFLOW TEARDOWN ---
    await browser.close();

    // Generate Final Excel Report
    try {
        if (allReportRows.length > 0) {
            const reportWb = xlsx.utils.book_new();
            const reportWs = xlsx.utils.json_to_sheet(allReportRows, { header: reportColumns });
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
        process.exit(0);
    } catch (reportErr) {
        console.error("Failed to generate Excel report:", reportErr);
        process.exit(1);
    }
}

// Execute everything
run();