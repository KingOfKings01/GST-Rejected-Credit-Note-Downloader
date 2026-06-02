const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

/**
 * Traverses the source directory to find all client and state subdirectories,
 * and consolidates all rejected notes files per state into a single Excel report.
 * @param {string} sourceDir - The main download/output directory
 * @param {function} logCallback - Function to log progress messages back to the UI
 */
async function consolidateReports(sourceDir, logCallback = console.log) {
    if (!sourceDir || !fs.existsSync(sourceDir)) {
        throw new Error(`Source directory "${sourceDir}" does not exist.`);
    }

    logCallback(`Starting consolidation process in: ${sourceDir}`);

    // Get list of clients
    const clients = fs.readdirSync(sourceDir).filter(name => {
        const fullPath = path.join(sourceDir, name);
        return fs.statSync(fullPath).isDirectory() && !name.startsWith('.') && name !== 'Consolidated';
    });

    if (clients.length === 0) {
        logCallback('No client folders found to consolidate.');
        return;
    }

    let totalConsolidated = 0;

    for (const client of clients) {
        const clientPath = path.join(sourceDir, client);
        const states = fs.readdirSync(clientPath).filter(name => {
            const fullPath = path.join(clientPath, name);
            return fs.statSync(fullPath).isDirectory() && !name.startsWith('.');
        });

        for (const state of states) {
            const statePath = path.join(clientPath, state);
            const months = fs.readdirSync(statePath).filter(name => {
                const fullPath = path.join(statePath, name);
                return fs.statSync(fullPath).isDirectory() && name !== 'Consolidated' && !name.startsWith('.');
            });

            // Group all rejected notes by section/category
            const sectionData = {}; // Format: { [sectionName]: { headers: [], rows: [], title: '' } }

            let username = client; // Default username fallback
            
            for (const month of months) {
                const monthPath = path.join(statePath, month);
                const files = fs.readdirSync(monthPath).filter(name => name.endsWith('.xlsx') || name.endsWith('.csv'));

                for (const file of files) {
                    const filePath = path.join(monthPath, file);
                    try {
                        const workbook = xlsx.readFile(filePath);
                        const firstSheetName = workbook.SheetNames[0];
                        const sheet = workbook.Sheets[firstSheetName];

                        // Recalculate sheet ref range if needed
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

                        // Dynamically find the header row by searching for "Status", "GSTIN", or "Remarks"
                        let headerIdx = -1;
                        for (let r = 0; r < Math.min(10, rows.length); r++) {
                            const row = rows[r];
                            if (row && row.some(cell => cell && (
                                String(cell).toLowerCase().includes('status') || 
                                String(cell).toLowerCase().includes('gstin') ||
                                String(cell).toLowerCase().includes('remarks')
                            ))) {
                                headerIdx = r;
                                break;
                            }
                        }

                        if (headerIdx !== -1 && rows.length > headerIdx) {
                            // Extract metadata dynamically
                            const rawTitle = rows[0] && rows[0][0] ? String(rows[0][0]) : '';
                            const sectionName = (headerIdx > 0 && rows[headerIdx - 1] && rows[headerIdx - 1][0]) 
                                ? String(rows[headerIdx - 1][0]).trim() 
                                : path.basename(file, path.extname(file));
                            const headers = rows[headerIdx] || [];
                            
                            // Try to extract GSTIN/username from title or data rows
                            if (rawTitle && rawTitle.includes('GSTIN')) {
                                const gstinMatch = rawTitle.match(/GSTIN\s*:\s*([A-Z0-9]+)/i);
                                if (gstinMatch) username = gstinMatch[1];
                            }

                            // Data rows start immediately after the header row
                            const dataRows = [];
                            for (let r = headerIdx + 1; r < rows.length; r++) {
                                const row = rows[r];
                                // Skip empty, divider, or too short rows
                                if (!row || row.length <= 1) continue;
                                // Ignore row if it's a copy of headers
                                if (row.some(c => c && String(c).includes('Recipient GSTIN'))) continue;
                                dataRows.push(row);
                            }

                            if (dataRows.length > 0) {
                                if (!sectionData[sectionName]) {
                                    sectionData[sectionName] = {
                                        headers: headers,
                                        rows: [],
                                        title: rawTitle
                                    };
                                }
                                sectionData[sectionName].rows.push(...dataRows);
                            }
                        }
                    } catch (readErr) {
                        logCallback(`⚠️ Warning: Failed to read file ${filePath}: ${readErr.message}`);
                    }
                }
            }

            const sections = Object.keys(sectionData);
            if (sections.length > 0) {
                // Generate Consolidated workbook
                const consolidatedWb = xlsx.utils.book_new();

                for (const sectionName of sections) {
                    const data = sectionData[sectionName];
                    
                    // Format output with top metadata
                    const formattedRows = [
                        [data.title || `Consolidated Report - ${sectionName}`],
                        [`Client Username: ${username}`, `State: ${state}`],
                        [],
                        [sectionName],
                        data.headers,
                        [],
                        ...data.rows
                    ];

                    const consolidatedWs = xlsx.utils.aoa_to_sheet(formattedRows);
                    xlsx.utils.book_append_sheet(consolidatedWb, consolidatedWs, sectionName.substring(0, 31));
                }

                // Ensure Consolidated folder exists inside the State folder
                const consolidatedFolder = path.join(statePath, 'Consolidated');
                if (!fs.existsSync(consolidatedFolder)) {
                    fs.mkdirSync(consolidatedFolder, { recursive: true });
                }

                const outputFilePath = path.join(consolidatedFolder, 'Consolidated_Rejected_Notes.xlsx');
                try {
                    xlsx.writeFile(consolidatedWb, outputFilePath);
                    logCallback(`✅ Consolidated report written for ${client} (${state}) -> ${outputFilePath}`);
                    totalConsolidated++;
                } catch (writeErr) {
                    logCallback(`❌ Error: Failed to write consolidated file for ${client} (${state}): ${writeErr.message}`);
                }
            }
        }
    }

    logCallback(`\n🎉 Consolidation complete! Processed ${totalConsolidated} consolidated files.`);
}

module.exports = { consolidateReports };
