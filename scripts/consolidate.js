const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { applyDateFormattingToSheet } = require('./utils');

/**
 * Traverses the source directory to find all client and state subdirectories,
 * and consolidates all rejected notes files into BOTH State-wise and Client-wise Excel reports.
 * @param {string} sourceDir - The main download/output directory
 * @param {function} logCallback - Function to log progress messages back to the UI
 */
async function consolidateReports(sourceDir, logCallback = console.log) {
    if (!sourceDir || !fs.existsSync(sourceDir)) {
        throw new Error(`Source directory "${sourceDir}" does not exist.`);
    }

    logCallback(`Starting consolidation in: ${sourceDir}`);

    const clients = fs.readdirSync(sourceDir).filter(name => {
        const fullPath = path.join(sourceDir, name);
        return fs.statSync(fullPath).isDirectory() && !name.startsWith('.') && name !== 'Consolidated' && name !== 'Client Wise Consolidation';
    });

    if (clients.length === 0) {
        logCallback('No client folders found to consolidate.');
        return;
    }

    let totalClientConsolidated = 0;
    let totalStateConsolidated = 0;

    for (const client of clients) {
        const clientPath = path.join(sourceDir, client);
        const states = fs.readdirSync(clientPath).filter(name => {
            const fullPath = path.join(clientPath, name);
            return fs.statSync(fullPath).isDirectory() && !name.startsWith('.') && name !== 'Consolidated' && name !== 'Client Wise Consolidation';
        });

        const clientSectionData = {}; // For client-wise output
        let username = client; // Default username fallback

        for (const state of states) {
            const stateSectionData = {}; // For state-wise output
            const statePath = path.join(clientPath, state);
            const months = fs.readdirSync(statePath).filter(name => {
                const fullPath = path.join(statePath, name);
                return fs.statSync(fullPath).isDirectory() && name !== 'Consolidated' && !name.startsWith('.');
            });

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

                        let headerIdx = -1;
                        for (let r = 0; r < Math.min(10, rows.length); r++) {
                            const row = rows[r];
                            if (row &&
                                row.filter(c => c !== null && c !== undefined && String(c).trim() !== '').length > 3 &&
                                row.some(cell => cell && (
                                    String(cell).toLowerCase().includes('status') ||
                                    String(cell).toLowerCase().includes('gstin') ||
                                    String(cell).toLowerCase().includes('remarks')
                                ))
                            ) {
                                headerIdx = r;
                                break;
                            }
                        }

                        if (headerIdx !== -1 && rows.length > headerIdx) {
                            const rawTitle = rows[0] && rows[0][0] ? String(rows[0][0]) : '';
                            const sectionName = (headerIdx > 0 && rows[headerIdx - 1] && rows[headerIdx - 1][0])
                                ? String(rows[headerIdx - 1][0]).trim()
                                : path.basename(file, path.extname(file));
                            const headers = rows[headerIdx] || [];

                            if (rawTitle && rawTitle.includes('GSTIN')) {
                                const gstinMatch = rawTitle.match(/GSTIN\s*:\s*([A-Z0-9]+)/i);
                                if (gstinMatch) username = gstinMatch[1];
                            }

                            const dataRowsStateWise = [];
                            const dataRowsClientWise = [];

                            for (let r = headerIdx + 1; r < rows.length; r++) {
                                const row = rows[r];
                                if (!row || row.length <= 1) continue;
                                if (row.some(c => c && String(c).includes('Recipient GSTIN'))) continue;

                                dataRowsStateWise.push([...row]);
                                dataRowsClientWise.push([state, ...row]); // Prepend State for client-wise
                            }

                            if (dataRowsStateWise.length > 0) {
                                // Add to state-wise data
                                if (!stateSectionData[sectionName]) {
                                    stateSectionData[sectionName] = { headers: [...headers], rows: [], title: rawTitle };
                                }
                                stateSectionData[sectionName].rows.push(...dataRowsStateWise);

                                // Add to client-wise data
                                if (!clientSectionData[sectionName]) {
                                    clientSectionData[sectionName] = { headers: ['State', ...headers], rows: [], title: rawTitle };
                                }
                                clientSectionData[sectionName].rows.push(...dataRowsClientWise);
                            }
                        }
                    } catch (readErr) {
                        logCallback(`⚠️ Warning: Failed to read file ${filePath}: ${readErr.message}`);
                    }
                }
            }

            // ──── Write ONE consolidated file per STATE ────
            const stateSections = Object.keys(stateSectionData);
            if (stateSections.length > 0) {
                const stateWb = xlsx.utils.book_new();
                for (const sectionName of stateSections) {
                    const data = stateSectionData[sectionName];
                    const formattedRows = [
                        [data.title || `Consolidated Report - ${sectionName}`],
                        [`Client: ${client} (${username}) | State: ${state}`],
                        [],
                        [sectionName],
                        data.headers,
                        [],
                        ...data.rows
                    ];
                    const ws = xlsx.utils.aoa_to_sheet(formattedRows);
                    applyDateFormattingToSheet(ws); // Fix dates in sheet!
                    const cleanSheetName = sectionName.replace(/[:\\/?*\[\]]/g, ' ').substring(0, 31).trim();
                    xlsx.utils.book_append_sheet(stateWb, ws, cleanSheetName);
                }

                const stateConsolidatedFolder = path.join(statePath, 'Consolidated');
                if (!fs.existsSync(stateConsolidatedFolder)) fs.mkdirSync(stateConsolidatedFolder, { recursive: true });
                const stateOutputFilePath = path.join(stateConsolidatedFolder, 'Consolidated_Rejected_Notes.xlsx');
                try {
                    xlsx.writeFile(stateWb, stateOutputFilePath, { cellDates: true });
                    totalStateConsolidated++;
                } catch (err) {
                    logCallback(`❌ Error writing state file for ${client}/${state}: ${err.message}`);
                }
            }
        } // end of states loop

        // ──── Write ONE consolidated file per CLIENT ────
        const clientSections = Object.keys(clientSectionData);
        if (clientSections.length > 0) {
            const clientWb = xlsx.utils.book_new();
            for (const sectionName of clientSections) {
                const data = clientSectionData[sectionName];
                const formattedRows = [
                    [data.title || `Consolidated Report - ${sectionName}`],
                    [`Client: ${client} (${username})`],
                    [],
                    [sectionName],
                    data.headers,
                    [],
                    ...data.rows
                ];
                const ws = xlsx.utils.aoa_to_sheet(formattedRows);
                applyDateFormattingToSheet(ws); // Fix dates in sheet!
                const cleanSheetName = sectionName.replace(/[:\\/?*\[\]]/g, ' ').substring(0, 31).trim();
                xlsx.utils.book_append_sheet(clientWb, ws, cleanSheetName);
            }

            const clientConsolidatedFolder = path.join(clientPath, 'Client Wise Consolidation');
            if (!fs.existsSync(clientConsolidatedFolder)) fs.mkdirSync(clientConsolidatedFolder, { recursive: true });
            const clientOutputFilePath = path.join(clientConsolidatedFolder, `${client}_Client_Wise_Consolidated.xlsx`);
            try {
                xlsx.writeFile(clientWb, clientOutputFilePath, { cellDates: true });
                logCallback(`✅ Consolidated (Client-wise & State-wise) created for client "${client}"`);
                totalClientConsolidated++;
            } catch (err) {
                logCallback(`❌ Error writing client file for ${client}: ${err.message}`);
            }
        }
    }

    logCallback(`\n🎉 Consolidation complete! Created ${totalClientConsolidated} client-wise and ${totalStateConsolidated} state-wise files.`);
}

module.exports = { consolidateReports };
