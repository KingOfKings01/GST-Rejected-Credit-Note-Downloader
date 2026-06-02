// Developer Configuration
const BYPASS_AUTH = true; // Set to true to disable authentication overlay during development

// State management
let loadedClients = [];
let isRunning = false;
let currentExcelFilePath = '';

// DOM Elements
const selectFy = document.getElementById('select-fy');
const selectPeriodFrom = document.getElementById('select-period-from');
const selectPeriodTo = document.getElementById('select-period-to');
const selectType = document.getElementById('select-type');
const btnBrowseOutput = document.getElementById('btn-browse-output');
const outputPathDisplay = document.getElementById('output-path-display');
const excelPathDisplay = document.getElementById('excel-path-display');
const btnBrowseExcel = document.getElementById('btn-browse-excel');
const btnRunAutomation = document.getElementById('btn-run-automation');
const btnStopAutomation = document.getElementById('btn-stop-automation');
const btnClearConsole = document.getElementById('btn-clear-console');
const consoleOutput = document.getElementById('console-output');
const clientSearchInput = document.getElementById('client-search');
const clientListContainer = document.getElementById('client-list');
const chkSelectAll = document.getElementById('chk-select-all');
const chkSkipExisting = document.getElementById('chk-skip-existing');

// Stats DOM Elements
const statTotal = document.getElementById('stat-total');
const statSelected = document.getElementById('stat-selected');
const statCompleted = document.getElementById('stat-completed');
const statFailed = document.getElementById('stat-failed');

// Footer DOM Elements
const progressText = document.getElementById('progress-text');
const progressBar = document.getElementById('progress-bar');
const progressPercentage = document.getElementById('progress-percentage');

// 1. Dynamic Financial Years & Months Options
function initializeDateOptions() {
    const now = new Date();
    const currentYear = now.getFullYear();
    // Indian Financial Year starts in April (Month index 3)
    const currentFYStartYear = now.getMonth() >= 3 ? currentYear : currentYear - 1;
    
    // Generate Current FY to next year, and past 3 years (Total 4 options)
    const fyOptions = [];
    for (let i = 0; i < 4; i++) {
        const startYr = currentFYStartYear - i;
        const endYrAbbr = ((startYr + 1) % 100).toString().padStart(2, '0');
        fyOptions.push(`${startYr}-${endYrAbbr}`);
    }

    // Populate FY dropdown
    selectFy.innerHTML = '';
    fyOptions.forEach(fy => {
        const option = document.createElement('option');
        option.value = fy;
        option.textContent = fy;
        selectFy.appendChild(option);
    });

    // Populate months when FY changes or initially
    selectFy.addEventListener('change', populateReturnPeriods);
    
    selectPeriodFrom.addEventListener('change', () => {
        if (selectPeriodFrom.selectedIndex > selectPeriodTo.selectedIndex) {
            selectPeriodTo.selectedIndex = selectPeriodFrom.selectedIndex;
        }
    });
    selectPeriodTo.addEventListener('change', () => {
        if (selectPeriodFrom.selectedIndex > selectPeriodTo.selectedIndex) {
            selectPeriodFrom.selectedIndex = selectPeriodTo.selectedIndex;
        }
    });

    populateReturnPeriods();
}

function populateReturnPeriods() {
    const selectedFy = selectFy.value;
    if (!selectedFy) return;

    const parts = selectedFy.split('-');
    const startYear = parseInt(parts[0], 10);
    const endYear = 2000 + parseInt(parts[1], 10);

    // List of months in Indian FY order (April to March)
    const baseMonths = [
        { name: "April", monthIndex: 3, year: startYear },
        { name: "May", monthIndex: 4, year: startYear },
        { name: "June", monthIndex: 5, year: startYear },
        { name: "July", monthIndex: 6, year: startYear },
        { name: "August", monthIndex: 7, year: startYear },
        { name: "September", monthIndex: 8, year: startYear },
        { name: "October", monthIndex: 9, year: startYear },
        { name: "November", monthIndex: 10, year: startYear },
        { name: "December", monthIndex: 11, year: startYear },
        { name: "January", monthIndex: 0, year: endYear },
        { name: "February", monthIndex: 1, year: endYear },
        { name: "March", monthIndex: 2, year: endYear }
    ];

    const now = new Date();
    
    // Filter months: if a month is in the future relative to current time, exclude it
    const filteredMonths = baseMonths.filter(month => {
        // Compare year and month index
        if (month.year > now.getFullYear()) {
            return false;
        }
        if (month.year === now.getFullYear() && month.monthIndex > now.getMonth()) {
            return false;
        }
        return true;
    });

    // Populate Period Dropdowns
    selectPeriodFrom.innerHTML = '';
    selectPeriodTo.innerHTML = '';
    
    // If all months are filtered out (shouldn't happen for past FY, only possible if system time is messed up)
    const monthsToShow = filteredMonths.length > 0 ? filteredMonths : baseMonths;

    monthsToShow.forEach(m => {
        const optionFrom = document.createElement('option');
        optionFrom.value = m.name;
        optionFrom.textContent = m.name;
        selectPeriodFrom.appendChild(optionFrom);

        const optionTo = document.createElement('option');
        optionTo.value = m.name;
        optionTo.textContent = m.name;
        selectPeriodTo.appendChild(optionTo);
    });

    // Select latest available month by default for "To Period", and 3 months ago (or index length - 3) for "From Period"
    if (selectPeriodTo.options.length > 0) {
        selectPeriodTo.selectedIndex = selectPeriodTo.options.length - 1;
        selectPeriodFrom.selectedIndex = Math.max(0, selectPeriodTo.options.length - 3);
    }
}

// 2. Client excel loading
async function loadClientsExcel(filePath) {
    if (!filePath) return;
    
    appendConsoleLine(`System: Parsing Excel workbook at ${filePath}...`, 'info');
    
    const result = await window.electronAPI.parseExcel(filePath);
    
    if (result.success) {
        excelPathDisplay.textContent = result.filePath;
        excelPathDisplay.title = result.filePath;
        currentExcelFilePath = result.filePath;
        
        const skipChecked = chkSkipExisting.checked;
        loadedClients = result.clients.map(c => {
            return {
                ...c,
                selected: skipChecked && c.status === 'success' ? false : true
            };
        });
        
        appendConsoleLine(`System: Successfully loaded ${loadedClients.length} clients from Excel file.`, 'success');
        renderClientsList();
        updateStats();
    } else {
        appendConsoleLine(`System Error: Failed to parse Excel file. ${result.error}`, 'error');
        alert(`System Error: Failed to parse Excel file.\n\n${result.error}\n\nPlease close the file in Excel and try again.`);
        excelPathDisplay.textContent = 'Error loading file';
        loadedClients = [];
        currentExcelFilePath = '';
        renderClientsList();
        updateStats();
    }
}

// 3. Render Clients Table
function renderClientsList() {
    const query = clientSearchInput.value.toLowerCase().trim();
    clientListContainer.innerHTML = '';

    const filtered = loadedClients.filter(c => {
        return (c.clientName && c.clientName.toLowerCase().includes(query)) ||
               (c.stateName && c.stateName.toLowerCase().includes(query)) ||
               (c.clientState && c.clientState.toLowerCase().includes(query)) ||
               (c.gstNo && c.gstNo.toLowerCase().includes(query)) ||
               (c.username && c.username.toLowerCase().includes(query));
    });

    if (filtered.length === 0) {
        if (loadedClients.length === 0) {
            clientListContainer.innerHTML = '<div style="padding: 1rem; color: var(--text-muted); text-align: center; font-size: 0.85rem;">Please browse and select a client login Excel file.</div>';
        } else {
            clientListContainer.innerHTML = '<div style="padding: 1rem; color: var(--text-muted); text-align: center; font-size: 0.85rem;">No clients match search filters.</div>';
        }
        return;
    }

    filtered.forEach((client, idx) => {
        const item = document.createElement('div');
        item.className = 'client-item';
        
        // Checkbox cell
        const chkCell = document.createElement('div');
        chkCell.className = 'chk-cell';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = client.selected !== false;
        checkbox.addEventListener('change', () => {
            client.selected = checkbox.checked;
            updateStats();
        });
        chkCell.appendChild(checkbox);
        
        // Serial Number cell (read-only)
        const snoCell = document.createElement('div');
        snoCell.className = 'sno-cell';
        snoCell.textContent = client.srNo || (idx + 1);

        // Name cell
        const nameCell = document.createElement('div');
        nameCell.className = 'client-name-cell';
        nameCell.textContent = client.clientName || client.clientState || 'Unknown Client';
        nameCell.title = nameCell.textContent;

        // State cell
        const stateCell = document.createElement('div');
        stateCell.className = 'client-state-cell';
        stateCell.textContent = client.stateName || 'N/A';
        stateCell.title = stateCell.textContent;

        // GSTIN cell
        const gstCell = document.createElement('div');
        gstCell.className = 'client-gst-cell';
        gstCell.textContent = client.gstNo || 'N/A';

        // Status cell
        const statusCell = document.createElement('div');
        statusCell.className = 'client-status-cell';
        
        const badge = document.createElement('span');
        badge.className = `badge badge-${client.status}`;
        badge.id = `badge-${client.username}`;
        
        if (client.status === 'pending') badge.textContent = 'Pending';
        else if (client.status === 'running') badge.textContent = 'Running';
        else if (client.status === 'success') badge.textContent = 'Success';
        else if (client.status === 'failed') badge.textContent = 'Failed';
        else if (client.status === 'zip_pending') {
            badge.className = 'badge badge-warning';
            badge.textContent = 'Zip Pending';
        }
        else if (client.status === 'zip_ready') {
            badge.className = 'badge badge-success';
            badge.textContent = 'zips is ready to download';
        }
        
        statusCell.appendChild(badge);

        item.appendChild(chkCell);
        item.appendChild(snoCell);
        item.appendChild(nameCell);
        item.appendChild(stateCell);
        item.appendChild(gstCell);
        item.appendChild(statusCell);

        clientListContainer.appendChild(item);
    });
}

// 4. Update Stats counters
function updateStats() {
    statTotal.textContent = loadedClients.length;
    const selected = loadedClients.filter(c => c.selected !== false).length;
    statSelected.textContent = selected;

    const completed = loadedClients.filter(c => c.status === 'success').length;
    statCompleted.textContent = completed;

    const failed = loadedClients.filter(c => c.status === 'failed').length;
    statFailed.textContent = failed;
}

// 5. Console Utility functions
function appendConsoleLine(text, type = '') {
    const line = document.createElement('div');
    line.className = `console-line ${type}`;
    line.textContent = text;
    consoleOutput.appendChild(line);
    consoleOutput.scrollTop = consoleOutput.scrollHeight;
}

function clearConsole() {
    consoleOutput.innerHTML = '';
}

// 6. Handle Run Operations
function setRunningState(running) {
    isRunning = running;
    
    // Toggle Inputs disabled state
    document.querySelectorAll('select').forEach(el => el.disabled = running);
    clientSearchInput.disabled = running;
    btnBrowseExcel.disabled = running;
    btnBrowseOutput.disabled = running;
    chkSelectAll.disabled = running;
    chkSkipExisting.disabled = running;
    const chkOnlyReadyZips = document.getElementById('chk-only-ready-zips');
    if (chkOnlyReadyZips) chkOnlyReadyZips.disabled = running;

    btnRunAutomation.disabled = running;
    btnStopAutomation.disabled = !running;
}

async function startAutomation() {
    let clientsToRun = loadedClients.filter(c => c.selected !== false);
    
    const onlyReadyZips = document.getElementById('chk-only-ready-zips').checked;
    if (onlyReadyZips) {
        clientsToRun = clientsToRun.filter(c => c.status === 'zip_ready');
    }

    if (clientsToRun.length === 0) {
        alert(onlyReadyZips ? 'No clients with ready zips to download.' : 'Please select at least one client to run.');
        return;
    }

    const downloadFolder = outputPathDisplay.textContent.trim();
    if (!downloadFolder) {
        alert('Please select an output download directory.');
        return;
    }

    // Reset statuses of all selected clients to pending
    clientsToRun.forEach(c => {
        c.status = 'pending';
    });
    renderClientsList();
    updateStats();

    setRunningState(true);
    clearConsole();
    appendConsoleLine('System: Starting GST Rejected Credit Note Automation Process...', 'info');

    // Build configuration payload
    const config = {
        clients: clientsToRun,
        selections: {
            financialYear: selectFy.value,
            returnPeriod: selectPeriodTo.value,
            returnPeriodFrom: selectPeriodFrom.value,
            returnPeriodTo: selectPeriodTo.value,
            returnType: selectType.value,
            browserZoom: document.getElementById('select-zoom').value
        },
        downloadFolder,
        excelFilePath: currentExcelFilePath
    };

    progressText.textContent = 'Initializing automation...';
    progressBar.style.width = '0%';
    progressPercentage.textContent = '0%';

    // Send trigger to Main Process
    window.electronAPI.startAutomation(config);
}

function stopAutomation() {
    appendConsoleLine('\nSystem: Sending cancellation request to running automation process...', 'error');
    window.electronAPI.stopAutomation();
}

// 7. Parse Progress Tags from Console Output
function parseConsoleProgress(logText) {
    // Format: CLIENT_PROGRESS:START:current:total:username:clientName
    const startMatch = logText.match(/CLIENT_PROGRESS:START:(\d+):(\d+):([^:]+):(.+)/);
    if (startMatch) {
        const current = parseInt(startMatch[1], 10);
        const total = parseInt(startMatch[2], 10);
        const username = startMatch[3].trim();
        const clientName = startMatch[4].trim();
        
        progressText.textContent = `Processing (${current}/${total}): ${clientName}`;
        const pct = Math.round(((current - 1) / total) * 100);
        progressBar.style.width = `${pct}%`;
        progressPercentage.textContent = `${pct}%`;

        // Update target client status inside memory & DOM
        const targetClient = loadedClients.find(c => c.username === username);
        if (targetClient) {
            targetClient.status = 'running';
            const badge = document.getElementById(`badge-${targetClient.username}`);
            if (badge) {
                badge.className = 'badge badge-running';
                badge.textContent = 'Running';
            }
            updateStats();
        }
    }

    // Zip Pending Match: CLIENT_PROGRESS:ZIP_PENDING:username:month:readyAt:pendingSection
    const zipPendingMatch = logText.match(/CLIENT_PROGRESS:ZIP_PENDING:([^:]+):([^:]+):([^:]+):(.+)/);
    if (zipPendingMatch) {
        const username = zipPendingMatch[1].trim();
        const targetClient = loadedClients.find(c => c.username === username);
        if (targetClient) {
            targetClient.status = 'zip_pending';
            const month = zipPendingMatch[2].trim();
            const readyAt = zipPendingMatch[3].trim();
            const pendingSection = zipPendingMatch[4].trim();
            targetClient.excelStatus = `Zip Pending | ${month} | ${readyAt} | ${pendingSection}`;
            const badge = document.getElementById(`badge-${username}`);
            if (badge) {
                badge.className = 'badge badge-warning';
                badge.textContent = 'Zip Pending';
            }
            updateStats();
        }
        return;
    }

    // Success Match: CLIENT_PROGRESS:SUCCESS:username
    const successMatch = logText.match(/CLIENT_PROGRESS:SUCCESS:(.+)/);
    if (successMatch) {
        const username = successMatch[1].trim();
        const targetClient = loadedClients.find(c => c.username === username);
        if (targetClient) {
            targetClient.status = 'success';
            targetClient.excelStatus = 'Success';
            const badge = document.getElementById(`badge-${username}`);
            if (badge) {
                badge.className = 'badge badge-success';
                badge.textContent = 'Success';
            }
            updateStats();
        }
        return;
    }

    // Failure Match: CLIENT_PROGRESS:FAILED:username: reason
    const failedMatch = logText.match(/CLIENT_PROGRESS:FAILED:([^:]+)/);
    if (failedMatch) {
        const username = failedMatch[1].trim();
        const targetClient = loadedClients.find(c => c.username === username);
        if (targetClient) {
            targetClient.status = 'failed';
            targetClient.excelStatus = 'Failed';
            const badge = document.getElementById(`badge-${username}`);
            if (badge) {
                badge.className = 'badge badge-failed';
                badge.textContent = 'Failed';
            }
            updateStats();
        }
        return;
    }
}

// 8. IPC Listeners
window.electronAPI.onAutomationLog((data) => {
    // Output directly to terminal
    // Parse tags to extract logs versus execution errors
    const lines = data.split('\n');
    lines.forEach(line => {
        if (!line.trim()) return;
        
        // Hide internal process state communication tags from output
        if (line.includes('CLIENT_PROGRESS:')) {
            parseConsoleProgress(line);
            return;
        }

        if (line.startsWith('ERROR:') || line.startsWith('❌')) {
            appendConsoleLine(line, 'error');
        } else if (line.startsWith('🎉') || line.startsWith('✅')) {
            appendConsoleLine(line, 'success');
        } else if (line.includes('⚠️')) {
            appendConsoleLine(line, 'warning');
        } else {
            appendConsoleLine(line);
        }
    });
});

window.electronAPI.onAutomationFinished((code) => {
    setRunningState(false);
    if (code === 0) {
        appendConsoleLine('\nSystem: Automation completed successfully.', 'success');
        progressText.textContent = 'Completed';
        progressBar.style.width = '100%';
        progressPercentage.textContent = '100%';
    } else {
        appendConsoleLine(`\nSystem: Automation stopped or exited with code ${code}.`, 'error');
        progressText.textContent = 'Execution Stopped';
    }
    updateStats();
});

// 9. Event Listeners
btnBrowseExcel.addEventListener('click', async () => {
    const selectedFile = await window.electronAPI.selectExcel();
    if (selectedFile) {
        loadClientsExcel(selectedFile);
    }
});

btnBrowseOutput.addEventListener('click', async () => {
    const selectedDir = await window.electronAPI.selectDirectory();
    if (selectedDir) {
        outputPathDisplay.textContent = selectedDir;
    }
});

btnRunAutomation.addEventListener('click', () => {
    startAutomation();
});

btnStopAutomation.addEventListener('click', () => {
    stopAutomation();
});

btnClearConsole.addEventListener('click', () => {
    clearConsole();
});

clientSearchInput.addEventListener('input', () => {
    renderClientsList();
});

chkSelectAll.addEventListener('change', () => {
    const isChecked = chkSelectAll.checked;
    loadedClients.forEach(c => {
        c.selected = isChecked;
    });
    renderClientsList();
    updateStats();
});

chkSkipExisting.addEventListener('change', () => {
    const skipChecked = chkSkipExisting.checked;
    loadedClients.forEach(c => {
        if (c.status === 'success') {
            c.selected = !skipChecked;
        }
    });
    renderClientsList();
    updateStats();
});

// Initialization
document.addEventListener('DOMContentLoaded', () => {
    initializeDateOptions();
    renderClientsList(); // Display default "Please browse..." message
});

document.addEventListener('DOMContentLoaded', () => {
    const step1View = document.getElementById('view-step-1');
    const step2View = document.getElementById('view-step-2');
    
    const proceedButton = document.getElementById('btn-proceed-to-clients');
    const backButton = document.getElementById('btn-back-to-setup');
    const excelPathDisplay = document.getElementById('excel-path-display');

    // 1. Move from Form Configuration to Client Panel
    proceedButton.addEventListener('click', () => {
        // Validation: Verify an input sheet file was chosen before letting them proceed
        if (excelPathDisplay.innerText === 'No file selected') {
            alert('Please select a valid Client Login Excel File before moving to Step 2.');
            return;
        }

        step1View.classList.add('hidden-panel');
        step2View.classList.remove('hidden-panel');
    });

    // 2. Go Back to adjust form data settings
    backButton.addEventListener('click', () => {
        step2View.classList.add('hidden-panel');
        step1View.classList.remove('hidden-panel');
    });
});

// --- Auth Provider Integration ---
document.addEventListener('DOMContentLoaded', () => {
    const authPanel = document.getElementById('auth-panel');
    const appContent = document.getElementById('app-content');
    const authForm = document.getElementById('auth-form');
    const authEmailInput = document.getElementById('auth-email');
    const authOtpGroup = document.getElementById('auth-otp-group');
    const authOtpInput = document.getElementById('auth-otp');
    const authEmailGroup = document.getElementById('auth-email-group');
    const authErrorMsg = document.getElementById('auth-error-msg');
    const authSubmitBtn = document.getElementById('btn-auth-submit');
    const logoutBtn = document.getElementById('btn-logout');

    let currentStage = 'email'; // 'email' or 'otp'
    let generatedOtp = '';
    let targetEmail = '';

    // Always show auth panel on startup
    if (BYPASS_AUTH) {
        document.getElementById('user-display-email').textContent = 'Developer Mode';
        showAppContent();
    } else {
        showAuthPanel();
    }

    authForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        if (currentStage === 'email') {
            const email = authEmailInput.value.trim();
            if (!email) return;
            targetEmail = email;
            requestOtp(email);
        } else if (currentStage === 'otp') {
            const userOtp = authOtpInput.value.trim();
            console.log('--- OTP COMPARISON ---');
            console.log('User OTP input:', JSON.stringify(userOtp));
            console.log('Generated OTP:', JSON.stringify(generatedOtp));
            console.log('Match result:', userOtp == generatedOtp);
            console.log('----------------------');
            if (userOtp == generatedOtp || userOtp === String(generatedOtp)) {
                // Save user email in main process for auto-logout tracking
                window.electronAPI.setAuthEmail(targetEmail);
                document.getElementById('user-display-email').textContent = targetEmail;
                showAppContent();
            } else {
                showAuthError('Invalid verification code. Please try again.');
            }
        }
    });

    logoutBtn.addEventListener('click', async () => {
        if (targetEmail) {
            await window.electronAPI.authTrackLogout(targetEmail);
        }
        window.electronAPI.setAuthEmail(null);
        showAuthPanel();
    });

    async function requestOtp(email) {
        setAuthLoading(true);
        hideAuthError();

        try {
            const result = await window.electronAPI.authAuthorize(email);
            if (result.success) {
                generatedOtp = result.otp;
                // Transition to OTP stage
                currentStage = 'otp';
                authEmailGroup.classList.add('hidden-panel');
                authOtpGroup.classList.remove('hidden-panel');
                authOtpInput.required = true;
                authOtpInput.focus();
                authSubmitBtn.textContent = 'Verify OTP';
            } else {
                showAuthError(result.message || 'Access denied. Your email is not authorized or is blocked.');
                showAuthPanel();
            }
        } catch (err) {
            console.error('Authentication request failed:', err);
            showAuthError('Failed to connect to authorization service.');
            showAuthPanel();
        } finally {
            setAuthLoading(false);
        }
    }

    function setAuthLoading(isLoading) {
        authSubmitBtn.disabled = isLoading;
        authEmailInput.disabled = isLoading;
        authOtpInput.disabled = isLoading;
        if (isLoading) {
            authSubmitBtn.textContent = currentStage === 'email' ? 'Sending code...' : 'Verifying...';
        } else {
            authSubmitBtn.textContent = currentStage === 'email' ? 'Send Verification Code' : 'Verify OTP';
        }
    }

    function showAuthError(message) {
        authErrorMsg.textContent = message;
        authErrorMsg.classList.remove('hidden-panel');
    }

    function hideAuthError() {
        authErrorMsg.classList.add('hidden-panel');
        authErrorMsg.textContent = '';
    }

    function showAppContent() {
        authPanel.classList.add('hidden-panel');
        appContent.classList.remove('hidden-panel');
    }

    function showAuthPanel() {
        appContent.classList.add('hidden-panel');
        authPanel.classList.remove('hidden-panel');
        authEmailGroup.classList.remove('hidden-panel');
        authOtpGroup.classList.add('hidden-panel');
        authEmailInput.value = '';
        authOtpInput.value = '';
        authOtpInput.required = false;
        currentStage = 'email';
        generatedOtp = '';
        targetEmail = '';
        authSubmitBtn.textContent = 'Send Verification Code';
        hideAuthError();
    }

    // --- REPORT CONSOLIDATOR INTEGRATION ---
    const tabDownloader = document.getElementById('tab-downloader');
    const tabConsolidator = document.getElementById('tab-consolidator');
    const viewStep1 = document.getElementById('view-step-1');
    const viewStep2 = document.getElementById('view-step-2');
    const viewConsolidator = document.getElementById('view-consolidator');
    
    const btnBrowseConsolidateSrc = document.getElementById('btn-browse-consolidate-src');
    const consolidateSrcDisplay = document.getElementById('consolidate-src-display');
    const btnRunConsolidation = document.getElementById('btn-run-consolidation');
    const consolidateConsoleOutput = document.getElementById('consolidate-console-output');

    let consolidateSourceDir = '';
    let downloaderWasOnStep2 = false;

    // Tab Switching
    tabDownloader.addEventListener('click', () => {
        tabDownloader.classList.add('active');
        tabConsolidator.classList.remove('active');
        
        viewConsolidator.classList.add('hidden-panel');
        if (downloaderWasOnStep2) {
            viewStep2.classList.remove('hidden-panel');
        } else {
            viewStep1.classList.remove('hidden-panel');
        }
    });

    tabConsolidator.addEventListener('click', () => {
        tabConsolidator.classList.add('active');
        tabDownloader.classList.remove('active');
        
        downloaderWasOnStep2 = !viewStep2.classList.contains('hidden-panel');
        
        viewStep1.classList.add('hidden-panel');
        viewStep2.classList.add('hidden-panel');
        viewConsolidator.classList.remove('hidden-panel');
    });

    // Browse Source Folder for Consolidation
    btnBrowseConsolidateSrc.addEventListener('click', async () => {
        const dir = await window.electronAPI.selectDirectory();
        if (dir) {
            consolidateSourceDir = dir;
            consolidateSrcDisplay.textContent = dir;
            consolidateSrcDisplay.title = dir;
            btnRunConsolidation.disabled = false;
            appendConsolidateLog(`System: Source folder set to ${dir}`);
        }
    });

    // Run Consolidation Action
    btnRunConsolidation.addEventListener('click', async () => {
        if (!consolidateSourceDir) return;
        
        btnRunConsolidation.disabled = true;
        btnBrowseConsolidateSrc.disabled = true;
        consolidateConsoleOutput.innerHTML = '';
        appendConsolidateLog('System: Initializing report consolidation...', 'info');

        const result = await window.electronAPI.runConsolidation(consolidateSourceDir);
        
        btnRunConsolidation.disabled = false;
        btnBrowseConsolidateSrc.disabled = false;
        
        if (result.success) {
            appendConsolidateLog('System: Consolidation process finished successfully.', 'success');
        } else {
            appendConsolidateLog(`System Error: Consolidation failed. ${result.error}`, 'error');
        }
    });

    // Handle consolidation logs from the main process
    window.electronAPI.onConsolidationLog((message) => {
        let type = 'info';
        if (message.includes('✅')) type = 'success';
        if (message.includes('❌') || message.includes('Error')) type = 'error';
        if (message.includes('⚠️')) type = 'warning';
        appendConsolidateLog(message, type);
    });

    function appendConsolidateLog(text, type = 'info') {
        const line = document.createElement('div');
        line.className = `console-line ${type}`;
        line.textContent = text;
        consolidateConsoleOutput.appendChild(line);
        consolidateConsoleOutput.scrollTop = consolidateConsoleOutput.scrollHeight;
    }

    function showNotification(client) {
        if (!("Notification" in window)) return;
        const title = 'GST Zip File Ready';
        const options = {
            body: `Zip file is ready to download for client: ${client.clientName || client.username}`,
            requireInteraction: true
        };
        if (Notification.permission === 'granted') {
            new Notification(title, options);
        } else if (Notification.permission !== 'denied') {
            Notification.requestPermission().then(permission => {
                if (permission === 'granted') {
                    new Notification(title, options);
                }
            });
        }
    }

    // Request notification permission on startup
    if (typeof Notification !== 'undefined' && Notification.permission !== 'granted' && Notification.permission !== 'denied') {
        Notification.requestPermission();
    }

    // Start global countdown timer for zip pending statuses
    setInterval(() => {
        loadedClients.forEach(c => {
            // If the client status is currently active, success, failed or pending, skip timer overwrite
            if (c.status === 'running' || c.status === 'pending' || c.status === 'success' || c.status === 'failed') {
                return;
            }
            
            if (c.excelStatus && c.excelStatus.startsWith('Zip Pending')) {
                const parts = c.excelStatus.split('|').map(s => s.trim());
                if (parts.length >= 4) {
                    const readyAt = parseInt(parts[2], 10);
                    const now = Date.now();
                    
                    const badge = document.getElementById(`badge-${c.username}`);
                    if (now < readyAt) {
                        c.status = 'zip_pending';
                        if (badge) {
                            const diffSeconds = Math.max(0, Math.floor((readyAt - now) / 1000));
                            const mins = Math.floor(diffSeconds / 60);
                            const secs = diffSeconds % 60;
                            badge.className = 'badge badge-warning';
                            badge.textContent = `Zip Pending (${mins}:${secs.toString().padStart(2, '0')})`;
                        }
                    } else {
                        if (c.status === 'zip_pending') {
                            showNotification(c);
                        }
                        c.status = 'zip_ready';
                        if (badge) {
                            badge.className = 'badge badge-success';
                            badge.textContent = 'zips is ready to download';
                        }
                    }
                }
            }
        });
    }, 1000);
});