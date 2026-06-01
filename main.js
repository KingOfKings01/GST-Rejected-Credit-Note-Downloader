const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const xlsx = require('xlsx');

let mainWindow;
let activeChildProcess = null;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1000,
        height: 750,
        minWidth: 800,
        minHeight: 600,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false
        },
        autoHideMenuBar: true
    });

    mainWindow.loadFile(path.join(__dirname, 'src', 'index.html'));

    // Open devtools in development environment if needed
    // mainWindow.webContents.openDevTools();

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

// IPC Handler for Directory Selection
ipcMain.handle('select-directory', async () => {
    if (!mainWindow) return null;
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openDirectory']
    });
    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }
    return result.filePaths[0];
});

// IPC Handler for Excel File Selection
ipcMain.handle('select-excel', async () => {
    if (!mainWindow) return null;
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile'],
        filters: [
            { name: 'Excel Files', extensions: ['xlsx', 'xlsm'] }
        ]
    });
    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }
    return result.filePaths[0];
});

// IPC Handler for Excel File Parsing
ipcMain.handle('parse-excel', async (event, customFilePath) => {
    try {
        let filePath = customFilePath;
        if (!filePath) {
            // Default path relative to project workspace
            filePath = path.join(__dirname, 'inputs', 'GST Auto Login all client.xlsm');
        }

        if (!fs.existsSync(filePath)) {
            throw new Error(`File not found at: ${filePath}`);
        }

        // Test write access (check if file is locked/open in Excel)
        try {
            const fd = fs.openSync(filePath, 'r+');
            fs.closeSync(fd);
        } catch (lockErr) {
            if (lockErr.code === 'EBUSY' || lockErr.code === 'EACCES') {
                throw new Error(`The Excel file '${path.basename(filePath)}' is currently open or locked by Excel. Please close it!`);
            }
        }

        const workbook = xlsx.readFile(filePath);
        const sheetName = 'customer data and filing return';
        
        if (!workbook.SheetNames.includes(sheetName)) {
            throw new Error(`Sheet "${sheetName}" not found in Excel file. Available sheets: ${workbook.SheetNames.join(', ')}`);
        }

        const sheet = workbook.Sheets[sheetName];
        const rawData = xlsx.utils.sheet_to_json(sheet);
        
        // Map raw row fields to normalized client objects
        const clients = rawData.map(row => {
            const excelStatus = (row['STATUS'] || '').toString().trim();
            let status = 'pending';
            if (excelStatus.toLowerCase() === 'success') {
                status = 'success';
            } else if (excelStatus.toLowerCase().startsWith('failed') || excelStatus.length > 0) {
                status = 'failed';
            }
            return {
                srNo: row['SR.NO'] || '',
                clientState: row[' CLIENT + STATE'] || '',
                gstNo: row['GST .NO'] || '',
                username: row['USER ID'] || '',
                password: row['PASSWORD'] || '',
                clientName: row['CLIENT NAME'] || '',
                stateName: row['STATE NAME'] || '',
                status: status,
                excelStatus: excelStatus
            };
        }).filter(c => c.username && c.password);

        return { success: true, clients, filePath };
    } catch (error) {
        console.error('Error parsing excel:', error);
        return { success: false, error: error.message };
    }
});

// IPC Listener to Start Automation
ipcMain.on('start-automation', (event, config) => {
    if (activeChildProcess) {
        event.reply('automation-log', '⚠️ Automation is already running!\n');
        return;
    }

    try {
        const { excelFilePath } = config;
        if (excelFilePath && fs.existsSync(excelFilePath)) {
            // Verify write access to excel file before launching script
            try {
                const fd = fs.openSync(excelFilePath, 'r+');
                fs.closeSync(fd);
            } catch (lockErr) {
                if (lockErr.code === 'EBUSY' || lockErr.code === 'EACCES') {
                    throw new Error(`Cannot start: The Excel file '${path.basename(excelFilePath)}' is currently open or locked by Excel. Please close it first!`);
                }
            }
        }

        // Write the config to a temp JSON file in the OS temp directory
        const tempConfigPath = path.join(app.getPath('temp'), 'temp_run_config.json');
        fs.writeFileSync(tempConfigPath, JSON.stringify(config, null, 2), 'utf8');

        // Spawn script.js process
        let scriptPath = path.join(__dirname, 'scripts', 'script.js');
        if (scriptPath.includes('app.asar') && !scriptPath.includes('app.asar.unpacked')) {
            scriptPath = scriptPath.replace('app.asar', 'app.asar.unpacked');
        }
        
        console.log(`Spawning child process for script: ${scriptPath}`);
        activeChildProcess = spawn('node', [scriptPath, tempConfigPath], {
            cwd: path.dirname(scriptPath)
        });

        // Redirect stdout
        activeChildProcess.stdout.on('data', (data) => {
            if (mainWindow) {
                mainWindow.webContents.send('automation-log', data.toString());
            }
        });

        // Redirect stderr
        activeChildProcess.stderr.on('data', (data) => {
            if (mainWindow) {
                mainWindow.webContents.send('automation-log', `ERROR: ${data.toString()}`);
            }
        });

        // Handle process completion
        activeChildProcess.on('close', (code) => {
            activeChildProcess = null;
            if (mainWindow) {
                mainWindow.webContents.send('automation-finished', code);
            }
            // Clean up the temp config file
            try {
                if (fs.existsSync(tempConfigPath)) {
                    fs.unlinkSync(tempConfigPath);
                }
            } catch (e) {
                console.error('Failed to delete temp config:', e);
            }
        });

    } catch (err) {
        console.error('Failed to start automation:', err);
        if (mainWindow) {
            mainWindow.webContents.send('automation-log', `CRITICAL ERROR: ${err.message}\n`);
            mainWindow.webContents.send('automation-finished', 1);
        }
    }
});

// IPC Listener to Stop Automation
ipcMain.on('stop-automation', () => {
    if (activeChildProcess) {
        console.log('Killing active automation child process...');
        activeChildProcess.kill('SIGINT');
        activeChildProcess = null;
    }
});