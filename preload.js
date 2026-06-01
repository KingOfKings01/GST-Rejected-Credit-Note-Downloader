const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    selectDirectory: () => ipcRenderer.invoke('select-directory'),
    selectExcel: () => ipcRenderer.invoke('select-excel'),
    parseExcel: (filePath) => ipcRenderer.invoke('parse-excel', filePath),
    startAutomation: (config) => ipcRenderer.send('start-automation', config),
    stopAutomation: () => ipcRenderer.send('stop-automation'),
    onAutomationLog: (callback) => {
        const subscription = (event, value) => callback(value);
        ipcRenderer.on('automation-log', subscription);
        return () => ipcRenderer.removeListener('automation-log', subscription);
    },
    onAutomationFinished: (callback) => {
        const subscription = (event, value) => callback(value);
        ipcRenderer.on('automation-finished', subscription);
        return () => ipcRenderer.removeListener('automation-finished', subscription);
    }
});
