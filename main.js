const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('path');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1040,
    height: 760,
    minWidth: 900,
    minHeight: 640,
    backgroundColor: '#f5f1e8',
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, 'index.html'));
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

ipcMain.handle('select-xlsx-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '请选择单词Excel文件',
    properties: ['openFile'],
    filters: [{ name: 'Excel xlsx', extensions: ['xlsx'] }]
  });

  if (result.canceled || !result.filePaths.length) {
    return null;
  }

  return result.filePaths[0];
});

ipcMain.handle('exit-app', () => {
  app.quit();
});