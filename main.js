const { app, BrowserWindow, dialog, ipcMain, protocol } = require('electron');
const path = require('path');
const fs = require('fs').promises;

// -----------------------------------------------------------------------------
// Custom "app://" protocol.
// web-ifc loads its WebAssembly via fetch()/instantiateStreaming, which does NOT
// work from file:// URLs ("readAsync does not work with file:// URLs"). Serving
// the app over a privileged, fetch-capable scheme fixes that while staying 100%
// local (no HTTP server, no network).
// -----------------------------------------------------------------------------
protocol.registerSchemesAsPrivileged([
  {
    scheme: 'app',
    privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true },
  },
]);

const MIME = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.mjs': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.wasm': 'application/wasm', // required so instantiateStreaming accepts it
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
};

async function serveAppFile(request) {
  try {
    const { pathname } = new URL(request.url);
    let rel = decodeURIComponent(pathname).replace(/^\/+/, '') || 'index.html';
    const filePath = path.normalize(path.join(__dirname, rel));
    // Keep reads inside the app folder.
    if (path.relative(__dirname, filePath).startsWith('..')) {
      return new Response('Forbidden', { status: 403 });
    }
    const data = await fs.readFile(filePath);
    const ext = path.extname(filePath).toLowerCase();
    return new Response(data, {
      headers: { 'content-type': MIME[ext] || 'application/octet-stream' },
    });
  } catch (e) {
    console.error('[app:// 404]', request.url, '->', e.code || e.message);
    return new Response('Not found', { status: 404 });
  }
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    show: false,
    webPreferences: {
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
      sandbox: false,
    },
  });

  win.loadURL('app://bundle/index.html');
  win.once('ready-to-show', () => win.show());

  // Forward renderer console + crashes to the terminal — handy for this POC.
  win.webContents.on('console-message', (e) => {
    console.log(`[renderer] ${e.message}`);
  });
  win.webContents.on('render-process-gone', (_e, details) => {
    console.error('[renderer] process gone:', details.reason);
  });
}

app.whenReady().then(() => {
  protocol.handle('app', serveAppFile);
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// --- file IPC (all disk reads happen here, in the main process) --------------

ipcMain.handle('dialog:openIfcFile', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    title: 'Open IFC file',
    properties: ['openFile'],
    filters: [{ name: 'IFC Files', extensions: ['ifc'] }],
  });
  if (canceled || !filePaths?.length) return null;
  const pathName = filePaths[0];
  const data = await fs.readFile(pathName);
  return { path: pathName, data: Uint8Array.from(data).buffer };
});

// XKT is the fast path for big models: converted offline by convert.js, then
// loaded here by XKTLoaderPlugin. Same shape as the IFC handler — just a buffer.
ipcMain.handle('dialog:openXktFile', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    title: 'Open XKT file',
    properties: ['openFile'],
    filters: [{ name: 'XKT Files', extensions: ['xkt'] }],
  });
  if (canceled || !filePaths?.length) return null;
  const pathName = filePaths[0];
  const data = await fs.readFile(pathName);
  return { path: pathName, data: Uint8Array.from(data).buffer };
});

ipcMain.handle('readSampleIfc', async () => {
  try {
    const samplePath = path.join(__dirname, 'sample', 'Duplex.ifc');
    const data = await fs.readFile(samplePath);
    return { path: samplePath, data: Uint8Array.from(data).buffer };
  } catch {
    return null;
  }
});
