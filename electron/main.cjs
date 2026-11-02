'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { app, BrowserWindow, net, protocol, shell } = require('electron');

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'app',
    privileges: {
      standard: true,
      secure: true,
      supportFetchAPI: true,
      corsEnabled: true,
    },
  },
]);

app.setName('ONAT Disciplina de Pago');

// Keep Electron preferences/cache beside the portable executable on a USB drive.
// Imported spreadsheets stay in renderer memory and are never written to this folder.
if (app.isPackaged) {
  const portableRoot = process.env.PORTABLE_EXECUTABLE_DIR || path.dirname(process.execPath);
  const portableData = path.join(portableRoot, 'Datos locales ONAT');
  const sessionData = path.join(portableData, 'Sesion');
  try {
    fs.mkdirSync(sessionData, { recursive: true });
    app.setPath('userData', portableData);
    app.setPath('sessionData', sessionData);
  } catch {
    // If the drive is read-only, use Electron's standard per-user profile as a fallback.
  }
}

let mainWindow;

async function registerLocalAppProtocol() {
  const appRoot = path.resolve(app.getAppPath(), 'dist');
  await protocol.handle('app', (request) => {
    const requestUrl = new URL(request.url);
    if (requestUrl.hostname !== 'onat') {
      return new Response('Not found', { status: 404 });
    }

    const requestedPath = decodeURIComponent(requestUrl.pathname);
    const relativePath = requestedPath === '/' ? 'index.html' : requestedPath.replace(/^\/+/, '');
    const filePath = path.resolve(appRoot, relativePath);
    const pathFromRoot = path.relative(appRoot, filePath);
    if (pathFromRoot.startsWith('..') || path.isAbsolute(pathFromRoot)) {
      return new Response('Not found', { status: 404 });
    }

    return net.fetch(pathToFileURL(filePath).toString());
  });
}

async function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 960,
    minHeight: 640,
    show: false,
    backgroundColor: '#f4f7fa',
    icon: path.join(app.getAppPath(), 'assets', 'onat.ico'),
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      webSecurity: true,
    },
  });

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('https://')) {
      shell.openExternal(url).catch(() => {});
    }
    return { action: 'deny' };
  });
  mainWindow.webContents.on('will-navigate', (event, url) => {
    if (!url.startsWith('app://onat/')) event.preventDefault();
  });
  mainWindow.once('ready-to-show', () => {
    mainWindow.maximize();
    mainWindow.show();
  });

  await mainWindow.loadURL('app://onat/index.html');
}

app.whenReady().then(async () => {
  await registerLocalAppProtocol();
  await createWindow();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
