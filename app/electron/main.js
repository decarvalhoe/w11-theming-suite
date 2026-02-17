/**
 * Electron Main Process — W11 Theme Configurator
 *
 * Creates the application window, initialises the PowerShell bridge,
 * and registers all IPC handlers that the React renderer relies on.
 */

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const PowerShellBridge = require('./ps-bridge');

let mainWindow = null;
let psBridge = null;

const isDev = !app.isPackaged;

// -----------------------------------------------------------------------
// Valid sections allowlist — FIX C4: prevent PS injection via section param
// -----------------------------------------------------------------------
const VALID_SECTIONS = new Set(['DarkMode', 'AccentColor', 'DWM', 'Taskbar']);

// -----------------------------------------------------------------------
// Helper: escape a string for embedding in PowerShell single-quoted strings
// -----------------------------------------------------------------------
function psEscape(str) {
  if (typeof str !== 'string') return '';
  return str.replace(/'/g, "''");
}

// -----------------------------------------------------------------------
// Window creation
// -----------------------------------------------------------------------
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 960,
    minHeight: 640,
    backgroundColor: '#0a0a0a',
    darkTheme: true,
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }
}

// -----------------------------------------------------------------------
// IPC Handlers
// -----------------------------------------------------------------------
function registerIpcHandlers() {
  // --- Read current theme state from registry ---
  ipcMain.handle('ps:read-theme', async () => {
    try {
      const { stdout } = await psBridge.execute(
        'Get-W11RegistryTheme | ConvertTo-Json -Depth 10 -Compress'
      );
      return { success: true, data: JSON.parse(stdout) };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Apply a single registry section (DarkMode, AccentColor, DWM, Taskbar) ---
  ipcMain.handle('ps:apply-section', async (_event, section, config) => {
    try {
      // FIX C4: Validate section against allowlist
      if (!VALID_SECTIONS.has(section)) {
        return { success: false, error: `Invalid section: ${section}` };
      }
      const json = JSON.stringify(config);
      const cmd = `$cfg = '${psEscape(json)}' | ConvertFrom-Json; Set-W11RegistryTheme -Config $cfg -Section ${section}`;
      const { stderr } = await psBridge.execute(cmd);
      return { success: true, warning: stderr || null };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Apply win32Colors ---
  ipcMain.handle('ps:apply-win32colors', async (_event, colors) => {
    try {
      const config = { win32Colors: colors };
      const json = JSON.stringify(config);
      const cmd = `$cfg = '${psEscape(json)}' | ConvertFrom-Json; Set-W11RegistryTheme -Config $cfg`;
      const { stderr } = await psBridge.execute(cmd);
      return { success: true, warning: stderr || null };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Apply full theme config ---
  ipcMain.handle('ps:apply-full-theme', async (_event, config) => {
    try {
      const json = JSON.stringify(config);
      const cmd = `$cfg = '${psEscape(json)}' | ConvertFrom-Json; Set-W11RegistryTheme -Config $cfg`;
      const { stderr } = await psBridge.execute(cmd);
      return { success: true, warning: stderr || null };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- List presets ---
  ipcMain.handle('ps:list-presets', async () => {
    try {
      const { stdout } = await psBridge.execute(
        'Get-W11InstalledThemes | ConvertTo-Json -Depth 5 -Compress'
      );
      const data = stdout ? JSON.parse(stdout) : [];
      return { success: true, data: Array.isArray(data) ? data : [data] };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Load a preset ---
  ipcMain.handle('ps:load-preset', async (_event, name) => {
    try {
      const { stdout } = await psBridge.execute(
        `Get-W11ThemeConfig -PresetName '${psEscape(name)}' | ConvertTo-Json -Depth 10 -Compress`
      );
      return { success: true, data: JSON.parse(stdout) };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Save preset ---
  ipcMain.handle('ps:save-preset', async (_event, name, config) => {
    try {
      const projectRoot = path.resolve(__dirname, '..', '..');
      const safeDir = path.join(projectRoot, 'config', 'user');
      const filePath = path.join(safeDir, `${name}.json`);

      // FIX I1: Prevent path traversal — resolved path must stay inside config/user/
      if (!filePath.startsWith(safeDir + path.sep) && filePath !== safeDir) {
        return { success: false, error: 'Invalid preset name: path traversal detected.' };
      }

      const json = JSON.stringify(config, null, 2);
      const cmd = `Set-Content -Path '${psEscape(filePath)}' -Value '${psEscape(json)}' -Encoding UTF8`;
      await psBridge.execute(cmd);
      return { success: true };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Backup ---
  ipcMain.handle('ps:backup', async (_event, name) => {
    try {
      await psBridge.execute(`Backup-W11ThemeState -Name '${psEscape(name)}'`);
      return { success: true };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Restore ---
  ipcMain.handle('ps:restore', async (_event, name) => {
    try {
      await psBridge.execute(`Restore-W11ThemeState -Name '${psEscape(name)}'`);
      return { success: true };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Wallpaper ---
  ipcMain.handle('ps:apply-wallpaper', async (_event, config) => {
    try {
      const json = JSON.stringify({ wallpaper: config });
      const cmd = `$cfg = '${psEscape(json)}' | ConvertFrom-Json; Set-W11Wallpaper -Config $cfg`;
      const { stderr } = await psBridge.execute(cmd);
      return { success: true, warning: stderr || null };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Transparency ---
  ipcMain.handle('ps:apply-transparency', async (_event, subsystem, params) => {
    try {
      // FIX C4: Validate subsystem and escape all params
      const VALID_SUBSYSTEMS = new Set(['taskbar', 'taskbarTAP', 'startMenu', 'actionCenter', 'appWindows']);
      if (!VALID_SUBSYSTEMS.has(subsystem)) {
        return { success: false, error: `Unknown subsystem: ${subsystem}` };
      }

      let cmd = '';
      switch (subsystem) {
        case 'taskbar':
          cmd = `Set-W11NativeTaskbarTransparency -Style '${psEscape(params.style || 'clear')}'`;
          if (params.color) cmd += ` -Color '${psEscape(params.color)}'`;
          if (params.allMonitors) cmd += ' -AllMonitors';
          break;
        case 'taskbarTAP':
          cmd = `Invoke-TaskbarTAPInject -Mode '${psEscape(params.mode || 'Transparent')}'`;
          break;
        case 'startMenu':
          cmd = `Invoke-StartMenuTransparency -Mode '${psEscape(params.mode || 'Transparent')}'`;
          break;
        case 'actionCenter':
          cmd = `Invoke-ActionCenterTransparency -Mode '${psEscape(params.mode || 'Transparent')}'`;
          break;
        case 'appWindows':
          cmd = `Start-W11BackdropWatcher -Style '${psEscape(params.backdrop || 'mica')}'`;
          if (params.darkMode) cmd += ' -DarkMode';
          break;
      }
      const { stderr } = await psBridge.execute(cmd);
      return { success: true, warning: stderr || null };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Cursors ---
  ipcMain.handle('ps:apply-cursors', async (_event, config) => {
    try {
      const json = JSON.stringify({ cursors: config });
      const cmd = `$cfg = '${psEscape(json)}' | ConvertFrom-Json; Install-W11CursorScheme -Config $cfg -Activate`;
      const { stderr } = await psBridge.execute(cmd);
      return { success: true, warning: stderr || null };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Sounds ---
  ipcMain.handle('ps:apply-sounds', async (_event, config) => {
    try {
      const json = JSON.stringify({ sounds: config });
      const cmd = `$cfg = '${psEscape(json)}' | ConvertFrom-Json; Install-W11SoundScheme -Config $cfg -Activate`;
      const { stderr } = await psBridge.execute(cmd);
      return { success: true, warning: stderr || null };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Native file dialog ---
  ipcMain.handle('dialog:open-file', async (_event, filters) => {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openFile'],
      filters: filters || [{ name: 'All Files', extensions: ['*'] }]
    });
    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0];
  });
}

// -----------------------------------------------------------------------
// App lifecycle
// -----------------------------------------------------------------------
app.whenReady().then(async () => {
  // Start PS bridge
  psBridge = new PowerShellBridge();

  psBridge.on('ready', () => {
    if (mainWindow) {
      mainWindow.webContents.send('bridge:status', 'connected');
    }
  });

  psBridge.on('closed', () => {
    if (mainWindow) {
      mainWindow.webContents.send('bridge:status', 'disconnected');
    }
  });

  psBridge.on('error', (err) => {
    console.error('PS Bridge error:', err);
    if (mainWindow) {
      mainWindow.webContents.send('bridge:status', 'error');
    }
  });

  registerIpcHandlers();
  createWindow();

  try {
    await psBridge.start();
    // Create automatic session backup for undo
    await psBridge.execute("Backup-W11ThemeState -Name 'app-session' -Force");
    console.log('[PS Bridge] Ready. Session backup created.');
  } catch (err) {
    console.error('[PS Bridge] Failed to start:', err);
  }
});

app.on('window-all-closed', () => {
  if (psBridge) psBridge.stop();
  app.quit();
});
