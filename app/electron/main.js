/**
 * Electron Main Process — W11 Theme Configurator
 *
 * Creates the application window, initialises the PowerShell bridge,
 * and registers all IPC handlers that the React renderer relies on.
 */

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
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
// Bidirectional mapping: store (schema-friendly) <-> PS (registry key names)
// -----------------------------------------------------------------------

// Taskbar: store names → registry key names (outbound: store→PS)
const TASKBAR_STORE_TO_REG = {
  alignment: 'TaskbarAl',
  showSearch: 'SearchboxTaskbarMode',
  showTaskView: 'ShowTaskViewButton',
  showWidgets: 'TaskbarDa',
  showChat: 'TaskbarMn',
  useOLEDTransparency: 'UseOLEDTaskbarTransparency'
};

// Taskbar: registry key names → store names (inbound: PS→store)
const TASKBAR_REG_TO_STORE = Object.fromEntries(
  Object.entries(TASKBAR_STORE_TO_REG).map(([k, v]) => [v, k])
);

// Section name mapping: PS output sections → store sections
const SECTION_PS_TO_STORE = {
  DarkMode: 'mode',
  AccentColor: 'accentColor',
  DWM: 'dwm',
  Taskbar: 'taskbar'
};

// DarkMode: registry keys → store keys (camelCase of PascalCase)
const DARKMODE_REG_TO_STORE = {
  AppsUseLightTheme: 'appsUseLightTheme',
  SystemUsesLightTheme: 'systemUsesLightTheme'
};

// AccentColor: registry keys → store keys
const ACCENT_REG_TO_STORE = {
  AccentColor: 'color',  // special: ABGR DWord → hex
  AccentColorMenu: null,  // skip: derived from AccentColor
  StartColorMenu: null,   // skip: derived from AccentColor
  ColorPrevalence_Personalize: 'colorPrevalence',
  ColorPrevalence_DWM: 'colorPrevalenceTitleBars',
  EnableTransparency: 'enableTransparency',
  AutoColorization: 'autoColorization'
};

// DWM: registry keys → store keys (just lowercase first letter)
const DWM_REG_TO_STORE = {
  ColorizationColor: 'colorizationColor',
  ColorizationAfterglow: 'colorizationAfterglow',
  ColorizationColorBalance: 'colorizationColorBalance',
  ColorizationAfterglowBalance: 'colorizationAfterglowBalance',
  ColorizationBlurBalance: 'colorizationBlurBalance',
  ColorizationGlassReflectionIntensity: 'colorizationGlassReflectionIntensity',
  EnableAeroPeek: 'enableAeroPeek',
  ForceEffectMode: 'forceEffectMode',
  AccentColorInactive: 'accentColorInactive',
  EnableWindowColorization: 'enableWindowColorization'
};

/**
 * Convert DWord (signed int32) to 0xAARRGGBB hex string.
 * Windows stores DWM colors as ABGR (0xAABBGGRR), so we convert.
 */
function dwordToHex(dword) {
  if (dword == null) return null;
  // Convert signed int to unsigned
  const unsigned = dword >>> 0;
  return '0x' + unsigned.toString(16).toUpperCase().padStart(8, '0');
}

/**
 * Convert ABGR DWord to #RRGGBB hex string for AccentColor.
 */
function abgrDwordToHexRGB(dword) {
  if (dword == null) return null;
  const unsigned = dword >>> 0;
  const R = unsigned & 0xFF;
  const G = (unsigned >>> 8) & 0xFF;
  const B = (unsigned >>> 16) & 0xFF;
  return '#' + [R, G, B].map(c => c.toString(16).padStart(2, '0')).join('').toUpperCase();
}

/**
 * Map the PS Get-W11RegistryTheme output to store-compatible format.
 * PS returns: { DarkMode: { AppsUseLightTheme: 0 }, Taskbar: { TaskbarAl: 0 }, ... }
 * Store expects: { mode: { appsUseLightTheme: 0 }, taskbar: { alignment: 0 }, ... }
 */
function mapPsOutputToStore(psData) {
  const result = {};

  // DarkMode → mode
  if (psData.DarkMode) {
    result.mode = {};
    for (const [regKey, val] of Object.entries(psData.DarkMode)) {
      const storeKey = DARKMODE_REG_TO_STORE[regKey];
      if (storeKey) result.mode[storeKey] = val;
    }
  }

  // AccentColor → accentColor
  if (psData.AccentColor) {
    result.accentColor = {};
    for (const [regKey, val] of Object.entries(psData.AccentColor)) {
      const storeKey = ACCENT_REG_TO_STORE[regKey];
      if (storeKey === null) continue; // skip derived keys
      if (storeKey === 'color') {
        result.accentColor.color = abgrDwordToHexRGB(val);
      } else if (storeKey) {
        result.accentColor[storeKey] = val;
      }
    }
  }

  // DWM → dwm
  if (psData.DWM) {
    result.dwm = {};
    for (const [regKey, val] of Object.entries(psData.DWM)) {
      const storeKey = DWM_REG_TO_STORE[regKey];
      if (storeKey) {
        // DWM color values are DWords — convert to hex strings
        if (storeKey.toLowerCase().includes('color') && typeof val === 'number') {
          result.dwm[storeKey] = dwordToHex(val);
        } else {
          result.dwm[storeKey] = val;
        }
      }
    }
  }

  // Taskbar → taskbar
  if (psData.Taskbar) {
    result.taskbar = {};
    for (const [regKey, val] of Object.entries(psData.Taskbar)) {
      const storeKey = TASKBAR_REG_TO_STORE[regKey];
      if (storeKey) result.taskbar[storeKey] = val;
    }
  }

  return result;
}

/**
 * Map outbound taskbar config from store names to registry key names.
 * Store sends: { taskbar: { alignment: 1 } }
 * PS expects: { taskbar: { TaskbarAl: 1 } }
 */
function mapTaskbarStoreToPs(storeTaskbar) {
  const mapped = {};
  for (const [storeKey, val] of Object.entries(storeTaskbar)) {
    const regKey = TASKBAR_STORE_TO_REG[storeKey];
    if (regKey) {
      mapped[regKey] = val;
    } else {
      mapped[storeKey] = val; // pass through unknown keys
    }
  }
  return mapped;
}

/**
 * Normalize a preset/theme config: handle "colors" vs "win32Colors" key.
 * Presets may use "colors" (schema name) — store uses "win32Colors" (PS name).
 */
function normalizeThemeConfig(config) {
  if (!config) return config;
  const result = { ...config };
  // Map "colors" → "win32Colors" if present
  if (result.colors && !result.win32Colors) {
    result.win32Colors = result.colors;
    delete result.colors;
  }
  return result;
}

/**
 * Extract JSON from PS stdout that may contain Write-Host noise.
 * PS functions like Get-W11InstalledThemes emit "Found N theme(s)."
 * before piping objects to ConvertTo-Json, polluting stdout.
 * This extracts the first valid JSON array or object from the output.
 */
function extractJson(stdout) {
  if (!stdout) return null;
  // Find the first [ or { which starts the JSON payload
  const arrStart = stdout.indexOf('[');
  const objStart = stdout.indexOf('{');
  let start = -1;
  if (arrStart === -1 && objStart === -1) return null;
  if (arrStart === -1) start = objStart;
  else if (objStart === -1) start = arrStart;
  else start = Math.min(arrStart, objStart);
  let jsonStr = stdout.substring(start);

  try {
    return JSON.parse(jsonStr);
  } catch (_firstErr) {
    // Encoding fallback: if the PS bridge produced non-UTF-8 bytes, Node's
    // UTF-8 decoder inserts U+FFFD replacement characters which break JSON.
    // Strip replacement characters and nearby damaged characters, then retry.
    const cleaned = jsonStr.replace(/\uFFFD/g, '');
    try {
      return JSON.parse(cleaned);
    } catch (_secondErr) {
      console.error('[extractJson] JSON parse failed even after cleanup. First 200 chars:', jsonStr.substring(0, 200));
      return null;
    }
  }
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
    const devPort = process.env.DEV_PORT || '5173';
    mainWindow.loadURL(`http://localhost:${devPort}`);
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
      const psData = extractJson(stdout);
      if (!psData) return { success: false, error: 'No JSON from Get-W11RegistryTheme' };
      // Map PS registry output → store-compatible format
      const storeData = mapPsOutputToStore(psData);
      return { success: true, data: storeData };
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

      // Map store-friendly property names → PS/registry key names for Taskbar
      let mappedConfig = { ...config };
      if (section === 'Taskbar' && mappedConfig.taskbar) {
        mappedConfig.taskbar = mapTaskbarStoreToPs(mappedConfig.taskbar);
      }

      const json = JSON.stringify(mappedConfig);
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
      // Map taskbar store names → registry key names for PS
      let mappedConfig = { ...config };
      if (mappedConfig.taskbar) {
        mappedConfig.taskbar = mapTaskbarStoreToPs(mappedConfig.taskbar);
      }
      // Normalize win32Colors/colors
      mappedConfig = normalizeThemeConfig(mappedConfig);

      const json = JSON.stringify(mappedConfig);
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
      const data = extractJson(stdout);
      if (!data) return { success: true, data: [] };
      return { success: true, data: Array.isArray(data) ? data : [data] };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // --- Load a preset ---
  ipcMain.handle('ps:load-preset', async (_event, name, source) => {
    try {
      const projectRoot = path.resolve(__dirname, '..', '..');

      // Determine which directory to look in based on source
      const subDir = (source === 'user') ? 'user' : 'presets';
      const filePath = path.join(projectRoot, 'config', subDir, `${name}.json`);

      // For built-in presets, use Get-W11ThemeConfig (handles basedOn inheritance).
      // For user presets (or if PS fails), fall back to direct file read.
      if (source !== 'user') {
        try {
          const { stdout } = await psBridge.execute(
            `Get-W11ThemeConfig -PresetName '${psEscape(name)}' | ConvertTo-Json -Depth 10 -Compress`
          );
          let data = extractJson(stdout);
          if (data) {
            data = normalizeThemeConfig(data);
            return { success: true, data };
          }
        } catch (_psErr) {
          // Fall through to direct file read
        }
      }

      // Direct file read (works for both user and preset files)
      if (!fs.existsSync(filePath)) {
        return { success: false, error: `Preset file not found: ${name}.json` };
      }
      const raw = fs.readFileSync(filePath, 'utf8');
      let data = JSON.parse(raw);
      data = normalizeThemeConfig(data);
      return { success: true, data };
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

      // Use Node fs directly (avoids PS escaping issues with large JSON)
      fs.mkdirSync(safeDir, { recursive: true });
      const json = JSON.stringify(config, null, 2);
      fs.writeFileSync(filePath, json, 'utf8');
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
