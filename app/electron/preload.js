/**
 * Preload script — exposes a safe `window.themeAPI` bridge to the renderer
 * via Electron's contextBridge.
 */

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('themeAPI', {
  // ---- Theme State ----
  /** Read the current Windows theme state from the registry. */
  readTheme: () => ipcRenderer.invoke('ps:read-theme'),

  /** Apply a single section to the live system. */
  applySection: (section, config) => ipcRenderer.invoke('ps:apply-section', section, config),

  /** Apply win32Colors to the live system. */
  applyWin32Colors: (colors) => ipcRenderer.invoke('ps:apply-win32colors', colors),

  /** Apply the full theme config at once. */
  applyFullTheme: (config) => ipcRenderer.invoke('ps:apply-full-theme', config),

  // ---- Presets ----
  /** List all available presets. */
  listPresets: () => ipcRenderer.invoke('ps:list-presets'),

  /** Load a preset by name and return its config. */
  loadPreset: (name, source) => ipcRenderer.invoke('ps:load-preset', name, source),

  /** Save current config as a user preset. */
  savePreset: (name, config) => ipcRenderer.invoke('ps:save-preset', name, config),

  // ---- Backup / Undo ----
  /** Create a backup of the current theme state. */
  backup: (name) => ipcRenderer.invoke('ps:backup', name),

  /** Restore from a named backup. */
  restore: (name) => ipcRenderer.invoke('ps:restore', name),

  // ---- Wallpaper ----
  applyWallpaper: (config) => ipcRenderer.invoke('ps:apply-wallpaper', config),

  // ---- Transparency ----
  applyTransparency: (subsystem, params) => ipcRenderer.invoke('ps:apply-transparency', subsystem, params),

  // ---- Cursors / Sounds ----
  applyCursors: (config) => ipcRenderer.invoke('ps:apply-cursors', config),
  applySounds: (config) => ipcRenderer.invoke('ps:apply-sounds', config),

  // ---- Dialogs ----
  /** Open a native file dialog, returns selected path or null. */
  openFile: (filters) => ipcRenderer.invoke('dialog:open-file', filters),

  // ---- Bridge Status ----
  onBridgeStatus: (callback) => {
    const handler = (_event, status) => callback(status);
    ipcRenderer.on('bridge:status', handler);
    return () => ipcRenderer.removeListener('bridge:status', handler);
  }
});
