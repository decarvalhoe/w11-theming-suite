import { create } from 'zustand';

/**
 * Central Zustand store that mirrors the full theme JSON structure.
 * Every control reads/writes from here; the live-preview hook watches
 * for changes and dispatches the appropriate IPC call.
 */

const defaultTheme = {
  meta: { name: 'Untitled', version: '1.0.0', author: '', description: '', tags: [], basedOn: null },
  mode: { appsUseLightTheme: 0, systemUsesLightTheme: 0 },
  accentColor: {
    color: '#0078D4',
    colorPrevalence: 0,
    colorPrevalenceTitleBars: 0,
    autoColorization: 0,
    enableTransparency: 1,
    accentPalette: ['#99EBFF', '#4CC2FF', '#0078D4', '#0067C0', '#003E92', '#001A68', '#00003C', '#FFFFFF']
  },
  dwm: {
    colorizationColor: '0xC40078D4',
    colorizationAfterglow: '0xC40078D4',
    colorizationColorBalance: 50,
    colorizationAfterglowBalance: 10,
    colorizationBlurBalance: 70,
    colorizationGlassReflectionIntensity: 0,
    enableAeroPeek: 1,
    forceEffectMode: 0,
    accentColorInactive: '0xFF2B2B2B',
    enableWindowColorization: 1
  },
  taskbar: {
    alignment: 1,
    showSearch: 1,
    showTaskView: 0,
    showWidgets: 0,
    showChat: 0,
    useOLEDTransparency: 0
  },
  win32Colors: {
    ActiveTitle: '0 0 0', Background: '0 0 0', Hilight: '0 120 215', HilightText: '255 255 255',
    TitleText: '255 255 255', Window: '0 0 0', WindowText: '255 255 255', Scrollbar: '20 20 20',
    InactiveTitle: '40 40 40', InactiveTitleText: '128 128 128', Menu: '20 20 20', MenuText: '255 255 255',
    ActiveBorder: '0 0 0', InactiveBorder: '0 0 0', AppWorkspace: '0 0 0', ButtonFace: '30 30 30',
    ButtonShadow: '0 0 0', GrayText: '128 128 128', ButtonText: '255 255 255', ButtonHilight: '50 50 50',
    ButtonDkShadow: '0 0 0', ButtonLight: '40 40 40', InfoText: '255 255 255', InfoWindow: '20 20 20',
    GradientActiveTitle: '0 0 0', GradientInactiveTitle: '0 0 0', WindowFrame: '0 0 0',
    MenuHilight: '30 30 30', MenuBar: '0 0 0', HotTrackingColor: '0 120 215'
  },
  wallpaper: { mode: 'static', path: '', style: 10, tileWallpaper: 0, color: '#000000' },
  transparency: {
    taskbar: { enabled: false, style: 'clear', color: '#00000000', allMonitors: false },
    taskbarTAP: { enabled: false, mode: 'Transparent' },
    startMenu: { enabled: false, mode: 'Transparent' },
    actionCenter: { enabled: false, mode: 'Transparent' },
    appWindows: { enabled: false, backdrop: 'mica', darkMode: true },
    contextMenus: { enabled: false },
    persist: false
  },
  cursors: { schemeName: '', setFolder: '', roles: {} },
  sounds: { schemeName: '', setFolder: '', events: {}, explorerEvents: {} },
  visualStyles: {
    path: '%SystemRoot%\\resources\\Themes\\Aero\\Aero.msstyles',
    colorStyle: 'NormalColor', size: 'NormalSize',
    colorizationColor: '0xC40078D4', transparency: 1, composition: 1
  },
  desktopIcons: { computer: '', documents: '', network: '', recycleBinFull: '', recycleBinEmpty: '' },
  advanced: { registryOverrides: [], thirdParty: { msstyleEditorPath: '', micaForEveryonePath: '' } }
};

const useThemeStore = create((set, get) => ({
  // --- State ---
  theme: { ...defaultTheme },
  activeSection: 'mode',
  bridgeStatus: 'connecting',   // 'connecting' | 'connected' | 'disconnected' | 'error'
  lastAction: null,             // { text, time, success }
  dirty: false,                 // Has the user made changes since last save/load?

  // --- Actions ---

  /** Set the active sidebar section. */
  setActiveSection: (section) => set({ activeSection: section }),

  /** Update a single field within a theme section. */
  setField: (section, key, value) => set((state) => ({
    theme: {
      ...state.theme,
      [section]: {
        ...state.theme[section],
        [key]: value
      }
    },
    dirty: true
  })),

  /** Replace an entire section's config. */
  setSection: (section, data) => set((state) => ({
    theme: { ...state.theme, [section]: data },
    dirty: true
  })),

  /** Load a full theme config (from preset or backup). */
  loadTheme: (themeData) => set({
    theme: { ...defaultTheme, ...themeData },
    dirty: false
  }),

  /** Reset to defaults. */
  resetTheme: () => set({ theme: { ...defaultTheme }, dirty: false }),

  /** Update bridge status. */
  setBridgeStatus: (status) => set({ bridgeStatus: status }),

  /** Log last applied action. */
  setLastAction: (text, success = true) => set({
    lastAction: { text, time: new Date().toLocaleTimeString(), success }
  }),

  /** Mark as saved. */
  markClean: () => set({ dirty: false })
}));

export default useThemeStore;
