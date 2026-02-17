<p align="center">
  <img src="https://img.shields.io/badge/Windows%2011-25H2-0078D4?style=for-the-badge&logo=windows11&logoColor=white" alt="Windows 11 25H2"/>
  <img src="https://img.shields.io/badge/PowerShell-5.1+-5391FE?style=for-the-badge&logo=powershell&logoColor=white" alt="PowerShell 5.1+"/>
  <img src="https://img.shields.io/badge/Electron-34-47848F?style=for-the-badge&logo=electron&logoColor=white" alt="Electron 34"/>
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="MIT License"/>
  <img src="https://img.shields.io/badge/No%20Third--Party-Native%20Only-FF6B35?style=for-the-badge" alt="No Third-Party"/>
</p>

# Windows 11 Theming Suite

**A comprehensive, native Windows 11 theming toolkit with a full desktop GUI -- no third-party software required.**

Apply system-wide transparency, custom backdrops, cursor schemes, sound packs, wallpapers, and full registry-level theming -- through a graphical desktop application or directly via 51 PowerShell commands.

---

## Highlights

- **Desktop GUI** -- Electron + React app with live preview for every configurable parameter (148+)
- **Fully native** -- uses only documented Microsoft APIs (DWM, SWCA, XAML Diagnostics)
- **System-wide transparency** -- taskbar, Start Menu, Action Center, app windows, context menus
- **DLL injection engine** -- custom ShellTAP DLL for XAML-level property manipulation
- **Persistent** -- auto-recovers effects after process restarts via WMI event monitoring
- **Theme orchestrator** -- one command installs cursors, sounds, wallpaper, registry, and transparency
- **Modular architecture** -- 10 PowerShell modules, 2 native C++ DLLs, JSON config schema

---

## Desktop Application

The GUI application provides visual access to every configurable theme element with instant live preview:

**13 configurable sections:**
Mode, Accent Color, DWM Colorization, Taskbar, Win32 Colors (26 entries), Wallpaper, Transparency (6 subsystems), Cursors (15 roles), Sounds (29 events), Visual Styles, Desktop Icons, Advanced (registry overrides), Presets

**8 specialized controls:**
Toggle switch, Slider, Dropdown, Color Picker (#RRGGBB + "R G B"), Alpha Color Picker (0xAARRGGBB), File Picker, 8-slot Palette Editor, Registry Override Editor

### Running the App

```bash
# Install dependencies
cd app
npm install

# Development mode (with hot reload)
npm run dev

# Production build
npm run package
```

The app spawns a persistent PowerShell process that loads the full module and executes commands via IPC. Changes are applied instantly to the live system via registry writes + `WM_SETTINGCHANGE` broadcasts.

---

## Architecture

```
w11-theming-suite/
|-- w11-theming-suite.psm1        Root module loader
|-- w11-theming-suite.psd1        Module manifest (51 commands)
|-- app/                          Electron + React desktop GUI
|   |-- electron/
|   |   |-- main.js               Main process + IPC handlers
|   |   |-- preload.js            Context bridge (themeAPI)
|   |   +-- ps-bridge.js          Persistent PowerShell child process
|   |-- src/
|   |   |-- App.jsx               Root component + section routing
|   |   |-- store/themeStore.js   Zustand store (148+ params)
|   |   |-- hooks/useLivePreview.js  Debounced live preview engine
|   |   |-- components/
|   |   |   |-- controls/         8 reusable input controls
|   |   |   |-- layout/           Sidebar, TopBar, StatusBar
|   |   |   +-- sections/         13 settings sections
|   |   +-- styles/global.css     Dark theme UI
|   +-- package.json              Electron 34, React 19, Vite 6
|-- config/
|   |-- schema.json               JSON Schema for theme validation
|   |-- presets/                   Built-in theme presets (6 themes)
|   +-- user/                     User-created themes
|-- modules/
|   |-- ConfigManager/            Theme config loading, validation, merging
|   |-- RegistryConfigurator/     Registry-based theming (dark mode, accent, DWM)
|   |-- ThemeFileBuilder/         .theme file generation
|   |-- CursorSchemeBuilder/      Cursor scheme installation
|   |-- SoundSchemeBuilder/       Sound scheme installation
|   |-- WallpaperManager/         Wallpaper and slideshow management
|   |-- BackupRestore/            Full theme state backup/restore
|   |-- NativeTaskbarTransparency/ DWM + SWCA + TAP transparency engine
|   |-- TranslucentTBIntegration/ TranslucentTB fallback support
|   +-- ThemeOrchestrator/        High-level install/uninstall/switch pipeline
|-- native/
|   |-- TaskbarTAP/               Taskbar XAML injection DLL (C++)
|   |-- ShellTAP/                 Generic Shell XAML injection DLL (C++)
|   +-- bin/                      Pre-built x64 binaries
|-- scripts/                      Standalone utility scripts
+-- tests/                        Diagnostic and integration tests
```

---

## Requirements

| Requirement | Minimum |
|-------------|---------|
| **OS** | Windows 11 Build 22621+ (22H2) |
| **Tested** | Windows 11 Build 26200 (25H2) |
| **PowerShell** | 5.1+ (ships with Windows) |
| **Node.js** | 18+ (for the desktop app only) |
| **Privileges** | Administrator (for DLL injection) |
| **Third-party** | None |

---

## Quick Start

### Option 1: Desktop Application (Recommended)

```bash
git clone https://github.com/decarvalhoe/w11-theming-suite.git
cd w11-theming-suite/app
npm install
npm run dev
```

### Option 2: PowerShell CLI

```powershell
# 1. Clone the repository
git clone https://github.com/decarvalhoe/w11-theming-suite.git
cd w11-theming-suite

# 2. Import the module
Import-Module .\w11-theming-suite.psd1 -Force

# 3. Apply a built-in theme
Install-W11Theme -PresetName "deep-dark"

# 4. Or apply individual transparency effects
Set-W11NativeTaskbarTransparency -Style clear
Invoke-StartMenuTransparency -Mode Transparent
Invoke-ActionCenterTransparency -Mode Acrylic
Start-W11BackdropWatcher -Style mica -DarkMode -IncludeContextMenus
```

---

## Transparency Features

### Taskbar

```powershell
# SWCA approach (fast, no injection)
Set-W11NativeTaskbarTransparency -Style clear
Set-W11NativeTaskbarTransparency -Style acrylic -Color '#80000000' -AllMonitors

# TAP approach (XAML-level, deeper control)
Invoke-TaskbarTAPInject -Mode Transparent
```

### Start Menu

```powershell
# Discovery mode (find XAML element names)
Invoke-StartMenuDiscovery

# Apply transparency
Invoke-StartMenuTransparency -Mode Transparent
```

### Action Center & Notifications

```powershell
Invoke-ActionCenterDiscovery
Invoke-ActionCenterTransparency -Mode Acrylic
```

### App Windows (Persistent)

```powershell
# Apply Mica/Acrylic/Tabbed backdrop to ALL windows, present and future
Start-W11BackdropWatcher -Style mica -DarkMode -IncludeContextMenus

# Check status
Stop-W11BackdropWatcher
```

### Context Menus

```powershell
# One-shot: apply to currently visible menus
Set-W11ContextMenuBackdrop -Style acrylic

# Persistent: handled by BackdropWatcher
Start-W11BackdropWatcher -Style acrylic -IncludeContextMenus
```

---

## Theme Configuration

Themes are defined as JSON files validated against `config/schema.json`:

```json
{
  "meta": {
    "name": "My Theme",
    "version": "1.0.0",
    "author": "You"
  },
  "mode": {
    "appsUseLightTheme": 0,
    "systemUsesLightTheme": 0
  },
  "transparency": {
    "taskbar": { "enabled": true, "style": "clear" },
    "startMenu": { "enabled": true, "mode": "Transparent" },
    "actionCenter": { "enabled": true, "mode": "Acrylic" },
    "appWindows": { "enabled": true, "backdrop": "mica", "darkMode": true },
    "contextMenus": { "enabled": true },
    "persist": true
  }
}
```

### Built-in Presets

| Preset | Description |
|--------|-------------|
| `deep-dark` | Full dark mode with transparent taskbar |
| `pure-black` | OLED-optimized pure black theme |
| `macos-monterey` | macOS-inspired light theme |
| `ubuntu-yaru` | Ubuntu Yaru color scheme |
| `retro-xp` | Windows XP nostalgic theme |
| `custom-template` | Blank template for custom themes |

---

## Persistence & Auto-Recovery

Register all transparency effects for login persistence with automatic re-injection when processes restart:

```powershell
Register-W11TransparencyPersistence `
    -TaskbarStyle clear `
    -TaskbarTAP `
    -StartMenu `
    -ActionCenter `
    -AppWindows -AppWindowsBackdrop mica -AppWindowsDarkMode `
    -ContextMenus

# Check status
Get-W11TransparencyPersistence

# Remove
Unregister-W11TransparencyPersistence
```

The persistence script monitors `explorer.exe`, `StartMenuExperienceHost.exe`, and `ShellExperienceHost.exe` via WMI `__InstanceCreationEvent` and re-injects automatically within seconds of a process restart.

---

## How It Works

### DWM Backdrop (App Windows & Context Menus)
Uses the documented `DwmSetWindowAttribute` API with the [SetMica technique](https://github.com/tringi/setmica):
1. `DwmExtendFrameIntoClientArea(MARGINS -1)` -- sheet of glass
2. `DwmSetWindowAttribute(DWMWA_SYSTEMBACKDROP_TYPE)` -- set material
3. `DwmSetWindowAttribute(DWMWA_CAPTION_COLOR, COLOR_NONE)` -- remove caption paint

### SWCA Taskbar Transparency
Uses the undocumented `SetWindowCompositionAttribute` API to directly control the taskbar's composition accent state.

### ShellTAP DLL Injection (Start Menu, Action Center)
1. PowerShell writes a config struct to named shared memory
2. `CreateRemoteThread(LoadLibraryW)` injects `ShellTAP.dll` into the target process
3. The DLL calls `InitializeXamlDiagnosticsEx` from within the target process
4. Uses `GetPropertyValuesChain` + `SetProperty` to modify XAML elements (opacity, visibility, brush)

### BackdropWatcher (Persistent)
A C# class running on a dedicated thread with a Win32 message pump, using `SetWinEventHook` to monitor:
- `EVENT_OBJECT_SHOW` (0x8002) -- new windows appearing
- `EVENT_SYSTEM_FOREGROUND` (0x0003) -- focus changes
- `EVENT_SYSTEM_MENUPOPUPSTART` (0x0006) -- context menus opening

### Electron GUI Architecture
The desktop app uses a persistent PowerShell child process as its backend:
1. `ps-bridge.js` spawns `powershell.exe`, imports the module, and maintains a command queue with sentinel-based completion detection
2. IPC handlers in `main.js` bridge Electron's `ipcMain` to PowerShell commands with input validation and escaping
3. The React renderer uses Zustand for state management and a debounced `useLivePreview` hook that applies changes to the live system as the user adjusts controls

---

## All Exported Commands (51)

<details>
<summary>Click to expand full command list</summary>

**Theme Management**
- `Install-W11Theme` / `Uninstall-W11Theme` / `Switch-W11Theme`
- `Get-W11InstalledThemes`

**Configuration**
- `Get-W11ThemeConfig` / `New-W11ThemeConfig` / `Merge-W11ThemeConfig`

**Backup & Restore**
- `Backup-W11ThemeState` / `Restore-W11ThemeState`
- `Get-W11ThemeBackups` / `Remove-W11ThemeBackup`

**Registry**
- `Set-W11RegistryTheme` / `Get-W11RegistryTheme`

**Theme Files**
- `New-W11ThemeFile`

**Cursors**
- `Install-W11CursorScheme` / `Uninstall-W11CursorScheme` / `Get-W11CursorSchemes`

**Sounds**
- `Install-W11SoundScheme` / `Uninstall-W11SoundScheme` / `Get-W11SoundSchemes`

**Wallpaper**
- `Set-W11Wallpaper` / `Get-W11Wallpaper`

**Taskbar Transparency (SWCA)**
- `Set-W11NativeTaskbarTransparency` / `Get-W11NativeTaskbarTransparency` / `Remove-W11NativeTaskbarTransparency`
- `Register-W11TaskbarTransparencyStartup` / `Unregister-W11TaskbarTransparencyStartup`

**Taskbar Transparency (TAP)**
- `Invoke-TaskbarTAPInject` / `Set-TaskbarTAPMode` / `Get-TaskbarExplorerPid`

**Shell Transparency (ShellTAP)**
- `Invoke-ShellTAPInject` / `Set-ShellTAPMode`
- `Invoke-StartMenuDiscovery` / `Invoke-StartMenuTransparency`
- `Invoke-ActionCenterDiscovery` / `Invoke-ActionCenterTransparency`

**Context Menus**
- `Set-W11ContextMenuBackdrop` / `Get-W11ContextMenuBackdropStatus`

**App Window Backdrop**
- `Set-W11WindowBackdrop` / `Set-W11WindowColors`
- `Start-W11BackdropWatcher` / `Stop-W11BackdropWatcher`
- `Register-W11BackdropWatcherStartup` / `Unregister-W11BackdropWatcherStartup`

**Unified Persistence**
- `Register-W11TransparencyPersistence` / `Unregister-W11TransparencyPersistence` / `Get-W11TransparencyPersistence`

**TranslucentTB Integration**
- `Test-W11TranslucentTBInstalled` / `Get-W11TranslucentTBConfig` / `Set-W11TranslucentTBConfig`

**Utility**
- `Get-W11ProjectRoot`

</details>

---

## Development

### Desktop App

```bash
cd app
npm install
npm run dev      # Vite dev server + Electron with hot reload
npm run build    # Build production bundle
npm run package  # Build + electron-builder installer
```

**Tech stack:** Electron 34, React 19, Vite 6, Zustand 5, react-colorful

### Building the Native DLLs

Requires Visual Studio Build Tools with MSVC x64:

```cmd
cd native\ShellTAP
build.cmd

cd native\TaskbarTAP
build.cmd
```

Pre-built binaries are included in `native/bin/`.

### Branch Strategy

| Branch | Purpose |
|--------|---------|
| `DEV` | Active development (default) |
| `STAGE` | Pre-release testing |
| `PROD` | Stable releases |

---

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines on submitting issues, feature requests, and pull requests.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

## Code of Conduct

This project follows the Contributor Covenant Code of Conduct. See [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md).
