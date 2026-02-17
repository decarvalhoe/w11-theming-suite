Set-StrictMode -Version Latest

# ===========================================================================
# NativeTaskbarTransparency.psm1
# ===========================================================================
# Windows 11 native DWM theming -- no third-party software required.
#
# Uses DwmSetWindowAttribute (dwmapi.dll) + DwmExtendFrameIntoClientArea
# following the "SetMica" technique (github.com/tringi/setmica):
#   1. ExtendFrame(MARGINS -1) -- "sheet of glass" for backdrop rendering
#   2. SetBackdropType(Mica/Acrylic/Tabbed) -- set the material
#   3. SetCaptionColor(COLOR_NONE) -- remove caption paint to show backdrop
#
# WHAT WORKS from an external process (confirmed on Build 26200 / 25H2):
#
#   COLORS (border, caption, text, dark mode):
#   [YES] Notepad (WinUI3)            -- border, caption, text, dark mode
#   [YES] Terminal (WinUI3)            -- border, caption, text, dark mode
#   [YES] File Explorer (CabinetWClass)-- border, caption, text, dark mode
#   [YES] UWP apps (ApplicationFrame)  -- border, caption, text, dark mode
#   [NO]  Chrome/Edge/Electron         -- Chromium renders its own frame
#   [NO]  ConsoleWindowClass           -- INVALID_HANDLE (protected process)
#   [NO]  TaskManagerWindow            -- INVALID_HANDLE (protected process)
#
#   BACKDROPS (Mica, Acrylic, Tabbed):
#   [YES] Terminal                     -- full backdrop (transparent client area)
#   [PARTIAL] Notepad, Explorer, UWP   -- title bar area only (opaque client)
#   [NO]  Chrome/Edge/Electron         -- paints own client area
#   [NO]  Taskbar (Shell_TrayWnd)      -- XAML taskbar ignores DWM + SWCA
#
#   NOTES:
#   - API calls return S_OK even when effects are not visually rendered.
#   - Apps that paint their own non-client area (Chromium) override DWM.
#   - TaskbarAnimations=0 or VisualFX="Best Performance" may suppress effects.
#   - Outdated GPU drivers may cause backdrop effects to silently fail.
#   - The old SWCA API is kept as fallback for taskbar transparency.
#
# Requires: Windows 11 Build 22621+ (22H2), PowerShell 5.1+
# ===========================================================================

# ---------------------------------------------------------------------------
# P/Invoke: NEW DWM-based approach (official Microsoft API)
# ---------------------------------------------------------------------------
$dwmTypeDefinition = @'
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Threading;

namespace W11ThemeSuite {
    // MARGINS struct for DwmExtendFrameIntoClientArea
    [StructLayout(LayoutKind.Sequential)]
    public struct MARGINS {
        public int cxLeftWidth;
        public int cxRightWidth;
        public int cyTopHeight;
        public int cyBottomHeight;

        // Constructor: all margins to same value (-1 = sheet of glass)
        public MARGINS(int allMargins) {
            cxLeftWidth = allMargins;
            cxRightWidth = allMargins;
            cyTopHeight = allMargins;
            cyBottomHeight = allMargins;
        }
    }

    public static class DwmHelper {
        // DwmSetWindowAttribute - THE official Microsoft API for window effects
        [DllImport("dwmapi.dll", PreserveSig = true)]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

        [DllImport("dwmapi.dll", PreserveSig = true)]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref uint pvAttribute, int cbAttribute);

        // DwmExtendFrameIntoClientArea - CRITICAL for making backdrops visible!
        // Without this call, DwmSetWindowAttribute sets the backdrop type but
        // the effect is NOT rendered. You must extend the DWM frame into the
        // client area using MARGINS(-1) ("sheet of glass") for the backdrop
        // material to actually appear.
        [DllImport("dwmapi.dll", PreserveSig = true)]
        public static extern int DwmExtendFrameIntoClientArea(IntPtr hwnd, ref MARGINS pMarInset);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        // EnumWindows for applying to all visible windows
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        // DWMWINDOWATTRIBUTE values (official Microsoft enum)
        public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        public const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
        public const int DWMWA_BORDER_COLOR = 34;
        public const int DWMWA_CAPTION_COLOR = 35;
        public const int DWMWA_TEXT_COLOR = 36;
        public const int DWMWA_SYSTEMBACKDROP_TYPE = 38;

        // DWM_SYSTEMBACKDROP_TYPE values
        public const int DWMSBT_AUTO = 0;              // Let DWM decide
        public const int DWMSBT_NONE = 1;              // No backdrop
        public const int DWMSBT_MAINWINDOW = 2;        // Mica
        public const int DWMSBT_TRANSIENTWINDOW = 3;   // Desktop Acrylic
        public const int DWMSBT_TABBEDWINDOW = 4;      // Mica Alt

        // DWMWA_COLOR special values
        public const uint DWMWA_COLOR_NONE = 0xFFFFFFFE;    // No border
        public const uint DWMWA_COLOR_DEFAULT = 0xFFFFFFFF;  // System default

        // Extend the DWM frame into the client area.
        // MARGINS(-1) = "sheet of glass" -- makes the entire window area
        // eligible for DWM backdrop rendering.
        public static int ExtendFrame(IntPtr hwnd) {
            var margins = new MARGINS(-1);
            return DwmExtendFrameIntoClientArea(hwnd, ref margins);
        }

        // Reset the DWM frame extension to default (no extension).
        public static int ResetFrame(IntPtr hwnd) {
            var margins = new MARGINS(0);
            return DwmExtendFrameIntoClientArea(hwnd, ref margins);
        }

        // Apply backdrop type to a window.
        // Uses the SetMica technique (github.com/tringi/setmica):
        //   1. DwmExtendFrameIntoClientArea -- extend DWM frame
        //   2. DwmSetWindowAttribute(DWMWA_SYSTEMBACKDROP_TYPE) -- set backdrop
        //   3. DwmSetWindowAttribute(DWMWA_CAPTION_COLOR, COLOR_NONE) -- remove
        //      caption color so the backdrop shows through the title bar
        // Without steps 1+3, the API succeeds but the backdrop is NOT visible.
        public static int SetBackdropType(IntPtr hwnd, int backdropType) {
            if (backdropType == DWMSBT_AUTO || backdropType == DWMSBT_NONE) {
                // Resetting: clear backdrop, restore default caption color, reset frame
                int hr = DwmSetWindowAttribute(hwnd, DWMWA_SYSTEMBACKDROP_TYPE, ref backdropType, sizeof(int));
                uint defaultColor = DWMWA_COLOR_DEFAULT;
                DwmSetWindowAttribute(hwnd, DWMWA_CAPTION_COLOR, ref defaultColor, sizeof(uint));
                ResetFrame(hwnd);
                return hr;
            } else {
                // Applying the full SetMica sequence:
                // Step 1: Extend DWM frame into client area (sheet of glass)
                ExtendFrame(hwnd);
                // Step 2: Set the backdrop type (Mica, Acrylic, Tabbed)
                int hr = DwmSetWindowAttribute(hwnd, DWMWA_SYSTEMBACKDROP_TYPE, ref backdropType, sizeof(int));
                // Step 3: Remove caption color so backdrop shows through title bar
                uint noneColor = DWMWA_COLOR_NONE;
                DwmSetWindowAttribute(hwnd, DWMWA_CAPTION_COLOR, ref noneColor, sizeof(uint));
                return hr;
            }
        }

        // Apply immersive dark mode to a window
        public static int SetDarkMode(IntPtr hwnd, bool enable) {
            int value = enable ? 1 : 0;
            return DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
        }

        // Set border color (COLORREF = 0x00BBGGRR)
        public static int SetBorderColor(IntPtr hwnd, uint color) {
            return DwmSetWindowAttribute(hwnd, DWMWA_BORDER_COLOR, ref color, sizeof(uint));
        }

        // Set caption color (COLORREF = 0x00BBGGRR)
        public static int SetCaptionColor(IntPtr hwnd, uint color) {
            return DwmSetWindowAttribute(hwnd, DWMWA_CAPTION_COLOR, ref color, sizeof(uint));
        }

        // Set text color (COLORREF = 0x00BBGGRR)
        public static int SetTextColor(IntPtr hwnd, uint color) {
            return DwmSetWindowAttribute(hwnd, DWMWA_TEXT_COLOR, ref color, sizeof(uint));
        }

        // Get all visible top-level windows
        public static List<IntPtr> GetVisibleWindows() {
            var windows = new List<IntPtr>();
            EnumWindows((hWnd, lParam) => {
                if (IsWindowVisible(hWnd)) {
                    // GWL_EXSTYLE = -20, WS_EX_TOOLWINDOW = 0x80, WS_EX_NOACTIVATE = 0x08000000
                    int exStyle = GetWindowLong(hWnd, -20);
                    bool isToolWindow = (exStyle & 0x80) != 0;
                    bool hasTitle = GetWindowTextLength(hWnd) > 0;

                    // Only include windows with title bars (real app windows)
                    if (!isToolWindow && hasTitle) {
                        windows.Add(hWnd);
                    }
                }
                return true;
            }, IntPtr.Zero);
            return windows;
        }

        // Get window class name
        public static string GetWindowClassName(IntPtr hwnd) {
            var sb = new System.Text.StringBuilder(256);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }
    }

    // =====================================================================
    // BackdropWatcher -- event-driven persistent backdrop for ALL app windows
    // Uses SetWinEventHook to detect new windows + focus changes, then applies
    // DWM backdrop automatically. Runs on a dedicated thread with a message pump.
    // =====================================================================
    public static class BackdropWatcher {
        // --- P/Invoke (watcher-specific) ---
        public delegate void WinEventDelegate(
            IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(
            uint eventMin, uint eventMax, IntPtr hmodWinEventProc,
            WinEventDelegate lpfnWinEventProc,
            uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll")]
        private static extern bool GetMessageW(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        private static extern bool TranslateMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessageW(ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern bool PostThreadMessage(uint idThread, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

        [StructLayout(LayoutKind.Sequential)]
        public struct MSG {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public int pt_x;
            public int pt_y;
        }

        // --- Constants ---
        private const uint EVENT_OBJECT_SHOW = 0x8002;
        private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
        private const uint WINEVENT_OUTOFCONTEXT = 0x0000;
        private const uint WINEVENT_SKIPOWNPROCESS = 0x0002;
        private const uint WM_QUIT = 0x0012;
        private const uint GA_ROOT = 2;

        // --- State ---
        private static Thread _thread;
        private static uint _threadId;
        private static volatile bool _running;
        private static int _backdropType = 2;
        private static bool _darkMode = false;
        private static bool _includeContextMenus = false;
        private static int _appliedCount = 0;
        private static HashSet<IntPtr> _appliedWindows = new HashSet<IntPtr>();

        // Pin the delegate so GC does not collect it while hook is active
        private static WinEventDelegate _callback;

        // System window classes to skip
        private static readonly HashSet<string> SystemClasses = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "Shell_TrayWnd", "Shell_SecondaryTrayWnd", "Progman", "WorkerW",
            "Windows.UI.Core.CoreWindow", "ForegroundStaging", "MultitaskingViewFrame",
            "XamlExplorerHostIslandWindow", "ConsoleWindowClass", "TaskManagerWindow",
            "ApplicationFrameInputSinkWindow", "Windows.Internal.Shell.TabProxyWindow",
            "EdgeUiInputTopWndClass", "EdgeUiInputWndClass",
            "Shell_InputSwitchTopLevelWindow", "LockScreenInputOcclusionWindow"
        };

        // Context menu classes
        private static readonly HashSet<string> MenuClasses = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "#32768", "Xaml_WindowedPopupClass"
        };

        // --- Public API ---
        public static bool IsRunning { get { return _running; } }
        public static int AppliedCount { get { return _appliedCount; } }

        public static void Start(int backdropType, bool darkMode, bool includeContextMenus) {
            if (_running) return;

            _backdropType = backdropType;
            _darkMode = darkMode;
            _includeContextMenus = includeContextMenus;
            _appliedCount = 0;
            _appliedWindows.Clear();
            _running = true;

            _thread = new Thread(WatcherThread);
            _thread.IsBackground = true;
            _thread.Name = "W11BackdropWatcher";
            _thread.Start();

            // Wait for thread to initialize
            for (int i = 0; i < 50 && _threadId == 0; i++) {
                Thread.Sleep(100);
            }
        }

        public static void Stop() {
            if (!_running) return;
            _running = false;

            if (_threadId != 0) {
                PostThreadMessage(_threadId, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            }

            if (_thread != null && _thread.IsAlive) {
                _thread.Join(5000);
            }

            _threadId = 0;
            _thread = null;
        }

        public static void UpdateSettings(int backdropType, bool darkMode, bool includeContextMenus) {
            _backdropType = backdropType;
            _darkMode = darkMode;
            _includeContextMenus = includeContextMenus;
        }

        // --- Internal ---
        private static void WatcherThread() {
            _threadId = GetCurrentThreadId();
            _callback = new WinEventDelegate(OnWinEvent);

            IntPtr hookShow = SetWinEventHook(
                EVENT_OBJECT_SHOW, EVENT_OBJECT_SHOW,
                IntPtr.Zero, _callback, 0, 0,
                WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);

            IntPtr hookFG = SetWinEventHook(
                EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
                IntPtr.Zero, _callback, 0, 0,
                WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);

            MSG msg;
            while (_running && GetMessageW(out msg, IntPtr.Zero, 0, 0)) {
                TranslateMessage(ref msg);
                DispatchMessageW(ref msg);
            }

            if (hookShow != IntPtr.Zero) UnhookWinEvent(hookShow);
            if (hookFG != IntPtr.Zero) UnhookWinEvent(hookFG);
            _callback = null;
        }

        private static void OnWinEvent(
            IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (idObject != 0 || hwnd == IntPtr.Zero) return;

            IntPtr rootHwnd = GetAncestor(hwnd, GA_ROOT);
            if (rootHwnd == IntPtr.Zero) rootHwnd = hwnd;

            if (_appliedWindows.Contains(rootHwnd)) return;
            if (!DwmHelper.IsWindowVisible(rootHwnd)) return;

            var sb = new System.Text.StringBuilder(256);
            DwmHelper.GetClassName(rootHwnd, sb, sb.Capacity);
            string className = sb.ToString();

            if (MenuClasses.Contains(className)) {
                if (_includeContextMenus) {
                    DwmHelper.SetBackdropType(rootHwnd, _backdropType);
                    if (_darkMode) DwmHelper.SetDarkMode(rootHwnd, true);
                }
                return;
            }

            if (SystemClasses.Contains(className)) return;

            int exStyle = DwmHelper.GetWindowLong(rootHwnd, -20);
            if ((exStyle & 0x80) != 0) return;
            if (DwmHelper.GetWindowTextLength(rootHwnd) == 0) return;

            int hr = DwmHelper.SetBackdropType(rootHwnd, _backdropType);
            if (hr == 0) {
                if (_darkMode) DwmHelper.SetDarkMode(rootHwnd, true);
                _appliedWindows.Add(rootHwnd);
                _appliedCount++;
            }
        }
    }
}
'@

Add-Type -TypeDefinition $dwmTypeDefinition -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# P/Invoke: OLD SWCA-based approach (undocumented, kept as fallback for
# the taskbar window specifically)
# ---------------------------------------------------------------------------
$swcaTypeDefinition = @'
using System;
using System.Runtime.InteropServices;

namespace W11ThemeSuite {
    [StructLayout(LayoutKind.Sequential)]
    public struct AccentPolicy {
        public int AccentState;
        public int AccentFlags;
        public uint GradientColor; // ABGR format
        public int AnimationId;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WindowCompositionAttributeData {
        public int Attribute;
        public IntPtr Data;
        public int SizeOfData;
    }

    public static class TaskbarTransparency {
        [DllImport("user32.dll")]
        public static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WindowCompositionAttributeData data);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        public static bool Apply(IntPtr hwnd, int accentState, uint gradientColor) {
            var accent = new AccentPolicy {
                AccentState = accentState,
                AccentFlags = 2,
                GradientColor = gradientColor,
                AnimationId = 0
            };

            var data = new WindowCompositionAttributeData {
                Attribute = 19,
                SizeOfData = Marshal.SizeOf(accent)
            };

            var accentPtr = Marshal.AllocHGlobal(data.SizeOfData);
            try {
                Marshal.StructureToPtr(accent, accentPtr, false);
                data.Data = accentPtr;
                return SetWindowCompositionAttribute(hwnd, ref data) != 0;
            }
            finally {
                Marshal.FreeHGlobal(accentPtr);
            }
        }
    }
}
'@

Add-Type -TypeDefinition $swcaTypeDefinition -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# Module-scoped constants
# ---------------------------------------------------------------------------
$script:RegistryBasePath = 'HKCU:\Software\w11-theming-suite\TaskbarTransparency'
$script:RunRegistryPath  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$script:RunValueName     = 'W11TaskbarTransparency'
$script:StartupDir       = Join-Path $env:LOCALAPPDATA 'w11-theming-suite\TaskbarTransparency'

# ---------------------------------------------------------------------------
# DWM backdrop type map (official API values)
# ---------------------------------------------------------------------------
$script:BackdropMap = @{
    'auto'    = 0  # DWMSBT_AUTO
    'none'    = 1  # DWMSBT_NONE
    'mica'    = 2  # DWMSBT_MAINWINDOW
    'acrylic' = 3  # DWMSBT_TRANSIENTWINDOW
    'tabbed'  = 4  # DWMSBT_TABBEDWINDOW (Mica Alt)
}

# ---------------------------------------------------------------------------
# Old SWCA style map (kept for backward-compat fallback)
# ---------------------------------------------------------------------------
$script:StyleMap = @{
    'clear'   = @{ AccentState = 2; DefaultColor = '00000000' }
    'blur'    = @{ AccentState = 3; DefaultColor = '00000000' }
    'acrylic' = @{ AccentState = 4; DefaultColor = 'CC000000' }
    'opaque'  = @{ AccentState = 1; DefaultColor = 'FF000000' }
    'normal'  = @{ AccentState = 0; DefaultColor = '00000000' }
}

# ---------------------------------------------------------------------------
# Mapping from old SWCA style names to new DWM backdrop types
# ---------------------------------------------------------------------------
$script:LegacyToDwmMap = @{
    'clear'   = 1  # DWMSBT_NONE
    'blur'    = 3  # DWMSBT_TRANSIENTWINDOW (acrylic is closest to blur)
    'acrylic' = 3  # DWMSBT_TRANSIENTWINDOW
    'opaque'  = 1  # DWMSBT_NONE
    'normal'  = 0  # DWMSBT_AUTO
}

# ===========================================================================
# Private Helper Functions
# ===========================================================================

function ConvertTo-COLORREF {
    <#
    .SYNOPSIS
        Converts an #RRGGBB or #AARRGGBB hex color string to a COLORREF uint32.
    .DESCRIPTION
        The Windows DwmSetWindowAttribute API expects colors in COLORREF format
        which is BGR byte order: 0x00BBGGRR. This function converts standard
        RGB hex notation to that format.

        Special string values 'none' and 'default' return the corresponding
        DWM sentinel values (DWMWA_COLOR_NONE and DWMWA_COLOR_DEFAULT).

        Examples:
          #FF0000 (red)   -> 0x000000FF
          #00FF00 (green) -> 0x0000FF00
          #0000FF (blue)  -> 0x00FF0000
          #000000 (black) -> 0x00000000
          'none'          -> 0xFFFFFFFE
          'default'       -> 0xFFFFFFFF
    .PARAMETER HexColor
        A hex color string (#RRGGBB or #AARRGGBB) or a special keyword
        ('none', 'default').
    .OUTPUTS
        System.UInt32 - The color in COLORREF (0x00BBGGRR) format.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$HexColor
    )

    # Handle special keyword values
    # NOTE: PowerShell 5.1 treats hex literals > 0x7FFFFFFF as signed int64/int32
    # which FAILS [uint32] cast. Use decimal literals or explicit conversion.
    if ($HexColor -eq 'none') {
        # DWMWA_COLOR_NONE = 0xFFFFFFFE = 4294967294 decimal
        return ([uint32]4294967294)
    }
    if ($HexColor -eq 'default') {
        # DWMWA_COLOR_DEFAULT = 0xFFFFFFFF = 4294967295 decimal
        return ([uint32]4294967295)
    }

    $hex = $HexColor.TrimStart('#')

    # Strip alpha channel if present (COLORREF does not use alpha)
    if ($hex.Length -eq 8) {
        $hex = $hex.Substring(2)  # drop AA prefix
    }

    if ($hex.Length -ne 6) {
        throw "Invalid color format '$HexColor'. Expected #RRGGBB or #AARRGGBB."
    }

    $r = [Convert]::ToByte($hex.Substring(0, 2), 16)
    $g = [Convert]::ToByte($hex.Substring(2, 2), 16)
    $b = [Convert]::ToByte($hex.Substring(4, 2), 16)

    # Pack as COLORREF: 0x00BBGGRR
    [uint32]$colorref = ([uint32]$b -shl 16) -bor ([uint32]$g -shl 8) -bor $r
    return $colorref
}

function ConvertTo-ABGRColor {
    <#
    .SYNOPSIS
        Converts an ARGB hex color string to an ABGR uint for the old SWCA API.
    .DESCRIPTION
        Kept for backward compatibility with Set-W11NativeTaskbarTransparency.
        Accepts colors in #AARRGGBB or #RRGGBB format.
    .PARAMETER HexColor
        Hex color string such as '#CC000000', 'FF336699', or '#336699'.
    .OUTPUTS
        System.UInt32 - The color in ABGR format.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$HexColor
    )

    $hex = $HexColor.TrimStart('#')

    if ($hex.Length -eq 6) {
        $hex = "FF$hex"
    }

    if ($hex.Length -ne 8) {
        throw "Invalid color format '$HexColor'. Expected #AARRGGBB or #RRGGBB."
    }

    $a = [Convert]::ToByte($hex.Substring(0, 2), 16)
    $r = [Convert]::ToByte($hex.Substring(2, 2), 16)
    $g = [Convert]::ToByte($hex.Substring(4, 2), 16)
    $b = [Convert]::ToByte($hex.Substring(6, 2), 16)

    [uint32]$abgr = ([uint32]$a -shl 24) -bor ([uint32]$b -shl 16) -bor ([uint32]$g -shl 8) -bor $r
    return $abgr
}

function Get-TopLevelWindows {
    <#
    .SYNOPSIS
        Returns visible top-level application windows, optionally filtering
        out shell/system windows.
    .DESCRIPTION
        Calls [W11ThemeSuite.DwmHelper]::GetVisibleWindows() and then
        optionally excludes known system window classes such as
        Shell_TrayWnd, Progman, WorkerW, etc.
    .PARAMETER ExcludeSystemWindows
        If set, filters out known Windows shell window classes that should
        not receive backdrop/color changes.
    .OUTPUTS
        System.IntPtr[] - Array of window handles.
    #>
    [CmdletBinding()]
    param(
        [switch]$ExcludeSystemWindows
    )

    $systemClasses = @(
        'Shell_TrayWnd',
        'Shell_SecondaryTrayWnd',
        'Progman',
        'WorkerW',
        'Windows.UI.Core.CoreWindow',
        'ForegroundStaging',
        'MultitaskingViewFrame',
        'XamlExplorerHostIslandWindow',
        'ConsoleWindowClass',     # Protected process -- returns INVALID_HANDLE
        'TaskManagerWindow'       # Protected process -- returns INVALID_HANDLE
    )

    $windows = [W11ThemeSuite.DwmHelper]::GetVisibleWindows()

    if ($ExcludeSystemWindows) {
        $filtered = [System.Collections.Generic.List[IntPtr]]::new()
        foreach ($hwnd in $windows) {
            $className = [W11ThemeSuite.DwmHelper]::GetWindowClassName($hwnd)
            if ($className -notin $systemClasses) {
                $filtered.Add($hwnd)
            }
        }
        return $filtered.ToArray()
    }

    return $windows.ToArray()
}

function Get-TaskbarHandles {
    <#
    .SYNOPSIS
        Finds the main and secondary taskbar window handles.
    .DESCRIPTION
        Locates the primary taskbar (Shell_TrayWnd) and all secondary monitor
        taskbars (Shell_SecondaryTrayWnd) using FindWindow / FindWindowEx.
    .PARAMETER IncludeSecondary
        If set, also returns handles for secondary-monitor taskbars.
    .OUTPUTS
        System.IntPtr[] - Array of taskbar window handles.
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludeSecondary
    )

    $handles = @()

    # Primary taskbar
    $primary = [W11ThemeSuite.DwmHelper]::FindWindow('Shell_TrayWnd', $null)
    if ($primary -ne [IntPtr]::Zero) {
        $handles += $primary
    }

    # Secondary taskbars (multi-monitor)
    if ($IncludeSecondary) {
        $child = [IntPtr]::Zero
        do {
            $child = [W11ThemeSuite.DwmHelper]::FindWindowEx(
                [IntPtr]::Zero, $child, 'Shell_SecondaryTrayWnd', $null
            )
            if ($child -ne [IntPtr]::Zero) {
                $handles += $child
            }
        } while ($child -ne [IntPtr]::Zero)
    }

    return $handles
}

function Save-TransparencyConfig {
    <#
    .SYNOPSIS
        Persists the current transparency configuration to the registry.
    #>
    [CmdletBinding()]
    param(
        [string]$Style,
        [string]$Color,
        [bool]$AllMonitors,
        [bool]$Enabled
    )

    if (-not (Test-Path $script:RegistryBasePath)) {
        New-Item -Path $script:RegistryBasePath -Force | Out-Null
    }

    Set-ItemProperty -Path $script:RegistryBasePath -Name 'Style'       -Value $Style
    Set-ItemProperty -Path $script:RegistryBasePath -Name 'Color'       -Value $Color
    Set-ItemProperty -Path $script:RegistryBasePath -Name 'AllMonitors' -Value ([int]$AllMonitors)
    Set-ItemProperty -Path $script:RegistryBasePath -Name 'Enabled'     -Value ([int]$Enabled)
}

# ===========================================================================
# Public Functions
# ===========================================================================

function Set-W11WindowBackdrop {
    <#
    .SYNOPSIS
        Applies a DWM system backdrop type to one or all visible windows.
    .DESCRIPTION
        Uses the official DwmSetWindowAttribute API with DWMWA_SYSTEMBACKDROP_TYPE
        (attribute 38) to set the backdrop material on application windows.

        This works on windows whose app framework supports DWM backdrops.
        Confirmed working: Notepad (WinUI3), Terminal (WinUI3), File Explorer,
        UWP apps. Apps that render their own frame (Chrome, Electron) will
        ignore this. Supported backdrop types:

          auto    - Let DWM decide the backdrop (system default)
          none    - No system backdrop (opaque)
          mica    - Mica material (subtle desktop tinting)
          acrylic - Desktop Acrylic (frosted glass blur)
          tabbed  - Mica Alt / Tabbed (variant of Mica)

        NOTE: The taskbar (Shell_TrayWnd) is a XAML Islands window on Win11
        22H2+ and may not respond to this API. Use Set-W11NativeTaskbarTransparency
        for best-effort taskbar transparency.
    .PARAMETER Style
        The backdrop type to apply. Valid values: auto, none, mica, acrylic, tabbed.
    .PARAMETER WindowHandle
        Apply the backdrop to a specific window handle (IntPtr). If omitted and
        -AllWindows is not specified, an error is shown.
    .PARAMETER AllWindows
        Enumerate all visible top-level application windows and apply the backdrop
        to each one. System windows (taskbar, desktop, etc.) are excluded.
    .PARAMETER DarkMode
        Also force immersive dark mode (dark title bars) on the affected windows
        via DWMWA_USE_IMMERSIVE_DARK_MODE.
    .EXAMPLE
        Set-W11WindowBackdrop -Style mica -AllWindows
        # Apply Mica backdrop to all visible application windows.
    .EXAMPLE
        Set-W11WindowBackdrop -Style acrylic -AllWindows -DarkMode
        # Apply Desktop Acrylic with dark title bars to all windows.
    .EXAMPLE
        $hwnd = [W11ThemeSuite.DwmHelper]::FindWindow('CabinetWClass', $null)
        Set-W11WindowBackdrop -Style tabbed -WindowHandle $hwnd
        # Apply Mica Alt to a specific Explorer window.
    .OUTPUTS
        System.Int32 - Number of windows successfully affected.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [ValidateSet('auto', 'none', 'mica', 'acrylic', 'tabbed')]
        [string]$Style = 'mica',

        [Parameter()]
        [IntPtr]$WindowHandle = [IntPtr]::Zero,

        [switch]$AllWindows,

        [switch]$DarkMode
    )

    try {
        $backdropType = $script:BackdropMap[$Style]

        # Determine target windows
        if ($AllWindows) {
            $targets = @(Get-TopLevelWindows -ExcludeSystemWindows)
        }
        elseif ($WindowHandle -ne [IntPtr]::Zero) {
            $targets = @($WindowHandle)
        }
        else {
            Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
            Write-Host 'Specify -WindowHandle or -AllWindows.'
            return 0
        }

        if ($targets.Count -eq 0) {
            Write-Host '[WARN]  ' -ForegroundColor Yellow -NoNewline
            Write-Host 'No eligible windows found.'
            return 0
        }

        $successCount = 0
        foreach ($hwnd in $targets) {
            $hr = [W11ThemeSuite.DwmHelper]::SetBackdropType($hwnd, $backdropType)
            if ($hr -eq 0) {
                $successCount++
            }
            else {
                $className = [W11ThemeSuite.DwmHelper]::GetWindowClassName($hwnd)
                Write-Verbose "DwmSetWindowAttribute failed on 0x$($hwnd.ToString('X')) ($className) HRESULT=0x$($hr.ToString('X8'))"
            }

            # Optionally set dark mode
            if ($DarkMode) {
                [W11ThemeSuite.DwmHelper]::SetDarkMode($hwnd, $true) | Out-Null
            }
        }

        Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
        Write-Host "Backdrop '$Style' applied to $successCount of $($targets.Count) window(s)."

        return $successCount
    }
    catch {
        Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
        Write-Host "Failed to set window backdrop: $_"
        return 0
    }
}

function Set-W11WindowColors {
    <#
    .SYNOPSIS
        Sets border, caption, and/or text colors on one or all visible windows.
    .DESCRIPTION
        Uses the official DwmSetWindowAttribute API with DWMWA_BORDER_COLOR (34),
        DWMWA_CAPTION_COLOR (35), and DWMWA_TEXT_COLOR (36) to customize window
        chrome colors on Windows 11 22H2+ (build 22621+).

        Colors are specified in standard #RRGGBB hex format. The API uses COLORREF
        (0x00BBGGRR) internally; this function handles the conversion.

        Special values:
          'none'    - Remove the border entirely (DWMWA_COLOR_NONE = 0xFFFFFFFE)
          'default' - Reset to system default color (DWMWA_COLOR_DEFAULT = 0xFFFFFFFF)

        NOTE: These attributes affect the window's non-client area (title bar and
        border). Confirmed working on Notepad, Terminal, File Explorer, and UWP
        apps. Chrome/Edge/Electron apps render their own frame and will not
        respond. Protected processes (ConsoleWindowClass, TaskManagerWindow)
        return INVALID_HANDLE. Taskbar/XAML Islands are also unaffected.
    .PARAMETER BorderColor
        Hex color for the window border (#RRGGBB), or 'none' to remove borders,
        or 'default' to restore system default.
    .PARAMETER CaptionColor
        Hex color for the title bar / caption area (#RRGGBB), or 'default'.
    .PARAMETER TextColor
        Hex color for the title bar text (#RRGGBB), or 'default'.
    .PARAMETER WindowHandle
        Apply colors to a specific window handle (IntPtr).
    .PARAMETER AllWindows
        Enumerate all visible top-level application windows and apply colors
        to each. System windows are excluded.
    .PARAMETER DarkMode
        Also force immersive dark mode on the affected windows.
    .EXAMPLE
        Set-W11WindowColors -BorderColor '#FF0000' -AllWindows
        # Set red borders on all application windows.
    .EXAMPLE
        Set-W11WindowColors -CaptionColor '#1A1A2E' -TextColor '#FFFFFF' -AllWindows -DarkMode
        # Dark caption with white text and dark mode title bars.
    .EXAMPLE
        Set-W11WindowColors -BorderColor 'none' -CaptionColor 'default' -AllWindows
        # Remove borders and reset caption to system default.
    .OUTPUTS
        System.Int32 - Number of windows successfully affected.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$BorderColor,

        [Parameter()]
        [string]$CaptionColor,

        [Parameter()]
        [string]$TextColor,

        [Parameter()]
        [IntPtr]$WindowHandle = [IntPtr]::Zero,

        [switch]$AllWindows,

        [switch]$DarkMode
    )

    try {
        # Validate that at least one color parameter was provided
        if (-not $BorderColor -and -not $CaptionColor -and -not $TextColor) {
            Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
            Write-Host 'Specify at least one of -BorderColor, -CaptionColor, or -TextColor.'
            return 0
        }

        # Pre-convert colors
        $borderRef  = if ($BorderColor)  { ConvertTo-COLORREF -HexColor $BorderColor }  else { $null }
        $captionRef = if ($CaptionColor) { ConvertTo-COLORREF -HexColor $CaptionColor } else { $null }
        $textRef    = if ($TextColor)    { ConvertTo-COLORREF -HexColor $TextColor }    else { $null }

        # Determine target windows
        if ($AllWindows) {
            $targets = @(Get-TopLevelWindows -ExcludeSystemWindows)
        }
        elseif ($WindowHandle -ne [IntPtr]::Zero) {
            $targets = @($WindowHandle)
        }
        else {
            Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
            Write-Host 'Specify -WindowHandle or -AllWindows.'
            return 0
        }

        if ($targets.Count -eq 0) {
            Write-Host '[WARN]  ' -ForegroundColor Yellow -NoNewline
            Write-Host 'No eligible windows found.'
            return 0
        }

        $successCount = 0
        foreach ($hwnd in $targets) {
            $ok = $true

            if ($null -ne $borderRef) {
                $hr = [W11ThemeSuite.DwmHelper]::SetBorderColor($hwnd, $borderRef)
                if ($hr -ne 0) {
                    $ok = $false
                    Write-Verbose "SetBorderColor failed on 0x$($hwnd.ToString('X')) HRESULT=0x$($hr.ToString('X8'))"
                }
            }

            if ($null -ne $captionRef) {
                $hr = [W11ThemeSuite.DwmHelper]::SetCaptionColor($hwnd, $captionRef)
                if ($hr -ne 0) {
                    $ok = $false
                    Write-Verbose "SetCaptionColor failed on 0x$($hwnd.ToString('X')) HRESULT=0x$($hr.ToString('X8'))"
                }
            }

            if ($null -ne $textRef) {
                $hr = [W11ThemeSuite.DwmHelper]::SetTextColor($hwnd, $textRef)
                if ($hr -ne 0) {
                    $ok = $false
                    Write-Verbose "SetTextColor failed on 0x$($hwnd.ToString('X')) HRESULT=0x$($hr.ToString('X8'))"
                }
            }

            if ($DarkMode) {
                [W11ThemeSuite.DwmHelper]::SetDarkMode($hwnd, $true) | Out-Null
            }

            if ($ok) { $successCount++ }
        }

        $parts = @()
        if ($BorderColor)  { $parts += "border=$BorderColor" }
        if ($CaptionColor) { $parts += "caption=$CaptionColor" }
        if ($TextColor)    { $parts += "text=$TextColor" }
        $colorSummary = $parts -join ', '

        Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
        Write-Host "Colors ($colorSummary) applied to $successCount of $($targets.Count) window(s)."

        return $successCount
    }
    catch {
        Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
        Write-Host "Failed to set window colors: $_"
        return 0
    }
}

function Set-W11NativeTaskbarTransparency {
    <#
    .SYNOPSIS
        Applies a transparency effect to the Windows 11 taskbar (backward-compatible).
    .DESCRIPTION
        This function is kept for backward compatibility with scripts that used the
        original SWCA-based approach. It now tries TWO approaches in order:

        1. DwmSetWindowAttribute with DWMWA_SYSTEMBACKDROP_TYPE (official API).
           Old style names are mapped to DWM backdrop types:
             clear   -> DWMSBT_NONE (1)
             blur    -> DWMSBT_TRANSIENTWINDOW (3) (acrylic is closest)
             acrylic -> DWMSBT_TRANSIENTWINDOW (3)
             opaque  -> DWMSBT_NONE (1)
             normal  -> DWMSBT_AUTO (0)

        2. SetWindowCompositionAttribute (undocumented SWCA API) as a fallback.
           This was the original approach but has ZERO visual effect on Windows 11
           22H2+ XAML taskbar, even though it returns success.

        IMPORTANT LIMITATION: Neither API may produce visible results on the
        Windows 11 22H2+ XAML taskbar. The taskbar is rendered through a XAML
        composition layer that does not respect these window attributes.
        For actual taskbar transparency, a TAP (Taskbar Appearance Plugin)
        injection approach like TranslucentTB is required.

        For ALL OTHER WINDOWS: Use Set-W11WindowBackdrop instead, which uses the
        official DwmSetWindowAttribute API and works perfectly.
    .PARAMETER Style
        The transparency style to apply:
          clear   - Fully transparent (DWM: none, SWCA: transparent gradient)
          blur    - Gaussian blur (DWM: acrylic, SWCA: blur behind)
          acrylic - Frosted-glass acrylic (DWM: acrylic, SWCA: acrylic blur)
          opaque  - Solid colored (DWM: none, SWCA: gradient)
          normal  - Reset to Windows default (DWM: auto, SWCA: disabled)
    .PARAMETER Color
        ARGB hex color for the SWCA gradient overlay (e.g. '#CC000000').
        Only used by the SWCA fallback path. If omitted, a sensible default
        is used for the chosen style.
    .PARAMETER AllMonitors
        Also apply the effect to taskbars on secondary monitors.
    .EXAMPLE
        Set-W11NativeTaskbarTransparency -Style clear
    .EXAMPLE
        Set-W11NativeTaskbarTransparency -Style acrylic -Color '#AA1A1A2E' -AllMonitors
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [ValidateSet('clear', 'blur', 'acrylic', 'opaque', 'normal')]
        [string]$Style = 'clear',

        [Parameter(Position = 1)]
        [ValidatePattern('^#?([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$')]
        [string]$Color,

        [switch]$AllMonitors
    )

    try {
        # Get taskbar handles
        $handles = @(Get-TaskbarHandles -IncludeSecondary:$AllMonitors)

        if ($handles.Count -eq 0) {
            Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
            Write-Host 'Could not find the taskbar window. Is Explorer running?'
            return
        }

        $successCount = 0
        $dwmBackdropType = $script:LegacyToDwmMap[$Style]

        foreach ($hwnd in $handles) {
            $applied = $false

            # --- Approach 1: DwmSetWindowAttribute (official API, try first) ---
            $hr = [W11ThemeSuite.DwmHelper]::SetBackdropType($hwnd, $dwmBackdropType)
            if ($hr -eq 0) {
                Write-Verbose "DWM backdrop type $dwmBackdropType applied to taskbar 0x$($hwnd.ToString('X'))"
                $applied = $true
            }
            else {
                Write-Verbose "DWM approach failed on taskbar 0x$($hwnd.ToString('X')) HRESULT=0x$($hr.ToString('X8')), trying SWCA fallback..."
            }

            # --- Approach 2: SWCA fallback (undocumented, may have no visual effect) ---
            $styleInfo   = $script:StyleMap[$Style]
            $accentState = $styleInfo.AccentState

            $effectiveColor = $Color
            if (-not $effectiveColor) {
                $effectiveColor = '#' + $styleInfo.DefaultColor
            }

            $gradientColor = ConvertTo-ABGRColor -HexColor $effectiveColor

            $swcaResult = [W11ThemeSuite.TaskbarTransparency]::Apply($hwnd, $accentState, $gradientColor)
            if ($swcaResult -and -not $applied) {
                Write-Verbose "SWCA fallback returned success on taskbar 0x$($hwnd.ToString('X')) (may have no visual effect on Win11 22H2+)"
                $applied = $true
            }

            if ($applied) {
                $successCount++
            }
            else {
                Write-Host '[WARN]  ' -ForegroundColor Yellow -NoNewline
                Write-Host "Both DWM and SWCA failed on taskbar handle 0x$($hwnd.ToString('X'))"
            }
        }

        if ($successCount -gt 0) {
            # Save state to registry
            $effectiveColor2 = $Color
            if (-not $effectiveColor2) {
                $effectiveColor2 = '#' + $script:StyleMap[$Style].DefaultColor
            }
            Save-TransparencyConfig -Style $Style -Color $effectiveColor2 `
                -AllMonitors $AllMonitors.IsPresent -Enabled $true

            Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
            Write-Host "Taskbar transparency set to '$Style' on $successCount taskbar(s)."
            Write-Host '        ' -NoNewline
            Write-Host '(Note: Visual effect on Win11 22H2+ XAML taskbar may be limited. Use Set-W11WindowBackdrop for app windows.)' -ForegroundColor DarkGray
        }
        else {
            Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
            Write-Host 'Failed to apply transparency to any taskbar.'
        }
    }
    catch {
        Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
        Write-Host "Failed to set taskbar transparency: $_"
    }
}

function Get-W11NativeTaskbarTransparency {
    <#
    .SYNOPSIS
        Returns the currently saved taskbar transparency configuration.
    .DESCRIPTION
        Reads the persisted configuration from the registry key
        HKCU:\Software\w11-theming-suite\TaskbarTransparency and returns it
        as a PSCustomObject with Style, Color, AllMonitors, and Enabled
        properties.
    .EXAMPLE
        Get-W11NativeTaskbarTransparency
    .EXAMPLE
        $config = Get-W11NativeTaskbarTransparency
        if ($config.Enabled) { Write-Host "Transparency is active: $($config.Style)" }
    .OUTPUTS
        PSCustomObject with properties: Style, Color, AllMonitors, Enabled
    #>
    [CmdletBinding()]
    param()

    try {
        if (-not (Test-Path $script:RegistryBasePath)) {
            Write-Host '[INFO]  ' -ForegroundColor Cyan -NoNewline
            Write-Host 'No saved configuration found. Taskbar transparency has not been configured.'
            return [PSCustomObject]@{
                Style       = 'normal'
                Color       = '#00000000'
                AllMonitors = $false
                Enabled     = $false
            }
        }

        $props = Get-ItemProperty -Path $script:RegistryBasePath -ErrorAction Stop

        return [PSCustomObject]@{
            Style       = if ($props.PSObject.Properties['Style'])       { $props.Style }       else { 'normal' }
            Color       = if ($props.PSObject.Properties['Color'])       { $props.Color }       else { '#00000000' }
            AllMonitors = if ($props.PSObject.Properties['AllMonitors']) { [bool]$props.AllMonitors } else { $false }
            Enabled     = if ($props.PSObject.Properties['Enabled'])     { [bool]$props.Enabled }     else { $false }
        }
    }
    catch {
        Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
        Write-Host "Failed to read configuration: $_"
    }
}

function Remove-W11NativeTaskbarTransparency {
    <#
    .SYNOPSIS
        Removes all window customizations and resets everything to system defaults.
    .DESCRIPTION
        Resets the taskbar to its default state using both DWM and SWCA approaches,
        resets all visible application windows to DWMSBT_AUTO backdrop type and
        default colors, removes the saved registry configuration, and unregisters
        any startup entries.

        This is the "undo everything" function.
    .EXAMPLE
        Remove-W11NativeTaskbarTransparency
    #>
    [CmdletBinding()]
    param()

    try {
        # Reset all taskbars to normal via both APIs
        $taskbarHandles = @(Get-TaskbarHandles -IncludeSecondary)
        foreach ($hwnd in $taskbarHandles) {
            # DWM: reset to auto
            [W11ThemeSuite.DwmHelper]::SetBackdropType($hwnd, 0) | Out-Null
            # SWCA: reset to disabled
            [W11ThemeSuite.TaskbarTransparency]::Apply($hwnd, 0, 0) | Out-Null
        }

        # Reset all application windows to defaults
        $appWindows = @(Get-TopLevelWindows -ExcludeSystemWindows)
        foreach ($hwnd in $appWindows) {
            [W11ThemeSuite.DwmHelper]::SetBackdropType($hwnd, 0) | Out-Null
            $defaultColor = [uint32]4294967295  # 0xFFFFFFFF - PS 5.1 safe
            [W11ThemeSuite.DwmHelper]::SetBorderColor($hwnd, $defaultColor)  | Out-Null
            [W11ThemeSuite.DwmHelper]::SetCaptionColor($hwnd, $defaultColor) | Out-Null
            [W11ThemeSuite.DwmHelper]::SetTextColor($hwnd, $defaultColor)    | Out-Null
        }

        # Remove saved configuration from registry
        if (Test-Path $script:RegistryBasePath) {
            Remove-Item -Path $script:RegistryBasePath -Recurse -Force -ErrorAction SilentlyContinue
        }

        # Clean up parent key if empty
        $parentPath = 'HKCU:\Software\w11-theming-suite'
        if (Test-Path $parentPath) {
            $children = Get-ChildItem -Path $parentPath -ErrorAction SilentlyContinue
            if ($null -eq $children -or $children.Count -eq 0) {
                Remove-Item -Path $parentPath -Force -ErrorAction SilentlyContinue
            }
        }

        # Remove startup registration
        Unregister-W11TaskbarTransparencyStartup

        $totalReset = $taskbarHandles.Count + $appWindows.Count
        Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
        Write-Host "All customizations removed. Reset $totalReset window(s) to system defaults."
    }
    catch {
        Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
        Write-Host "Failed to remove transparency: $_"
    }
}

function Register-W11TaskbarTransparencyStartup {
    <#
    .SYNOPSIS
        Registers taskbar transparency to apply automatically at user login.
    .DESCRIPTION
        Creates a VBScript wrapper (to suppress the PowerShell window flash)
        and a PowerShell script that re-applies the configured transparency
        style on login. The VBS is registered in the current user's Run key
        (HKCU:\Software\Microsoft\Windows\CurrentVersion\Run).

        The startup script imports this module and calls
        Set-W11NativeTaskbarTransparency with the saved parameters.
    .PARAMETER Style
        The transparency style to apply at startup.
    .PARAMETER Color
        ARGB hex color for the gradient overlay.
    .PARAMETER AllMonitors
        Whether to apply to secondary-monitor taskbars as well.
    .EXAMPLE
        Register-W11TaskbarTransparencyStartup -Style acrylic -Color '#CC000000' -AllMonitors
    .EXAMPLE
        Register-W11TaskbarTransparencyStartup -Style clear
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [ValidateSet('clear', 'blur', 'acrylic', 'opaque', 'normal')]
        [string]$Style = 'clear',

        [Parameter(Position = 1)]
        [ValidatePattern('^#?([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$')]
        [string]$Color,

        [switch]$AllMonitors
    )

    try {
        # Resolve default color if not supplied
        if (-not $Color) {
            $Color = '#' + $script:StyleMap[$Style].DefaultColor
        }

        # Ensure startup directory exists
        if (-not (Test-Path $script:StartupDir)) {
            New-Item -Path $script:StartupDir -ItemType Directory -Force | Out-Null
        }

        # Determine this module's path for import
        $modulePath = $PSScriptRoot

        # Build the PowerShell apply script
        $allMonitorsFlag = if ($AllMonitors) { ' -AllMonitors' } else { '' }

        $ps1Content = @"
# Auto-generated by w11-theming-suite NativeTaskbarTransparency
# This script is run at login via a VBS wrapper to apply taskbar transparency.

# Brief delay to ensure the taskbar is fully loaded after login
Start-Sleep -Seconds 3

Import-Module '$modulePath' -Force -ErrorAction Stop
Set-W11NativeTaskbarTransparency -Style '$Style' -Color '$Color'$allMonitorsFlag
"@

        $ps1Path = Join-Path $script:StartupDir 'apply-transparency.ps1'
        Set-Content -Path $ps1Path -Value $ps1Content -Encoding UTF8 -Force

        # Build the VBS wrapper to launch PowerShell hidden (no window flash)
        $vbsContent = @"
' Auto-generated by w11-theming-suite NativeTaskbarTransparency
' Launches the transparency apply script without a visible PowerShell window.
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & "$ps1Path" & """", 0, False
"@

        $vbsPath = Join-Path $script:StartupDir 'apply-transparency.vbs'
        Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII -Force

        # Register in HKCU Run key
        if (-not (Test-Path $script:RunRegistryPath)) {
            New-Item -Path $script:RunRegistryPath -Force | Out-Null
        }

        $wscriptCmd = "wscript.exe `"$vbsPath`""
        Set-ItemProperty -Path $script:RunRegistryPath -Name $script:RunValueName -Value $wscriptCmd

        # Persist configuration
        Save-TransparencyConfig -Style $Style -Color $Color `
            -AllMonitors $AllMonitors.IsPresent -Enabled $true

        Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
        Write-Host "Startup registration complete. Taskbar transparency ($Style) will be applied at login."
        Write-Host '        ' -NoNewline
        Write-Host "Scripts: $script:StartupDir" -ForegroundColor DarkGray
    }
    catch {
        Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
        Write-Host "Failed to register startup: $_"
    }
}

function Unregister-W11TaskbarTransparencyStartup {
    <#
    .SYNOPSIS
        Removes the login startup registration for taskbar transparency.
    .DESCRIPTION
        Deletes the VBS and PS1 scripts from the startup directory and
        removes the Run registry entry. Does not reset the current
        window state -- use Remove-W11NativeTaskbarTransparency for that.
    .EXAMPLE
        Unregister-W11TaskbarTransparencyStartup
    #>
    [CmdletBinding()]
    param()

    try {
        # Remove Run registry entry
        $runProps = Get-ItemProperty -Path $script:RunRegistryPath -ErrorAction SilentlyContinue
        if ($runProps -and $runProps.PSObject.Properties[$script:RunValueName]) {
            Remove-ItemProperty -Path $script:RunRegistryPath -Name $script:RunValueName -ErrorAction SilentlyContinue
        }

        # Remove startup scripts
        if (Test-Path $script:StartupDir) {
            $ps1Path = Join-Path $script:StartupDir 'apply-transparency.ps1'
            $vbsPath = Join-Path $script:StartupDir 'apply-transparency.vbs'

            if (Test-Path $ps1Path) { Remove-Item -Path $ps1Path -Force -ErrorAction SilentlyContinue }
            if (Test-Path $vbsPath) { Remove-Item -Path $vbsPath -Force -ErrorAction SilentlyContinue }

            # Remove directory if empty
            $remaining = Get-ChildItem -Path $script:StartupDir -ErrorAction SilentlyContinue
            if ($null -eq $remaining -or $remaining.Count -eq 0) {
                Remove-Item -Path $script:StartupDir -Force -ErrorAction SilentlyContinue
            }

            # Clean up parent if empty
            $parentDir = Split-Path $script:StartupDir -Parent
            if (Test-Path $parentDir) {
                $parentRemaining = Get-ChildItem -Path $parentDir -ErrorAction SilentlyContinue
                if ($null -eq $parentRemaining -or $parentRemaining.Count -eq 0) {
                    Remove-Item -Path $parentDir -Force -ErrorAction SilentlyContinue
                }
            }
        }

        Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
        Write-Host 'Startup registration removed.'
    }
    catch {
        Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
        Write-Host "Failed to unregister startup: $_"
    }
}

# ===========================================================================
# TAP (Taskbar Appearance Plugin) Injection
# ===========================================================================
# Uses InitializeXamlDiagnosticsEx to inject TaskbarTAP.dll into explorer.exe.
# This is the ONLY method that achieves true taskbar transparency on Win11 25H2,
# because the XAML taskbar continuously re-applies its own accent policy,
# overriding any external SWCA/DWM calls.
#
# The injected DLL implements IObjectWithSite + IVisualTreeServiceCallback2
# to find and modify Rectangle#BackgroundFill in the taskbar XAML tree.
# ===========================================================================

# P/Invoke for DLL injection via CreateRemoteThread + LoadLibraryW
# Two-stage approach (same as TranslucentTB):
#   Stage 1 (PowerShell): Inject TaskbarTAP.dll into explorer.exe via CreateRemoteThread
#   Stage 2 (inside DLL): DllMain spawns thread that calls InitializeXamlDiagnosticsEx
$tapTypeDefinition = @'
using System;
using System.Runtime.InteropServices;

namespace W11ThemeSuite {
    public static class TAPHelper {
        // Find explorer.exe process that owns the taskbar
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // Process manipulation for DLL injection
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, uint flAllocationType, uint flProtect);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, uint nSize, out uint lpNumberOfBytesWritten);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi)]
        public static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr CreateRemoteThread(IntPtr hProcess, IntPtr lpThreadAttributes, uint dwStackSize, IntPtr lpStartAddress, IntPtr lpParameter, uint dwCreationFlags, out uint lpThreadId);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, uint dwFreeType);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool CloseHandle(IntPtr hObject);

        // Constants
        public const uint PROCESS_ALL_ACCESS = 0x001F0FFF;
        public const uint MEM_COMMIT = 0x1000;
        public const uint MEM_RESERVE = 0x2000;
        public const uint MEM_RELEASE = 0x8000;
        public const uint PAGE_READWRITE = 0x04;
        public const uint INFINITE = 0xFFFFFFFF;
    }
}
'@

try {
    Add-Type -TypeDefinition $tapTypeDefinition -ErrorAction SilentlyContinue
} catch {
    # Type already loaded -- ignore
}

function Get-TaskbarExplorerPid {
    <#
    .SYNOPSIS
    Gets the PID of the explorer.exe process that owns the taskbar (Shell_TrayWnd).
    #>
    $hTaskbar = [W11ThemeSuite.TAPHelper]::FindWindow('Shell_TrayWnd', $null)
    if ($hTaskbar -eq [IntPtr]::Zero) {
        Write-Error "Taskbar window (Shell_TrayWnd) not found."
        return $null
    }

    $pid = [uint32]0
    [W11ThemeSuite.TAPHelper]::GetWindowThreadProcessId($hTaskbar, [ref]$pid) | Out-Null

    if ($pid -eq 0) {
        Write-Error "Could not get PID for Shell_TrayWnd."
        return $null
    }

    return $pid
}

function Invoke-TaskbarTAPInject {
    <#
    .SYNOPSIS
    Injects TaskbarTAP.dll into explorer.exe for taskbar transparency.

    .DESCRIPTION
    Two-stage injection (same approach as TranslucentTB):
      Stage 1: PowerShell injects TaskbarTAP.dll into explorer.exe via
               CreateRemoteThread + LoadLibraryW.
      Stage 2: The DLL's DllMain spawns a thread that calls
               InitializeXamlDiagnosticsEx from WITHIN explorer.exe,
               which registers the XAML visual tree watcher.

    This is the ONLY method that works on Windows 11 25H2 for true taskbar
    transparency, because the XAML taskbar continuously re-applies its own
    accent, overriding any SWCA or DWM-based approach.

    REQUIRES: Run as Administrator (for OpenProcess on explorer.exe).

    .PARAMETER Mode
    The appearance mode: 'Transparent', 'Acrylic', or 'Default'.

    .EXAMPLE
    Invoke-TaskbarTAPInject -Mode Transparent
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('Transparent', 'Acrylic', 'Default')]
        [string]$Mode = 'Transparent'
    )

    # Locate TaskbarTAP.dll
    $moduleRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $tapDll = Join-Path $moduleRoot 'native\bin\TaskbarTAP.dll'
    if (-not (Test-Path $tapDll)) {
        Write-Error "TaskbarTAP.dll not found at: $tapDll. Run native\TaskbarTAP\build.cmd first."
        return $false
    }

    $tapDllFull = (Resolve-Path $tapDll).Path
    Write-Verbose "TAP DLL: $tapDllFull"

    # Get the explorer.exe PID that owns the taskbar
    $explorerPid = Get-TaskbarExplorerPid
    if (-not $explorerPid) { return $false }
    Write-Verbose "Explorer PID: $explorerPid"

    Write-Host "Injecting TaskbarTAP.dll into explorer.exe (PID $explorerPid)..." -ForegroundColor Cyan

    # Stage 1: Inject DLL via CreateRemoteThread + LoadLibraryW
    # 1. Open the target process
    $hProcess = [W11ThemeSuite.TAPHelper]::OpenProcess(
        [W11ThemeSuite.TAPHelper]::PROCESS_ALL_ACCESS,
        $false,
        [uint32]$explorerPid
    )
    if ($hProcess -eq [IntPtr]::Zero) {
        $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Error "OpenProcess failed (error $err). Are you running as Administrator?"
        return $false
    }
    Write-Verbose "Opened explorer.exe process handle: $hProcess"

    try {
        # 2. Convert DLL path to UTF-16 bytes
        $dllPathBytes = [System.Text.Encoding]::Unicode.GetBytes($tapDllFull + "`0")
        $dllPathSize = [uint32]$dllPathBytes.Length

        # 3. Allocate memory in explorer.exe for the DLL path string
        $pRemoteMem = [W11ThemeSuite.TAPHelper]::VirtualAllocEx(
            $hProcess,
            [IntPtr]::Zero,
            $dllPathSize,
            ([W11ThemeSuite.TAPHelper]::MEM_COMMIT -bor [W11ThemeSuite.TAPHelper]::MEM_RESERVE),
            [W11ThemeSuite.TAPHelper]::PAGE_READWRITE
        )
        if ($pRemoteMem -eq [IntPtr]::Zero) {
            $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Error "VirtualAllocEx failed (error $err)"
            return $false
        }
        Write-Verbose "Allocated remote memory at: $pRemoteMem ($dllPathSize bytes)"

        # 4. Write the DLL path into explorer.exe's memory
        $bytesWritten = [uint32]0
        $wrote = [W11ThemeSuite.TAPHelper]::WriteProcessMemory(
            $hProcess,
            $pRemoteMem,
            $dllPathBytes,
            $dllPathSize,
            [ref]$bytesWritten
        )
        if (-not $wrote) {
            $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Error "WriteProcessMemory failed (error $err)"
            return $false
        }
        Write-Verbose "Wrote $bytesWritten bytes to remote process"

        # 5. Get address of LoadLibraryW in kernel32.dll
        $hKernel32 = [W11ThemeSuite.TAPHelper]::GetModuleHandle("kernel32.dll")
        if ($hKernel32 -eq [IntPtr]::Zero) {
            Write-Error "GetModuleHandle(kernel32.dll) failed"
            return $false
        }
        $pLoadLibraryW = [W11ThemeSuite.TAPHelper]::GetProcAddress($hKernel32, "LoadLibraryW")
        if ($pLoadLibraryW -eq [IntPtr]::Zero) {
            Write-Error "GetProcAddress(LoadLibraryW) failed"
            return $false
        }
        Write-Verbose "LoadLibraryW address: $pLoadLibraryW"

        # 6. Create a remote thread in explorer.exe that calls LoadLibraryW(dllPath)
        $threadId = [uint32]0
        $hThread = [W11ThemeSuite.TAPHelper]::CreateRemoteThread(
            $hProcess,
            [IntPtr]::Zero,
            0,
            $pLoadLibraryW,
            $pRemoteMem,
            0,
            [ref]$threadId
        )
        if ($hThread -eq [IntPtr]::Zero) {
            $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Error "CreateRemoteThread failed (error $err). Are you running as Administrator?"
            return $false
        }
        Write-Verbose "Created remote thread (TID $threadId), waiting for LoadLibrary..."

        # 7. Wait for the thread to finish (LoadLibrary + DllMain)
        [W11ThemeSuite.TAPHelper]::WaitForSingleObject($hThread, 10000) | Out-Null
        [W11ThemeSuite.TAPHelper]::CloseHandle($hThread) | Out-Null

        # 8. Free the remote memory (the DLL path string, no longer needed)
        [W11ThemeSuite.TAPHelper]::VirtualFreeEx(
            $hProcess,
            $pRemoteMem,
            0,
            [W11ThemeSuite.TAPHelper]::MEM_RELEASE
        ) | Out-Null

        Write-Host "DLL injected into explorer.exe!" -ForegroundColor Green
        Write-Host "Waiting for XAML Diagnostics initialization..." -ForegroundColor Gray

        # Stage 2 happens inside the DLL:
        # DllMain -> SelfInjectThread -> InitializeXamlDiagnosticsEx
        # This takes time (up to 30s with retries). Wait for shared memory to appear.
        $maxWait = 45  # seconds
        $waited = 0
        $sharedMemReady = $false

        while ($waited -lt $maxWait) {
            Start-Sleep -Milliseconds 1000
            $waited++

            try {
                $mmf = [System.IO.MemoryMappedFiles.MemoryMappedFile]::OpenExisting(
                    'W11ThemeSuite_TaskbarTAP_Mode')
                $mmf.Dispose()
                $sharedMemReady = $true
                break
            }
            catch {
                # Shared memory not yet created -- DLL is still initializing
                if ($waited % 5 -eq 0) {
                    Write-Verbose "Still waiting for TAP initialization... ($waited s)"
                }
            }
        }

        if (-not $sharedMemReady) {
            Write-Warning "TAP DLL was injected but shared memory not detected after ${maxWait}s."
            Write-Warning "The XAML Diagnostics initialization may have failed inside explorer."
            Write-Warning "Check that explorer.exe has XAML content (Win11 taskbar)."
            return $false
        }

        Write-Host "XAML Diagnostics initialized successfully!" -ForegroundColor Green

        # Set the desired mode via shared memory
        Set-TaskbarTAPMode -Mode $Mode

        Write-Host "Mode set to: $Mode" -ForegroundColor Green
        Write-Host "Use Set-TaskbarTAPMode to change appearance at runtime." -ForegroundColor Gray
        return $true
    }
    finally {
        [W11ThemeSuite.TAPHelper]::CloseHandle($hProcess) | Out-Null
    }
}

function Set-TaskbarTAPMode {
    <#
    .SYNOPSIS
    Changes the taskbar appearance mode via shared memory IPC with the injected TAP DLL.

    .DESCRIPTION
    Writes to a named shared memory region (W11ThemeSuite_TaskbarTAP_Mode) that
    the injected TaskbarTAP.dll monitors every 250ms. When a mode change is detected,
    the DLL updates the XAML visual tree accordingly.

    .PARAMETER Mode
    The appearance mode: 'Transparent' (0 opacity), 'Acrylic' (semi-transparent), or 'Default' (reset).

    .EXAMPLE
    Set-TaskbarTAPMode -Mode Transparent

    .EXAMPLE
    Set-TaskbarTAPMode -Mode Default   # Reset to normal taskbar
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Transparent', 'Acrylic', 'Default')]
        [string]$Mode
    )

    $modeMap = @{
        'Default'     = 0
        'Transparent' = 1
        'Acrylic'     = 2
    }
    $modeInt = $modeMap[$Mode]

    # Open the named shared memory (created by the DLL inside explorer.exe)
    $sharedMemName = 'W11ThemeSuite_TaskbarTAP_Mode'

    try {
        # Use .NET MemoryMappedFile to write the mode
        $mmf = [System.IO.MemoryMappedFiles.MemoryMappedFile]::OpenExisting($sharedMemName)
        $accessor = $mmf.CreateViewAccessor(0, 4)
        $accessor.Write(0, [int]$modeInt)
        $accessor.Dispose()
        $mmf.Dispose()
        Write-Verbose "TAP mode set to $Mode ($modeInt) via shared memory."
    }
    catch {
        Write-Error "Failed to set TAP mode. Is the TAP DLL injected? Error: $_"
        return $false
    }

    return $true
}

# ===========================================================================
# ShellTAP -- Generic XAML injection for any XAML-based process
# ===========================================================================
# Refactored from TaskbarTAP to support injection into:
#   - explorer.exe (taskbar)
#   - StartMenuExperienceHost.exe (Start Menu)
#   - ShellExperienceHost.exe (Action Center, Notifications)
#
# Uses the same two-stage injection pattern but with a configurable DLL
# that reads target element names from shared memory.
# ===========================================================================

function Invoke-ShellTAPInject {
    <#
    .SYNOPSIS
        Injects ShellTAP.dll into a target XAML process for transparency.
    .DESCRIPTION
        Generic XAML injection function. Writes a ShellTAPConfig struct to named
        shared memory, then injects ShellTAP.dll into the target process via
        CreateRemoteThread + LoadLibraryW.

        In discovery mode (no -TargetElements), the DLL logs ALL XAML elements
        to a discovery log file for analysis.

        In targeting mode, the DLL matches elements by name+type and applies
        the specified appearance mode.

        REQUIRES: Run as Administrator.
    .PARAMETER TargetProcess
        The process name to inject into (e.g., 'StartMenuExperienceHost').
    .PARAMETER TargetId
        Unique identifier for this injection target (e.g., 'StartMenu').
        Used to name shared memory regions.
    .PARAMETER TargetElements
        Array of "Name:Type" strings to match in the XAML tree.
        Example: @("BackgroundFill:Rectangle", "BackgroundStroke:Rectangle")
        If omitted or empty, enters discovery mode.
    .PARAMETER Mode
        The appearance mode: 'Transparent', 'Acrylic', or 'Default'.
    .PARAMETER LogPath
        Custom path for the discovery/debug log file.
    .EXAMPLE
        # Discovery mode: log all XAML elements in Start Menu
        Invoke-ShellTAPInject -TargetProcess StartMenuExperienceHost -TargetId StartMenu
    .EXAMPLE
        # Apply transparency to known taskbar elements via ShellTAP
        Invoke-ShellTAPInject -TargetProcess explorer -TargetId Taskbar -TargetElements @("BackgroundFill:Rectangle", "BackgroundStroke:Rectangle") -Mode Transparent
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TargetProcess,

        [Parameter(Mandatory)]
        [string]$TargetId,

        [Parameter()]
        [string[]]$TargetElements = @(),

        [Parameter()]
        [ValidateSet('Transparent', 'Acrylic', 'Default')]
        [string]$Mode = 'Transparent',

        [Parameter()]
        [string]$LogPath
    )

    # Locate ShellTAP.dll
    $moduleRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $shellTapDll = Join-Path $moduleRoot 'native\bin\ShellTAP.dll'
    if (-not (Test-Path $shellTapDll)) {
        Write-Error "ShellTAP.dll not found at: $shellTapDll. Run native\ShellTAP\build.cmd first."
        return $false
    }
    $shellTapDllFull = (Resolve-Path $shellTapDll).Path

    # Find the target process
    $proc = Get-Process -Name $TargetProcess -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $proc) {
        Write-Error "Process '$TargetProcess' not found. Is it running?"
        return $false
    }
    $targetPid = [uint32]$proc.Id
    Write-Verbose "Target process: $TargetProcess (PID $targetPid)"

    # Build the ShellTAPConfig struct and write to shared memory
    $modeMap = @{ 'Default' = 0; 'Transparent' = 1; 'Acrylic' = 2 }
    $modeInt = $modeMap[$Mode]

    # Calculate struct size: version(4) + mode(4) + targetCount(4) + targetNames(8*64*2) + targetTypes(8*128*2) + logPath(260*2) + flags(4)
    # = 4 + 4 + 4 + 1024 + 2048 + 520 + 4 = 3608 bytes
    $configSize = 3608
    $configName = "W11ThemeSuite_ShellTAP_${TargetId}_Config"

    try {
        # Create shared memory for config
        $mmfConfig = [System.IO.MemoryMappedFiles.MemoryMappedFile]::CreateOrOpen(
            $configName, $configSize,
            [System.IO.MemoryMappedFiles.MemoryMappedFileAccess]::ReadWrite)
        $accessor = $mmfConfig.CreateViewAccessor(0, $configSize)

        # Write version
        $accessor.Write(0, [int]1)
        # Write mode
        $accessor.Write(4, [int]$modeInt)
        # Write targetCount
        $accessor.Write(8, [int]$TargetElements.Count)

        # Write target names and types (offset 12)
        $nameOffset = 12
        $typeOffset = 12 + (8 * 64 * 2)  # After names block

        for ($i = 0; $i -lt [Math]::Min($TargetElements.Count, 8); $i++) {
            $parts = $TargetElements[$i] -split ':', 2
            $name = $parts[0]
            $type = if ($parts.Count -gt 1) { $parts[1] } else { '*' }

            # Write name (wchar_t[64])
            $nameBytes = [System.Text.Encoding]::Unicode.GetBytes($name)
            $namePos = $nameOffset + ($i * 64 * 2)
            for ($b = 0; $b -lt [Math]::Min($nameBytes.Length, 126); $b++) {
                $accessor.Write($namePos + $b, $nameBytes[$b])
            }

            # Write type (wchar_t[128])
            $typeBytes = [System.Text.Encoding]::Unicode.GetBytes($type)
            $typePos = $typeOffset + ($i * 128 * 2)
            for ($b = 0; $b -lt [Math]::Min($typeBytes.Length, 254); $b++) {
                $accessor.Write($typePos + $b, $typeBytes[$b])
            }
        }

        # Write logPath (offset after types block)
        $logPathOffset = $typeOffset + (8 * 128 * 2)
        if ($LogPath) {
            $logBytes = [System.Text.Encoding]::Unicode.GetBytes($LogPath)
            for ($b = 0; $b -lt [Math]::Min($logBytes.Length, 518); $b++) {
                $accessor.Write($logPathOffset + $b, $logBytes[$b])
            }
        }

        $accessor.Dispose()
        # Keep mmfConfig alive -- the DLL will read it

        Write-Verbose "Config written to shared memory '$configName'"
        if ($TargetElements.Count -eq 0) {
            Write-Host '[INFO]  ' -ForegroundColor Cyan -NoNewline
            Write-Host "Discovery mode: all XAML elements will be logged."
        }
    }
    catch {
        Write-Error "Failed to create config shared memory: $_"
        return $false
    }

    # Set environment variable for TargetId (DLL reads it on init)
    [System.Environment]::SetEnvironmentVariable('W11_SHELLTAP_TARGET', $TargetId, 'Process')

    Write-Host "Injecting ShellTAP.dll into $TargetProcess (PID $targetPid)..." -ForegroundColor Cyan

    # Stage 1: Inject DLL via CreateRemoteThread + LoadLibraryW
    $hProcess = [W11ThemeSuite.TAPHelper]::OpenProcess(
        [W11ThemeSuite.TAPHelper]::PROCESS_ALL_ACCESS, $false, $targetPid)
    if ($hProcess -eq [IntPtr]::Zero) {
        $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Error "OpenProcess failed (error $err). Are you running as Administrator?"
        return $false
    }

    try {
        $dllPathBytes = [System.Text.Encoding]::Unicode.GetBytes($shellTapDllFull + "`0")
        $dllPathSize = [uint32]$dllPathBytes.Length

        $pRemoteMem = [W11ThemeSuite.TAPHelper]::VirtualAllocEx(
            $hProcess, [IntPtr]::Zero, $dllPathSize,
            ([W11ThemeSuite.TAPHelper]::MEM_COMMIT -bor [W11ThemeSuite.TAPHelper]::MEM_RESERVE),
            [W11ThemeSuite.TAPHelper]::PAGE_READWRITE)
        if ($pRemoteMem -eq [IntPtr]::Zero) {
            Write-Error "VirtualAllocEx failed"
            return $false
        }

        $bytesWritten = [uint32]0
        $wrote = [W11ThemeSuite.TAPHelper]::WriteProcessMemory(
            $hProcess, $pRemoteMem, $dllPathBytes, $dllPathSize, [ref]$bytesWritten)
        if (-not $wrote) {
            Write-Error "WriteProcessMemory failed"
            return $false
        }

        $hKernel32 = [W11ThemeSuite.TAPHelper]::GetModuleHandle("kernel32.dll")
        $pLoadLibraryW = [W11ThemeSuite.TAPHelper]::GetProcAddress($hKernel32, "LoadLibraryW")

        $threadId = [uint32]0
        $hThread = [W11ThemeSuite.TAPHelper]::CreateRemoteThread(
            $hProcess, [IntPtr]::Zero, 0, $pLoadLibraryW, $pRemoteMem, 0, [ref]$threadId)
        if ($hThread -eq [IntPtr]::Zero) {
            $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Error "CreateRemoteThread failed (error $err). Are you running as Administrator?"
            return $false
        }

        [W11ThemeSuite.TAPHelper]::WaitForSingleObject($hThread, 10000) | Out-Null
        [W11ThemeSuite.TAPHelper]::CloseHandle($hThread) | Out-Null
        [W11ThemeSuite.TAPHelper]::VirtualFreeEx(
            $hProcess, $pRemoteMem, 0, [W11ThemeSuite.TAPHelper]::MEM_RELEASE) | Out-Null

        Write-Host "DLL injected!" -ForegroundColor Green
        Write-Host "Waiting for XAML Diagnostics initialization..." -ForegroundColor Gray

        # Wait for mode shared memory to appear
        $modeName = "W11ThemeSuite_ShellTAP_${TargetId}_Mode"
        $maxWait = 45
        $waited = 0
        $ready = $false

        while ($waited -lt $maxWait) {
            Start-Sleep -Milliseconds 1000
            $waited++
            try {
                $mmf = [System.IO.MemoryMappedFiles.MemoryMappedFile]::OpenExisting($modeName)
                $mmf.Dispose()
                $ready = $true
                break
            }
            catch {
                if ($waited % 10 -eq 0) {
                    Write-Verbose "Still waiting for ShellTAP initialization... ($waited s)"
                }
            }
        }

        if (-not $ready) {
            Write-Warning "ShellTAP DLL injected but shared memory not detected after ${maxWait}s."
            return $false
        }

        Write-Host "XAML Diagnostics initialized!" -ForegroundColor Green

        if ($TargetElements.Count -eq 0) {
            $logDir = Split-Path $shellTapDllFull -Parent
            Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
            Write-Host "Discovery mode active. Check log in: $logDir"
        }
        else {
            Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
            Write-Host "ShellTAP active on $TargetProcess (target=$TargetId, mode=$Mode)."
        }

        return $true
    }
    finally {
        [W11ThemeSuite.TAPHelper]::CloseHandle($hProcess) | Out-Null
    }
}

function Set-ShellTAPMode {
    <#
    .SYNOPSIS
        Changes the appearance mode for an active ShellTAP injection.
    .PARAMETER TargetId
        The TargetId used when injecting (e.g., 'StartMenu', 'Taskbar').
    .PARAMETER Mode
        The appearance mode: 'Transparent', 'Acrylic', or 'Default'.
    .EXAMPLE
        Set-ShellTAPMode -TargetId StartMenu -Mode Transparent
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TargetId,

        [Parameter(Mandatory)]
        [ValidateSet('Transparent', 'Acrylic', 'Default')]
        [string]$Mode
    )

    $modeMap = @{ 'Default' = 0; 'Transparent' = 1; 'Acrylic' = 2 }
    $modeInt = $modeMap[$Mode]
    $modeName = "W11ThemeSuite_ShellTAP_${TargetId}_Mode"

    try {
        $mmf = [System.IO.MemoryMappedFiles.MemoryMappedFile]::OpenExisting($modeName)
        $accessor = $mmf.CreateViewAccessor(0, 4)
        $accessor.Write(0, [int]$modeInt)
        $accessor.Dispose()
        $mmf.Dispose()
        Write-Verbose "ShellTAP mode set to $Mode ($modeInt) for target '$TargetId'."
    }
    catch {
        Write-Error "Failed to set ShellTAP mode for '$TargetId'. Is the DLL injected? Error: $_"
        return $false
    }
    return $true
}

# ===========================================================================
# Backdrop Watcher -- persistent DWM backdrop for ALL app windows
# ===========================================================================

function Start-W11BackdropWatcher {
    <#
    .SYNOPSIS
        Starts an event-driven watcher that automatically applies a DWM backdrop
        to all current and future application windows.
    .DESCRIPTION
        Uses SetWinEventHook to monitor EVENT_OBJECT_SHOW (new windows) and
        EVENT_SYSTEM_FOREGROUND (focus changes). When a new top-level app window
        appears, it applies the specified DWM backdrop type via the SetMica
        technique (ExtendFrame + SetBackdropType + remove caption color).

        The watcher runs on a dedicated background thread with a proper Win32
        message pump. Call Stop-W11BackdropWatcher to stop it.

        System windows (taskbar, desktop, etc.) are automatically excluded.
        Context menus can optionally be included with -IncludeContextMenus.
    .PARAMETER Style
        The backdrop type to apply. Valid values: mica, acrylic, tabbed.
    .PARAMETER DarkMode
        Also force immersive dark mode (dark title bars) on affected windows.
    .PARAMETER IncludeContextMenus
        Also apply the backdrop to context menus (#32768, Xaml_WindowedPopupClass).
    .PARAMETER ApplyToExisting
        Immediately apply the backdrop to all existing visible windows before
        starting the watcher. Default: $true.
    .EXAMPLE
        Start-W11BackdropWatcher -Style mica -DarkMode
        # Apply Mica backdrop to all windows, including future ones.
    .EXAMPLE
        Start-W11BackdropWatcher -Style acrylic -IncludeContextMenus
        # Apply Acrylic to all windows and context menus.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [ValidateSet('mica', 'acrylic', 'tabbed')]
        [string]$Style = 'mica',

        [switch]$DarkMode,

        [switch]$IncludeContextMenus,

        [bool]$ApplyToExisting = $true
    )

    if ([W11ThemeSuite.BackdropWatcher]::IsRunning) {
        Write-Host '[WARN]  ' -ForegroundColor Yellow -NoNewline
        Write-Host 'Backdrop watcher is already running. Use Stop-W11BackdropWatcher first.'
        return
    }

    $backdropType = $script:BackdropMap[$Style]

    # Apply to existing windows first
    if ($ApplyToExisting) {
        $count = Set-W11WindowBackdrop -Style $Style -AllWindows -DarkMode:$DarkMode
        Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
        Write-Host "Applied '$Style' backdrop to $count existing window(s)."
    }

    # Start the event-driven watcher
    [W11ThemeSuite.BackdropWatcher]::Start($backdropType, $DarkMode.IsPresent, $IncludeContextMenus.IsPresent)

    Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
    Write-Host "Backdrop watcher started (style=$Style, darkMode=$($DarkMode.IsPresent), contextMenus=$($IncludeContextMenus.IsPresent))."
    Write-Host '        ' -NoNewline
    Write-Host 'New windows will automatically receive the backdrop. Use Stop-W11BackdropWatcher to stop.' -ForegroundColor DarkGray
}

function Stop-W11BackdropWatcher {
    <#
    .SYNOPSIS
        Stops the persistent backdrop watcher.
    .DESCRIPTION
        Unhooks the SetWinEventHook callbacks and stops the background message
        pump thread. Already-applied backdrops remain on their windows until
        those windows are closed or the backdrop is reset.
    .EXAMPLE
        Stop-W11BackdropWatcher
    #>
    [CmdletBinding()]
    param()

    if (-not [W11ThemeSuite.BackdropWatcher]::IsRunning) {
        Write-Host '[WARN]  ' -ForegroundColor Yellow -NoNewline
        Write-Host 'Backdrop watcher is not running.'
        return
    }

    $count = [W11ThemeSuite.BackdropWatcher]::AppliedCount
    [W11ThemeSuite.BackdropWatcher]::Stop()

    Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
    Write-Host "Backdrop watcher stopped. Applied backdrop to $count window(s) during session."
}

function Register-W11BackdropWatcherStartup {
    <#
    .SYNOPSIS
        Registers the backdrop watcher to start automatically at user login.
    .DESCRIPTION
        Creates a PowerShell script and VBScript wrapper that launch the
        backdrop watcher at login. The VBS runs PowerShell hidden (no console
        flash). Registered in HKCU\...\Run.
    .PARAMETER Style
        The backdrop type to apply at startup.
    .PARAMETER DarkMode
        Also force dark mode on affected windows.
    .PARAMETER IncludeContextMenus
        Also apply to context menus.
    .EXAMPLE
        Register-W11BackdropWatcherStartup -Style mica -DarkMode
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [ValidateSet('mica', 'acrylic', 'tabbed')]
        [string]$Style = 'mica',

        [switch]$DarkMode,

        [switch]$IncludeContextMenus
    )

    try {
        $startupDir = Join-Path $env:LOCALAPPDATA 'w11-theming-suite\BackdropWatcher'
        if (-not (Test-Path $startupDir)) {
            New-Item -Path $startupDir -ItemType Directory -Force | Out-Null
        }

        $modulePath = $PSScriptRoot
        $darkFlag = if ($DarkMode) { ' -DarkMode' } else { '' }
        $menuFlag = if ($IncludeContextMenus) { ' -IncludeContextMenus' } else { '' }

        $ps1Content = @"
# Auto-generated by w11-theming-suite BackdropWatcher
# Applies persistent DWM backdrop to all app windows at login.

Start-Sleep -Seconds 5

Import-Module '$modulePath' -Force -ErrorAction Stop
Start-W11BackdropWatcher -Style '$Style'$darkFlag$menuFlag

# Keep the process alive so the watcher thread stays running
while (`$true) { Start-Sleep -Seconds 60 }
"@

        $ps1Path = Join-Path $startupDir 'apply-backdrop.ps1'
        Set-Content -Path $ps1Path -Value $ps1Content -Encoding UTF8 -Force

        $vbsContent = @"
' Auto-generated by w11-theming-suite BackdropWatcher
' Launches the backdrop watcher without a visible PowerShell window.
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & "$ps1Path" & """", 0, False
"@

        $vbsPath = Join-Path $startupDir 'apply-backdrop.vbs'
        Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII -Force

        $runPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
        if (-not (Test-Path $runPath)) {
            New-Item -Path $runPath -Force | Out-Null
        }

        $wscriptCmd = "wscript.exe `"$vbsPath`""
        Set-ItemProperty -Path $runPath -Name 'W11BackdropWatcher' -Value $wscriptCmd

        Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
        Write-Host "Startup registration complete. Backdrop watcher ($Style) will run at login."
        Write-Host '        ' -NoNewline
        Write-Host "Scripts: $startupDir" -ForegroundColor DarkGray
    }
    catch {
        Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
        Write-Host "Failed to register startup: $_"
    }
}

function Unregister-W11BackdropWatcherStartup {
    <#
    .SYNOPSIS
        Removes the backdrop watcher login startup registration.
    .EXAMPLE
        Unregister-W11BackdropWatcherStartup
    #>
    [CmdletBinding()]
    param()

    try {
        $runPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
        $runProps = Get-ItemProperty -Path $runPath -ErrorAction SilentlyContinue
        if ($runProps -and $runProps.PSObject.Properties['W11BackdropWatcher']) {
            Remove-ItemProperty -Path $runPath -Name 'W11BackdropWatcher' -ErrorAction SilentlyContinue
        }

        $startupDir = Join-Path $env:LOCALAPPDATA 'w11-theming-suite\BackdropWatcher'
        if (Test-Path $startupDir) {
            Remove-Item -Path $startupDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
        Write-Host 'Backdrop watcher startup registration removed.'
    }
    catch {
        Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
        Write-Host "Failed to unregister startup: $_"
    }
}

# ---------------------------------------------------------------------------
# Module exports
# ---------------------------------------------------------------------------
Export-ModuleMember -Function @(
    'Set-W11WindowBackdrop',
    'Set-W11WindowColors',
    'Set-W11NativeTaskbarTransparency',
    'Get-W11NativeTaskbarTransparency',
    'Remove-W11NativeTaskbarTransparency',
    'Register-W11TaskbarTransparencyStartup',
    'Unregister-W11TaskbarTransparencyStartup',
    'Invoke-TaskbarTAPInject',
    'Set-TaskbarTAPMode',
    'Get-TaskbarExplorerPid',
    'Invoke-ShellTAPInject',
    'Set-ShellTAPMode',
    'Start-W11BackdropWatcher',
    'Stop-W11BackdropWatcher',
    'Register-W11BackdropWatcherStartup',
    'Unregister-W11BackdropWatcherStartup'
)
