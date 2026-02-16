Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# P/Invoke types for SetWindowCompositionAttribute
# ---------------------------------------------------------------------------
$typeDefinition = @'
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

        // Accent states (from reverse-engineered ACCENT_STATE enum):
        // 0 = ACCENT_DISABLED (default/opaque)
        // 1 = ACCENT_ENABLE_GRADIENT
        // 2 = ACCENT_ENABLE_TRANSPARENTGRADIENT
        // 3 = ACCENT_ENABLE_BLURBEHIND (blur like Aero)
        // 4 = ACCENT_ENABLE_ACRYLICBLURBEHIND (acrylic - frosted glass, RS4 1803+)
        // 5 = ACCENT_ENABLE_HOSTBACKDROP (host backdrop, RS5 1809+)
        // 6 = ACCENT_INVALID_STATE
        //
        // For fully transparent (clear): AccentState=2, GradientColor=0x00000000
        // For acrylic dark: AccentState=4, GradientColor=0xCC000000
        // For blur: AccentState=3, GradientColor=0x00000000
        // For opaque colored: AccentState=1, GradientColor=0xFF{BBGGRR}

        public static bool Apply(IntPtr hwnd, int accentState, uint gradientColor) {
            var accent = new AccentPolicy {
                AccentState = accentState,
                AccentFlags = 2, // ACCENT_FLAG_ALLOW_SET_TRANSPARENCY
                GradientColor = gradientColor,
                AnimationId = 0
            };

            var data = new WindowCompositionAttributeData {
                Attribute = 19, // WCA_ACCENT_POLICY
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

Add-Type -TypeDefinition $typeDefinition -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# Module-scoped constants
# ---------------------------------------------------------------------------
$script:RegistryBasePath = 'HKCU:\Software\w11-theming-suite\TaskbarTransparency'
$script:RunRegistryPath  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$script:RunValueName     = 'W11TaskbarTransparency'
$script:StartupDir       = Join-Path $env:LOCALAPPDATA 'w11-theming-suite\TaskbarTransparency'

# ---------------------------------------------------------------------------
# Style-to-AccentState mapping
# ---------------------------------------------------------------------------
$script:StyleMap = @{
    'clear'   = @{ AccentState = 2; DefaultColor = '00000000' }
    'blur'    = @{ AccentState = 3; DefaultColor = '00000000' }
    'acrylic' = @{ AccentState = 4; DefaultColor = 'CC000000' }
    'opaque'  = @{ AccentState = 1; DefaultColor = 'FF000000' }
    'normal'  = @{ AccentState = 0; DefaultColor = '00000000' }
}

# ===========================================================================
# Private Helper Functions
# ===========================================================================

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
    $primary = [W11ThemeSuite.TaskbarTransparency]::FindWindow('Shell_TrayWnd', $null)
    if ($primary -ne [IntPtr]::Zero) {
        $handles += $primary
    }

    # Secondary taskbars (multi-monitor)
    if ($IncludeSecondary) {
        $child = [IntPtr]::Zero
        do {
            $child = [W11ThemeSuite.TaskbarTransparency]::FindWindowEx(
                [IntPtr]::Zero, $child, 'Shell_SecondaryTrayWnd', $null
            )
            if ($child -ne [IntPtr]::Zero) {
                $handles += $child
            }
        } while ($child -ne [IntPtr]::Zero)
    }

    return $handles
}

function ConvertTo-ABGRColor {
    <#
    .SYNOPSIS
        Converts an ARGB hex color string to a ABGR uint for the Windows API.
    .DESCRIPTION
        Accepts colors in #AARRGGBB or #RRGGBB format (the '#' prefix is optional).
        When only 6 hex digits are provided, alpha defaults to FF (fully opaque).
        The Windows SetWindowCompositionAttribute API expects ABGR byte order.
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
        # No alpha supplied; default to FF
        $hex = "FF$hex"
    }

    if ($hex.Length -ne 8) {
        throw "Invalid color format '$HexColor'. Expected #AARRGGBB or #RRGGBB."
    }

    $a = [Convert]::ToByte($hex.Substring(0, 2), 16)
    $r = [Convert]::ToByte($hex.Substring(2, 2), 16)
    $g = [Convert]::ToByte($hex.Substring(4, 2), 16)
    $b = [Convert]::ToByte($hex.Substring(6, 2), 16)

    # Pack as ABGR (little-endian uint32)
    [uint32]$abgr = ([uint32]$a -shl 24) -bor ([uint32]$b -shl 16) -bor ([uint32]$g -shl 8) -bor $r
    return $abgr
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

function Set-W11NativeTaskbarTransparency {
    <#
    .SYNOPSIS
        Applies a transparency effect to the Windows 11 taskbar natively.
    .DESCRIPTION
        Uses the undocumented SetWindowCompositionAttribute API to set the
        taskbar accent policy, achieving results identical to TranslucentTB
        without any third-party software.
    .PARAMETER Style
        The transparency style to apply:
          clear   - Fully transparent (see-through)
          blur    - Gaussian blur behind the taskbar
          acrylic - Frosted-glass acrylic effect
          opaque  - Solid colored taskbar
          normal  - Resets to the Windows default
    .PARAMETER Color
        ARGB hex color for the gradient overlay (e.g. '#CC000000' for
        semi-transparent black). If omitted, a sensible default is used
        for the chosen style.
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
        # Resolve style settings
        $styleInfo   = $script:StyleMap[$Style]
        $accentState = $styleInfo.AccentState

        if (-not $Color) {
            $Color = '#' + $styleInfo.DefaultColor
        }

        $gradientColor = ConvertTo-ABGRColor -HexColor $Color

        # Get taskbar handles (force array to ensure .Count works)
        $handles = @(Get-TaskbarHandles -IncludeSecondary:$AllMonitors)

        if ($handles.Count -eq 0) {
            Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
            Write-Host 'Could not find the taskbar window. Is Explorer running?'
            return
        }

        $successCount = 0
        foreach ($hwnd in $handles) {
            $result = [W11ThemeSuite.TaskbarTransparency]::Apply($hwnd, $accentState, $gradientColor)
            if ($result) {
                $successCount++
            }
            else {
                Write-Host '[WARN]  ' -ForegroundColor Yellow -NoNewline
                Write-Host "Failed to apply to handle 0x$($hwnd.ToString('X'))"
            }
        }

        if ($successCount -gt 0) {
            # Save state to registry
            Save-TransparencyConfig -Style $Style -Color $Color `
                -AllMonitors $AllMonitors.IsPresent -Enabled $true

            Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
            Write-Host "Taskbar transparency set to '$Style' on $successCount taskbar(s)."
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
        as a PSCustomObject.
    .EXAMPLE
        Get-W11NativeTaskbarTransparency
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
        Removes taskbar transparency and resets to the Windows default.
    .DESCRIPTION
        Applies AccentState 0 (disabled) to all taskbars, removes the
        saved registry configuration, and unregisters any startup entries.
    .EXAMPLE
        Remove-W11NativeTaskbarTransparency
    #>
    [CmdletBinding()]
    param()

    try {
        # Reset all taskbars to normal
        $handles = @(Get-TaskbarHandles -IncludeSecondary)
        foreach ($hwnd in $handles) {
            [W11ThemeSuite.TaskbarTransparency]::Apply($hwnd, 0, 0) | Out-Null
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

        Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
        Write-Host 'Taskbar transparency removed and reset to system default.'
    }
    catch {
        Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
        Write-Host "Failed to remove taskbar transparency: $_"
    }
}

function Register-W11TaskbarTransparencyStartup {
    <#
    .SYNOPSIS
        Registers taskbar transparency to apply automatically at user login.
    .DESCRIPTION
        Creates a VBScript wrapper (to suppress the PowerShell window flash)
        and a PowerShell script that re-applies the configured transparency
        style on login. The VBS is registered in the current user's Run key.
    .PARAMETER Style
        The transparency style to apply at startup.
    .PARAMETER Color
        ARGB hex color for the gradient overlay.
    .PARAMETER AllMonitors
        Whether to apply to secondary-monitor taskbars as well.
    .EXAMPLE
        Register-W11TaskbarTransparencyStartup -Style acrylic -Color '#CC000000' -AllMonitors
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
        taskbar state (use Remove-W11NativeTaskbarTransparency for that).
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

# ---------------------------------------------------------------------------
# Module exports
# ---------------------------------------------------------------------------
Export-ModuleMember -Function @(
    'Set-W11NativeTaskbarTransparency',
    'Get-W11NativeTaskbarTransparency',
    'Remove-W11NativeTaskbarTransparency',
    'Register-W11TaskbarTransparencyStartup',
    'Unregister-W11TaskbarTransparencyStartup'
)
