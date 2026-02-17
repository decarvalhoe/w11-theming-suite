#Requires -Version 5.1
# RegistryConfigurator.psm1
# Core module for w11-theming-suite registry-based theme customizations.
# Provides functions to read and write Windows 11 theme-related registry keys.

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Internal helper: Add-W11NativeMethods
# Loads P/Invoke signatures for Win32 APIs used to broadcast theme changes.
# ---------------------------------------------------------------------------
function Add-W11NativeMethods {
    [CmdletBinding()]
    param()

    # Only add the type if it has not been loaded yet in this session.
    try {
        [W11ThemeSuite.NativeMethods] | Out-Null
        Write-Verbose 'W11ThemeSuite.NativeMethods already loaded.'
        return
    }
    catch {
        # Type not loaded yet - proceed with Add-Type below.
    }

    $csharpCode = @'
using System;
using System.Runtime.InteropServices;

namespace W11ThemeSuite
{
    public static class NativeMethods
    {
        // --- Constants ---
        public const uint SPI_SETCURSORS        = 0x0057;
        public const uint SPI_SETDESKWALLPAPER   = 0x0014;
        public const uint SPIF_UPDATEINIFILE     = 0x01;
        public const uint SPIF_SENDCHANGE        = 0x02;
        public const int  HWND_BROADCAST         = 0xFFFF;
        public const uint WM_SETTINGCHANGE       = 0x001A;

        // --- P/Invoke: SystemParametersInfo ---
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool SystemParametersInfo(
            uint  uiAction,
            uint  uiParam,
            string pvParam,
            uint  fWinIni
        );

        // --- P/Invoke: SendMessageTimeout ---
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(
            IntPtr  hWnd,
            uint    Msg,
            UIntPtr wParam,
            string  lParam,
            uint    fuFlags,
            uint    uTimeout,
            out UIntPtr lpdwResult
        );
    }
}
'@

    Add-Type -TypeDefinition $csharpCode -Language CSharp -ErrorAction Stop
    Write-Verbose 'W11ThemeSuite.NativeMethods loaded successfully.'
}

# ---------------------------------------------------------------------------
# Internal helper: ConvertTo-ABGRDword
# Converts a hex RGB color string (e.g. "#6E00FF") to a DWORD in ABGR format.
# Windows DWM stores accent colors as 0xAABBGGRR.
# ---------------------------------------------------------------------------
function ConvertTo-ABGRDword {
    [CmdletBinding()]
    [OutputType([uint32])]
    param(
        [Parameter(Mandatory)]
        [ValidatePattern('^#?[0-9A-Fa-f]{6}$')]
        [string]$HexColor,

        [uint32]$Alpha = 0xFF
    )

    # Strip leading '#' if present.
    $hex = $HexColor.TrimStart('#')

    $R = [Convert]::ToByte($hex.Substring(0, 2), 16)
    $G = [Convert]::ToByte($hex.Substring(2, 2), 16)
    $B = [Convert]::ToByte($hex.Substring(4, 2), 16)

    # ABGR layout: Alpha in high byte, then Blue, Green, Red in low byte.
    [uint32]$dword = ($Alpha -shl 24) -bor ($B -shl 16) -bor ($G -shl 8) -bor $R
    return $dword
}

# ---------------------------------------------------------------------------
# Internal helper: ConvertFrom-ABGRDword
# Converts an ABGR DWORD back to its component values and a hex RGB string.
# ---------------------------------------------------------------------------
function ConvertFrom-ABGRDword {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [uint32]$Dword
    )

    $R = $Dword -band 0xFF
    $G = ($Dword -shr 8) -band 0xFF
    $B = ($Dword -shr 16) -band 0xFF
    $A = ($Dword -shr 24) -band 0xFF

    return @{
        R      = $R
        G      = $G
        B      = $B
        A      = $A
        HexRGB = '#{0:X2}{1:X2}{2:X2}' -f $R, $G, $B
    }
}

# ---------------------------------------------------------------------------
# Internal helper: Send-SettingsChangeNotification
# Broadcasts WM_SETTINGCHANGE so that all running applications pick up the
# registry modifications without requiring a logoff/restart.
# ---------------------------------------------------------------------------
function Send-SettingsChangeNotification {
    [CmdletBinding()]
    param()

    Add-W11NativeMethods

    $HWND_BROADCAST = [IntPtr][long][W11ThemeSuite.NativeMethods]::HWND_BROADCAST
    $result = [UIntPtr]::Zero

    [W11ThemeSuite.NativeMethods]::SendMessageTimeout(
        $HWND_BROADCAST,
        [W11ThemeSuite.NativeMethods]::WM_SETTINGCHANGE,
        [UIntPtr]::Zero,
        'ImmersiveColorSet',
        0x0002,   # SMTO_ABORTIFHUNG
        5000,     # timeout in ms
        [ref]$result
    ) | Out-Null

    Write-Verbose 'Broadcasted WM_SETTINGCHANGE (ImmersiveColorSet) to all top-level windows.'
}

# ---------------------------------------------------------------------------
# Internal helper: Set-RegistryValue
# Writes a single registry value, creating the key path if it does not exist.
# ---------------------------------------------------------------------------
function Set-RegistryValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)]$Value
    )

    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
        Write-Verbose "Created registry key: $Path"
    }

    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction Stop
    Write-Verbose "Set $Path\$Name = $Value ($Type)"
}

# ===========================================================================
#  Public function: Set-W11RegistryTheme
# ===========================================================================
function Set-W11RegistryTheme {
    <#
    .SYNOPSIS
        Applies a theme configuration to Windows 11 registry keys.

    .DESCRIPTION
        Reads a PSCustomObject config (typically loaded from a YAML/JSON theme
        file) and writes the corresponding registry values for dark mode,
        accent color, DWM composition, and taskbar layout.

    .PARAMETER Config
        A PSCustomObject containing theme sections: mode, accentColor, dwm,
        taskbar, and optionally advanced.registryOverrides.

    .PARAMETER Section
        Limit the operation to one or more sections (DarkMode, AccentColor,
        DWM, Taskbar). If omitted, all sections present in Config are applied.

    .PARAMETER RefreshExplorer
        If specified, restarts explorer.exe after applying changes so that
        the shell picks up the new theme immediately.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [Parameter()]
        [ValidateSet('DarkMode', 'AccentColor', 'DWM', 'Taskbar')]
        [string[]]$Section,

        [Parameter()]
        [switch]$RefreshExplorer
    )

    # Load the registry map from the companion data file.
    $mapPath = Join-Path $PSScriptRoot 'RegistryMap.psd1'
    $registryMap = Import-PowerShellDataFile -Path $mapPath

    # Determine which sections to process.
    $sectionsToProcess = if ($Section) { $Section } else { @('DarkMode', 'AccentColor', 'DWM', 'Taskbar') }

    # -----------------------------------------------------------------
    # DarkMode section
    # -----------------------------------------------------------------
    if ('DarkMode' -in $sectionsToProcess -and $null -ne $Config.mode) {
        Write-Verbose '--- Applying DarkMode section ---'
        $modeConfig = $Config.mode

        # Map config properties to registry keys:
        #   mode.appsUseLightTheme  -> AppsUseLightTheme  (0=dark, 1=light)
        #   mode.systemUsesLightTheme -> SystemUsesLightTheme (0=dark, 1=light)
        foreach ($entry in $registryMap.DarkMode.GetEnumerator()) {
            $reg = $entry.Value
            # Match the registry key name to the config property
            $configPropName = $entry.Key.Substring(0,1).ToLower() + $entry.Key.Substring(1)
            $configValue = $modeConfig.PSObject.Properties[$configPropName]
            if ($null -eq $configValue) {
                $configValue = $modeConfig.PSObject.Properties[$entry.Key]
            }
            if ($null -ne $configValue -and $null -ne $configValue.Value) {
                $val = [int]$configValue.Value
                if ($PSCmdlet.ShouldProcess("$($reg.Path)\$($reg.Name)", "Set to $val")) {
                    Set-RegistryValue -Path $reg.Path -Name $reg.Name -Type $reg.Type -Value $val
                }
            }
        }
    }

    # -----------------------------------------------------------------
    # AccentColor section
    # -----------------------------------------------------------------
    if ('AccentColor' -in $sectionsToProcess -and $null -ne $Config.accentColor) {
        Write-Verbose '--- Applying AccentColor section ---'
        $accentConfig = $Config.accentColor

        # Convert the hex color to an ABGR DWORD for registry storage.
        if ($accentConfig.color) {
            $alpha = if ($accentConfig.PSObject.Properties['alpha'] -and $null -ne $accentConfig.alpha) { [uint32]$accentConfig.alpha } else { 0xFF }
            $abgrDword = ConvertTo-ABGRDword -HexColor $accentConfig.color -Alpha $alpha

            # Set the DWM AccentColor key.
            $regAccent = $registryMap.AccentColor.AccentColor
            if ($PSCmdlet.ShouldProcess("$($regAccent.Path)\$($regAccent.Name)", "Set accent to 0x$($abgrDword.ToString('X8'))")) {
                Set-RegistryValue -Path $regAccent.Path -Name $regAccent.Name -Type $regAccent.Type -Value ([BitConverter]::ToInt32([BitConverter]::GetBytes($abgrDword), 0))
            }

            # Set the Explorer Accent keys to the same ABGR value.
            $regMenu = $registryMap.AccentColor.AccentColorMenu
            if ($PSCmdlet.ShouldProcess("$($regMenu.Path)\$($regMenu.Name)", "Set accent menu color")) {
                Set-RegistryValue -Path $regMenu.Path -Name $regMenu.Name -Type $regMenu.Type -Value ([BitConverter]::ToInt32([BitConverter]::GetBytes($abgrDword), 0))
            }

            $regStart = $registryMap.AccentColor.StartColorMenu
            if ($PSCmdlet.ShouldProcess("$($regStart.Path)\$($regStart.Name)", "Set start menu color")) {
                Set-RegistryValue -Path $regStart.Path -Name $regStart.Name -Type $regStart.Type -Value ([BitConverter]::ToInt32([BitConverter]::GetBytes($abgrDword), 0))
            }
        }

        # ColorPrevalence: set BOTH the Personalize and DWM keys to the same value.
        if ($accentConfig.PSObject.Properties['colorPrevalence'] -and $null -ne $accentConfig.colorPrevalence) {
            $prevalenceValue = [int]$accentConfig.colorPrevalence

            $regPers = $registryMap.AccentColor.ColorPrevalence_Personalize
            if ($PSCmdlet.ShouldProcess("$($regPers.Path)\$($regPers.Name)", "Set colorPrevalence=$prevalenceValue")) {
                Set-RegistryValue -Path $regPers.Path -Name $regPers.Name -Type $regPers.Type -Value $prevalenceValue
            }

            $regDwm = $registryMap.AccentColor.ColorPrevalence_DWM
            if ($PSCmdlet.ShouldProcess("$($regDwm.Path)\$($regDwm.Name)", "Set colorPrevalence=$prevalenceValue")) {
                Set-RegistryValue -Path $regDwm.Path -Name $regDwm.Name -Type $regDwm.Type -Value $prevalenceValue
            }
        }

        # EnableTransparency
        if ($accentConfig.PSObject.Properties['enableTransparency'] -and $null -ne $accentConfig.enableTransparency) {
            $regTrans = $registryMap.AccentColor.EnableTransparency
            if ($PSCmdlet.ShouldProcess("$($regTrans.Path)\$($regTrans.Name)", "Set enableTransparency")) {
                Set-RegistryValue -Path $regTrans.Path -Name $regTrans.Name -Type $regTrans.Type -Value ([int]$accentConfig.enableTransparency)
            }
        }

        # AutoColorization
        if ($accentConfig.PSObject.Properties['autoColorization'] -and $null -ne $accentConfig.autoColorization) {
            $regAuto = $registryMap.AccentColor.AutoColorization
            if ($PSCmdlet.ShouldProcess("$($regAuto.Path)\$($regAuto.Name)", "Set autoColorization")) {
                Set-RegistryValue -Path $regAuto.Path -Name $regAuto.Name -Type $regAuto.Type -Value ([int]$accentConfig.autoColorization)
            }
        }
    }

    # -----------------------------------------------------------------
    # DWM section
    # -----------------------------------------------------------------
    if ('DWM' -in $sectionsToProcess -and $null -ne $Config.dwm) {
        Write-Verbose '--- Applying DWM section ---'
        $dwmConfig = $Config.dwm

        foreach ($entry in $registryMap.DWM.GetEnumerator()) {
            $keyName = $entry.Key          # e.g. ColorizationColor
            $reg      = $entry.Value

            # Look up the matching property in the config (case-insensitive via PSObject).
            # Config keys use camelCase matching the registry value names.
            $configValue = $dwmConfig.PSObject.Properties[$keyName]
            if ($null -eq $configValue) {
                # Try lowercase-first variant (e.g. colorizationColor)
                $camel = $keyName.Substring(0,1).ToLower() + $keyName.Substring(1)
                $configValue = $dwmConfig.PSObject.Properties[$camel]
            }

            if ($null -ne $configValue) {
                $val = $configValue.Value

                # If the value looks like a hex color string, convert it to a DWORD.
                if ($val -is [string] -and $val -match '^(0x)?#?[0-9A-Fa-f]{6,8}$') {
                    $hexStr = $val -replace '^0x', '' -replace '^#', ''
                    $unsigned = [Convert]::ToUInt32($hexStr, 16)
                    $val = [BitConverter]::ToInt32([BitConverter]::GetBytes($unsigned), 0)
                }

                if ($PSCmdlet.ShouldProcess("$($reg.Path)\$($reg.Name)", "Set DWM value to $val")) {
                    Set-RegistryValue -Path $reg.Path -Name $reg.Name -Type $reg.Type -Value ([int]$val)
                }
            }
        }
    }

    # -----------------------------------------------------------------
    # Taskbar section
    # -----------------------------------------------------------------
    if ('Taskbar' -in $sectionsToProcess -and $null -ne $Config.taskbar) {
        Write-Verbose '--- Applying Taskbar section ---'
        $taskbarConfig = $Config.taskbar

        foreach ($entry in $registryMap.Taskbar.GetEnumerator()) {
            $keyName = $entry.Key
            $reg     = $entry.Value

            # Match config property by the registry value name (case-insensitive).
            $configValue = $taskbarConfig.PSObject.Properties[$keyName]
            if ($null -eq $configValue) {
                $camel = $keyName.Substring(0,1).ToLower() + $keyName.Substring(1)
                $configValue = $taskbarConfig.PSObject.Properties[$camel]
            }

            if ($null -ne $configValue) {
                $val = [int]$configValue.Value
                if ($PSCmdlet.ShouldProcess("$($reg.Path)\$($reg.Name)", "Set taskbar value to $val")) {
                    Set-RegistryValue -Path $reg.Path -Name $reg.Name -Type $reg.Type -Value $val
                }
            }
        }
    }

    # -----------------------------------------------------------------
    # Advanced: raw registry overrides
    # -----------------------------------------------------------------
    if ($null -ne $Config.advanced -and $null -ne $Config.advanced.registryOverrides) {
        Write-Verbose '--- Applying advanced registry overrides ---'

        foreach ($override in $Config.advanced.registryOverrides) {
            $oPath  = $override.path
            $oName  = $override.name
            $oValue = $override.value
            $oType  = if ($override.type) { $override.type } else { 'DWord' }

            if ($PSCmdlet.ShouldProcess("$oPath\$oName", "Set override to $oValue ($oType)")) {
                Set-RegistryValue -Path $oPath -Name $oName -Type $oType -Value $oValue
            }
        }
    }

    # -----------------------------------------------------------------
    # Broadcast change notification to all windows.
    # -----------------------------------------------------------------
    Write-Verbose 'Broadcasting settings change notification...'
    Send-SettingsChangeNotification

    # -----------------------------------------------------------------
    # Optionally restart explorer.exe for a full shell refresh.
    # -----------------------------------------------------------------
    if ($RefreshExplorer) {
        Write-Verbose 'Restarting explorer.exe...'
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
        Start-Process explorer
        Write-Verbose 'Explorer restarted.'
    }

    Write-Verbose 'Set-W11RegistryTheme completed.'
}

# ===========================================================================
#  Public function: Get-W11RegistryTheme
# ===========================================================================
function Get-W11RegistryTheme {
    <#
    .SYNOPSIS
        Reads the current Windows 11 theme-related registry values.

    .DESCRIPTION
        Queries all registry keys defined in RegistryMap.psd1 and returns
        a PSCustomObject whose properties mirror the map structure, with
        current values filled in.  Missing keys gracefully return $null.

    .PARAMETER Section
        Limit the read to one or more sections (DarkMode, AccentColor, DWM,
        Taskbar). If omitted, all sections are returned.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter()]
        [ValidateSet('DarkMode', 'AccentColor', 'DWM', 'Taskbar')]
        [string[]]$Section
    )

    # Load the registry map.
    $mapPath = Join-Path $PSScriptRoot 'RegistryMap.psd1'
    $registryMap = Import-PowerShellDataFile -Path $mapPath

    $sectionsToRead = if ($Section) { $Section } else { @('DarkMode', 'AccentColor', 'DWM', 'Taskbar') }

    $result = [ordered]@{}

    foreach ($sectionName in $sectionsToRead) {
        $sectionData = $registryMap[$sectionName]
        if ($null -eq $sectionData) {
            Write-Warning "Section '$sectionName' not found in RegistryMap."
            continue
        }

        $sectionResult = [ordered]@{}

        foreach ($entry in $sectionData.GetEnumerator()) {
            $keyName = $entry.Key
            $reg     = $entry.Value

            try {
                $value = Get-ItemPropertyValue -Path $reg.Path -Name $reg.Name -ErrorAction Stop
                $sectionResult[$keyName] = $value
            }
            catch {
                # Key does not exist or is inaccessible - return $null gracefully.
                Write-Verbose "Could not read $($reg.Path)\$($reg.Name): $_"
                $sectionResult[$keyName] = $null
            }
        }

        $result[$sectionName] = [PSCustomObject]$sectionResult
    }

    return [PSCustomObject]$result
}

# ---------------------------------------------------------------------------
# Export public functions
# ---------------------------------------------------------------------------
Export-ModuleMember -Function @(
    'Set-W11RegistryTheme',
    'Get-W11RegistryTheme'
)
