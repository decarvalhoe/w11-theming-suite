#Requires -Version 5.1
<#
.SYNOPSIS
    WallpaperManager module for the w11-theming-suite project.
.DESCRIPTION
    Manages Windows 11 desktop wallpaper settings using registry keys and
    P/Invoke calls to apply changes immediately without requiring a logoff.
#>

# ---------------------------------------------------------------------------
# Internal: Add-W11NativeMethods
# ---------------------------------------------------------------------------

function Add-W11NativeMethods {
    <#
    .SYNOPSIS
        Loads the W11ThemeSuite.NativeMethods type via Add-Type if not already present.
    #>
    if (-not ([System.Management.Automation.PSTypeName]'W11ThemeSuite.NativeMethods').Type) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace W11ThemeSuite {
    public class NativeMethods {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SystemParametersInfo(
            uint uiAction, uint uiParam, string pvParam, uint fWinIni);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(
            IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam,
            uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);

        public const uint SPI_SETDESKWALLPAPER = 0x0014;
        public const uint SPIF_UPDATEINIFILE   = 0x01;
        public const uint SPIF_SENDCHANGE      = 0x02;
    }
}
"@
    }
}

# ---------------------------------------------------------------------------
# Public: Set-W11Wallpaper
# ---------------------------------------------------------------------------

function Set-W11Wallpaper {
    <#
    .SYNOPSIS
        Sets the desktop wallpaper on Windows 11.
    .DESCRIPTION
        Applies a wallpaper by writing the appropriate registry values under
        HKCU:\Control Panel\Desktop and calling SystemParametersInfo to
        refresh the desktop immediately.

        Accepts either a Config object (from a theme JSON) or explicit
        WallpaperPath / Style parameters.
    .PARAMETER Config
        A PSCustomObject containing a .wallpaper property with .path, .style,
        and optionally .mode / .images for slideshow support.
    .PARAMETER WallpaperPath
        Direct path to a wallpaper image file. If relative (no drive letter),
        the path is resolved relative to <ProjectRoot>\assets\wallpapers\.
    .PARAMETER Style
        Wallpaper fit style as an integer:
          0  = Center
          2  = Stretch
          6  = Fit
          10 = Fill (default)
    .EXAMPLE
        Set-W11Wallpaper -WallpaperPath 'my-wallpaper.jpg'
    .EXAMPLE
        Set-W11Wallpaper -Config $themeConfig
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [PSCustomObject]$Config,

        [Parameter()]
        [string]$WallpaperPath,

        [Parameter()]
        [int]$Style = 10
    )

    # Ensure native methods are available
    Add-W11NativeMethods

    # ------------------------------------------------------------------
    # Resolve parameters from Config if provided
    # ------------------------------------------------------------------
    $mode = 'static'

    if ($Config) {
        $wp = $Config.wallpaper
        if (-not $wp) {
            Write-Error 'Config object does not contain a .wallpaper property.'
            return
        }

        if ($wp.mode) { $mode = $wp.mode }

        # Use config path unless an explicit WallpaperPath was also given
        if (-not $WallpaperPath -and $wp.path) {
            $WallpaperPath = $wp.path
        }

        if ($wp.style -ne $null) {
            $Style = [int]$wp.style
        }
    }

    # ------------------------------------------------------------------
    # Determine the project root (two levels up from this module)
    # ------------------------------------------------------------------
    $ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\')).Path

    # ------------------------------------------------------------------
    # Handle slideshow mode
    # ------------------------------------------------------------------
    if ($mode -eq 'slideshow') {
        Write-Warning ('Slideshow mode requires applying a .theme file for full support. ' +
                        'Setting the first image as a static wallpaper fallback.')

        # Attempt to use the first image from config as a static fallback
        if ($Config.wallpaper.images -and $Config.wallpaper.images.Count -gt 0) {
            $WallpaperPath = $Config.wallpaper.images[0]
        }

        if (-not $WallpaperPath) {
            Write-Error 'No images found in slideshow configuration.'
            return
        }

        # Fall through to static application below with the first image
    }

    # ------------------------------------------------------------------
    # Validate wallpaper path
    # ------------------------------------------------------------------
    if (-not $WallpaperPath) {
        Write-Error 'No wallpaper path specified. Provide -WallpaperPath or a Config with .wallpaper.path.'
        return
    }

    # Resolve relative paths against the project assets folder
    if ($WallpaperPath -notmatch '^[A-Za-z]:\\') {
        $WallpaperPath = Join-Path $ProjectRoot "assets\wallpapers\$WallpaperPath"
    }

    $ResolvedPath = $WallpaperPath
    if (-not (Test-Path -LiteralPath $ResolvedPath)) {
        Write-Error "Wallpaper file not found: $ResolvedPath"
        return
    }

    # Normalise to a fully-qualified path
    $ResolvedPath = (Resolve-Path -LiteralPath $ResolvedPath).Path

    # ------------------------------------------------------------------
    # Determine TileWallpaper value
    # ------------------------------------------------------------------
    # TileWallpaper = 1 only when the legacy "Tile" style is desired
    # (WallpaperStyle 0 + TileWallpaper 1). For all modern styles, set to 0.
    $TileValue = '0'

    # ------------------------------------------------------------------
    # Write registry values
    # ------------------------------------------------------------------
    $regPath = 'HKCU:\Control Panel\Desktop'

    try {
        Set-ItemProperty -Path $regPath -Name 'Wallpaper'      -Value $ResolvedPath -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name 'WallpaperStyle'  -Value $Style.ToString() -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name 'TileWallpaper'   -Value $TileValue -ErrorAction Stop

        Write-Verbose "Registry updated: Wallpaper=$ResolvedPath, Style=$Style, Tile=$TileValue"
    }
    catch {
        Write-Error "Failed to update wallpaper registry values: $_"
        return
    }

    # ------------------------------------------------------------------
    # Apply immediately via SystemParametersInfo
    # ------------------------------------------------------------------
    $SPI  = [W11ThemeSuite.NativeMethods]::SPI_SETDESKWALLPAPER
    $flags = [W11ThemeSuite.NativeMethods]::SPIF_UPDATEINIFILE -bor `
             [W11ThemeSuite.NativeMethods]::SPIF_SENDCHANGE

    $result = [W11ThemeSuite.NativeMethods]::SystemParametersInfo($SPI, 0, $ResolvedPath, $flags)

    if (-not $result) {
        $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Error "SystemParametersInfo failed with Win32 error code $err."
        return
    }

    # ------------------------------------------------------------------
    # Output confirmation
    # ------------------------------------------------------------------
    $styleLabel = switch ($Style) {
        0  { 'Center' }
        2  { 'Stretch' }
        6  { 'Fit' }
        10 { 'Fill' }
        default { "Unknown ($Style)" }
    }

    Write-Output "Wallpaper applied: $ResolvedPath (Style: $styleLabel)"

    if ($mode -eq 'slideshow') {
        Write-Output 'Note: Only the first slideshow image was applied as a static fallback. Apply the generated .theme file for full slideshow rotation.'
    }
}

# ---------------------------------------------------------------------------
# Public: Get-W11Wallpaper
# ---------------------------------------------------------------------------

function Get-W11Wallpaper {
    <#
    .SYNOPSIS
        Retrieves the current desktop wallpaper settings.
    .DESCRIPTION
        Reads the wallpaper path, style, and tile values from the registry
        and returns them as a PSCustomObject.
    .EXAMPLE
        Get-W11Wallpaper
    #>
    [CmdletBinding()]
    param()

    $regPath = 'HKCU:\Control Panel\Desktop'

    try {
        $regValues = Get-ItemProperty -Path $regPath -ErrorAction Stop

        [PSCustomObject]@{
            Path          = $regValues.Wallpaper
            Style         = $regValues.WallpaperStyle
            TileWallpaper = $regValues.TileWallpaper
        }
    }
    catch {
        Write-Error "Failed to read wallpaper settings from the registry: $_"
    }
}
