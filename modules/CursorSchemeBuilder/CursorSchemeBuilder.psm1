#Requires -Version 5.1
<#
.SYNOPSIS
    CursorSchemeBuilder module for w11-theming-suite.

.DESCRIPTION
    Manages custom Windows 11 cursor schemes natively via the registry.
    Supports installing, activating, uninstalling, and enumerating cursor schemes.

    Registry locations:
      - Active cursors:  HKCU:\Control Panel\Cursors        (each role = REG_EXPAND_SZ)
      - Saved schemes:   HKCU:\Control Panel\Cursors\Schemes (scheme = comma-separated paths)
      - Active scheme:   HKCU:\Control Panel\Cursors -> (Default) value
#>

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# The 15 standard Windows cursor roles, in canonical registry order.
$script:CursorRoles = @(
    'Arrow',
    'Help',
    'AppStarting',
    'Wait',
    'Crosshair',
    'IBeam',
    'NWPen',
    'No',
    'SizeNS',
    'SizeWE',
    'SizeNWSE',
    'SizeNESW',
    'SizeAll',
    'UpArrow',
    'Hand'
)

$script:CursorsRegPath  = 'HKCU:\Control Panel\Cursors'
$script:SchemesRegPath  = 'HKCU:\Control Panel\Cursors\Schemes'

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

function Add-W11NativeMethods {
    <#
    .SYNOPSIS
        Loads P/Invoke definition for SystemParametersInfo if not already present.
    .DESCRIPTION
        Adds the [W11ThemeSuite.NativeMethods] type via Add-Type so that
        SystemParametersInfo can be called to refresh the active cursor set.
    #>
    if (-not ([System.Management.Automation.PSTypeName]'W11ThemeSuite.NativeMethods').Type) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace W11ThemeSuite {
    public class NativeMethods {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SystemParametersInfo(
            uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

        public const uint SPI_SETCURSORS      = 0x0057;
        public const uint SPIF_UPDATEINIFILE   = 0x01;
        public const uint SPIF_SENDCHANGE      = 0x02;
    }
}
"@
        Write-Verbose 'Loaded W11ThemeSuite.NativeMethods P/Invoke definitions.'
    }
}

# ---------------------------------------------------------------------------
# Public functions
# ---------------------------------------------------------------------------

function Install-W11CursorScheme {
    <#
    .SYNOPSIS
        Installs a custom cursor scheme from the project assets into the user profile
        and registers it in the Windows cursor scheme registry.

    .DESCRIPTION
        Copies .cur / .ani files from the project's assets\cursors\<setFolder> directory
        to a per-user location under %LOCALAPPDATA%\w11-theming-suite\cursors\<schemeName>,
        then writes the scheme entry to HKCU:\Control Panel\Cursors\Schemes.

        When -Activate is specified the scheme is immediately applied.

    .PARAMETER Config
        A PSCustomObject containing at minimum a .cursors property with:
          - schemeName  : Display name for the scheme.
          - setFolder   : Sub-folder name under assets\cursors\.
          - roles       : Hashtable mapping role names to filenames (e.g. Arrow = 'arrow.cur').

    .PARAMETER Activate
        If present, the scheme is applied immediately after installation.

    .EXAMPLE
        Install-W11CursorScheme -Config $themeConfig -Activate
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [switch]$Activate
    )

    # ---- Validate configuration ----
    if (-not $Config.cursors) {
        throw 'Config object does not contain a "cursors" property.'
    }
    $cursors = $Config.cursors

    if (-not $cursors.schemeName) {
        throw 'Config.cursors.schemeName is required.'
    }
    if (-not $cursors.setFolder) {
        throw 'Config.cursors.setFolder is required.'
    }

    $schemeName = $cursors.schemeName
    Write-Verbose "Installing cursor scheme: $schemeName"

    # ---- Resolve paths ----
    # Project root is two levels above this module's directory.
    $projectRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
    $sourceFolder = Join-Path $projectRoot "assets\cursors\$($cursors.setFolder)"

    if (-not (Test-Path $sourceFolder)) {
        throw "Source cursor folder not found: $sourceFolder"
    }

    $destFolder = Join-Path $env:LOCALAPPDATA "w11-theming-suite\cursors\$schemeName"

    # ---- Copy cursor files ----
    if (-not (Test-Path $destFolder)) {
        New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
        Write-Verbose "Created destination directory: $destFolder"
    }

    $cursorFiles = Get-ChildItem -Path $sourceFolder -Include '*.cur', '*.ani' -File -Recurse
    if ($cursorFiles.Count -eq 0) {
        Write-Warning "No .cur or .ani files found in $sourceFolder"
    }
    else {
        Copy-Item -Path $cursorFiles.FullName -Destination $destFolder -Force
        Write-Verbose "Copied $($cursorFiles.Count) cursor file(s) to $destFolder"
    }

    # ---- Build the scheme registry string ----
    # The scheme value is a comma-separated list of cursor paths for all 15 roles
    # in the canonical order. Missing roles get an empty string.
    $roleMap = @{}
    if ($cursors.roles) {
        # Normalise keys coming from JSON / PSCustomObject into a hashtable.
        if ($cursors.roles -is [hashtable]) {
            $roleMap = $cursors.roles
        }
        else {
            $cursors.roles.PSObject.Properties | ForEach-Object {
                $roleMap[$_.Name] = $_.Value
            }
        }
    }

    $schemeParts = foreach ($role in $script:CursorRoles) {
        if ($roleMap.ContainsKey($role) -and $roleMap[$role]) {
            Join-Path $destFolder $roleMap[$role]
        }
        else {
            ''   # empty position for undefined roles
        }
    }
    $schemeString = $schemeParts -join ','

    # ---- Register the scheme ----
    if (-not (Test-Path $script:SchemesRegPath)) {
        New-Item -Path $script:SchemesRegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $script:SchemesRegPath -Name $schemeName -Value $schemeString -Type ExpandString
    Write-Verbose "Registered scheme '$schemeName' in $($script:SchemesRegPath)"

    # ---- Activate if requested ----
    if ($Activate) {
        Write-Verbose 'Activating cursor scheme...'

        # Set the active scheme display name.
        Set-ItemProperty -Path $script:CursorsRegPath -Name '(Default)' -Value $schemeName

        # Write each individual cursor role to the active-cursors key.
        foreach ($role in $script:CursorRoles) {
            if ($roleMap.ContainsKey($role) -and $roleMap[$role]) {
                $cursorPath = Join-Path $destFolder $roleMap[$role]
                Set-ItemProperty -Path $script:CursorsRegPath -Name $role -Value $cursorPath -Type ExpandString
            }
        }

        # Notify the system to reload cursors.
        Add-W11NativeMethods
        $flags = [W11ThemeSuite.NativeMethods]::SPIF_UPDATEINIFILE -bor `
                 [W11ThemeSuite.NativeMethods]::SPIF_SENDCHANGE

        $result = [W11ThemeSuite.NativeMethods]::SystemParametersInfo(
            [W11ThemeSuite.NativeMethods]::SPI_SETCURSORS,
            0,
            [IntPtr]::Zero,
            $flags
        )

        if (-not $result) {
            Write-Warning 'SystemParametersInfo(SPI_SETCURSORS) returned false. Cursors may not have refreshed.'
        }
        else {
            Write-Verbose 'System notified to reload cursors.'
        }
    }

    # ---- Summary ----
    [PSCustomObject]@{
        SchemeName  = $schemeName
        Source      = $sourceFolder
        Destination = $destFolder
        FilesCopied = $cursorFiles.Count
        Activated   = [bool]$Activate
    }
}

function Uninstall-W11CursorScheme {
    <#
    .SYNOPSIS
        Removes a previously installed cursor scheme.

    .DESCRIPTION
        Deletes the scheme entry from the registry, removes the copied cursor files,
        and reverts to Windows defaults if the removed scheme was currently active.

    .PARAMETER SchemeName
        The display name of the scheme to remove.

    .EXAMPLE
        Uninstall-W11CursorScheme -SchemeName 'MyDarkCursors'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SchemeName
    )

    # ---- Verify the scheme exists ----
    $schemeExists = $false
    if (Test-Path $script:SchemesRegPath) {
        $existing = Get-ItemProperty -Path $script:SchemesRegPath -Name $SchemeName -ErrorAction SilentlyContinue
        if ($existing) {
            $schemeExists = $true
        }
    }

    if (-not $schemeExists) {
        Write-Warning "Cursor scheme '$SchemeName' is not registered. Nothing to uninstall."
        return
    }

    # ---- Remove scheme from registry ----
    Remove-ItemProperty -Path $script:SchemesRegPath -Name $SchemeName -ErrorAction Stop
    Write-Verbose "Removed scheme '$SchemeName' from registry."

    # ---- Delete cursor files ----
    $cursorFolder = Join-Path $env:LOCALAPPDATA "w11-theming-suite\cursors\$SchemeName"
    if (Test-Path $cursorFolder) {
        Remove-Item -Path $cursorFolder -Recurse -Force
        Write-Verbose "Deleted cursor files at $cursorFolder"
    }
    else {
        Write-Verbose "Cursor folder not found (already removed?): $cursorFolder"
    }

    # ---- Revert to Windows Default if this was the active scheme ----
    $activeName = (Get-ItemProperty -Path $script:CursorsRegPath -Name '(Default)' -ErrorAction SilentlyContinue).'(Default)'
    if ($activeName -eq $SchemeName) {
        Write-Verbose "Scheme '$SchemeName' was active. Reverting to Windows Default cursors."

        # Clear the active scheme name.
        Set-ItemProperty -Path $script:CursorsRegPath -Name '(Default)' -Value ''

        # Reset each cursor role to empty (system default).
        foreach ($role in $script:CursorRoles) {
            Set-ItemProperty -Path $script:CursorsRegPath -Name $role -Value '' -Type ExpandString
        }

        # Notify the system to reload cursors back to defaults.
        Add-W11NativeMethods
        $flags = [W11ThemeSuite.NativeMethods]::SPIF_UPDATEINIFILE -bor `
                 [W11ThemeSuite.NativeMethods]::SPIF_SENDCHANGE

        $result = [W11ThemeSuite.NativeMethods]::SystemParametersInfo(
            [W11ThemeSuite.NativeMethods]::SPI_SETCURSORS,
            0,
            [IntPtr]::Zero,
            $flags
        )

        if (-not $result) {
            Write-Warning 'SystemParametersInfo(SPI_SETCURSORS) returned false. Cursors may not have refreshed.'
        }
        else {
            Write-Verbose 'Cursors reverted to Windows Default.'
        }
    }

    Write-Verbose "Cursor scheme '$SchemeName' has been uninstalled."
}

function Get-W11CursorSchemes {
    <#
    .SYNOPSIS
        Enumerates all registered cursor schemes and identifies the active one.

    .DESCRIPTION
        Reads every scheme value from HKCU:\Control Panel\Cursors\Schemes and
        returns objects with the scheme name, its comma-separated path string,
        and whether it is the currently active scheme.

    .EXAMPLE
        Get-W11CursorSchemes | Format-Table -AutoSize
    #>
    [CmdletBinding()]
    param()

    # Determine the currently active scheme name.
    $activeName = (Get-ItemProperty -Path $script:CursorsRegPath -Name '(Default)' -ErrorAction SilentlyContinue).'(Default)'
    Write-Verbose "Currently active scheme: '$activeName'"

    # Read all registered schemes.
    if (-not (Test-Path $script:SchemesRegPath)) {
        Write-Verbose 'No schemes registry key found. Returning empty list.'
        return
    }

    $schemesKey = Get-ItemProperty -Path $script:SchemesRegPath -ErrorAction SilentlyContinue
    if (-not $schemesKey) {
        Write-Verbose 'No scheme values found.'
        return
    }

    # Iterate over all properties, skipping the PS* metadata properties.
    $schemesKey.PSObject.Properties | Where-Object {
        $_.Name -notmatch '^PS(Path|Drive|Provider|ParentPath|ChildName)$'
    } | ForEach-Object {
        [PSCustomObject]@{
            Name     = $_.Name
            Paths    = $_.Value
            IsActive = ($_.Name -eq $activeName)
        }
    }
}

# ---------------------------------------------------------------------------
# Module initialisation
# ---------------------------------------------------------------------------

# Export public functions (also declared in the .psd1 manifest).
Export-ModuleMember -Function @(
    'Install-W11CursorScheme',
    'Uninstall-W11CursorScheme',
    'Get-W11CursorSchemes'
)
