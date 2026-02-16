#Requires -Version 5.1
<#
.SYNOPSIS
    SoundSchemeBuilder module for w11-theming-suite.
.DESCRIPTION
    Manages custom Windows 11 sound schemes natively via the registry.
    Supports installing, activating, uninstalling, and listing sound schemes.

    Registry structure:
      HKCU:\AppEvents\Schemes              -> (Default) = active scheme name
      HKCU:\AppEvents\Schemes\Names\<Name> -> (Default) = display name
      HKCU:\AppEvents\Schemes\Apps\.Default\<Event>\<Scheme> -> (Default) = wav path
      HKCU:\AppEvents\Schemes\Apps\Explorer\<Event>\<Scheme> -> (Default) = wav path
      ...\.Current -> (Default) = currently active wav path
#>

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Standard system sound events registered under Apps\.Default
$Script:DefaultEvents = @(
    '.Default'
    'AppGPFault'
    'Close'
    'CriticalBatteryAlarm'
    'DeviceConnect'
    'DeviceDisconnect'
    'DeviceFail'
    'FaxBeep'
    'LowBatteryAlarm'
    'MailBeep'
    'Maximize'
    'MenuCommand'
    'MenuPopup'
    'Minimize'
    'Open'
    'PrintComplete'
    'RestoreDown'
    'RestoreUp'
    'SystemAsterisk'
    'SystemExclamation'
    'SystemExit'
    'SystemHand'
    'SystemNotification'
    'SystemQuestion'
    'WindowsLogoff'
    'WindowsLogon'
    'WindowsUAC'
)

# Explorer-specific sound events registered under Apps\Explorer
$Script:ExplorerEvents = @(
    'BlockedPopup'
    'EmptyRecycleBin'
    'FeedDiscovered'
    'Navigating'
    'SecurityBand'
)

# Registry base paths
$Script:SchemesRoot   = 'HKCU:\AppEvents\Schemes'
$Script:NamesRoot     = 'HKCU:\AppEvents\Schemes\Names'
$Script:AppsDefault   = 'HKCU:\AppEvents\Schemes\Apps\.Default'
$Script:AppsExplorer  = 'HKCU:\AppEvents\Schemes\Apps\Explorer'

# Local storage base for installed sound files
$Script:SoundsStore   = Join-Path $env:LOCALAPPDATA 'w11-theming-suite\sounds'

# ---------------------------------------------------------------------------
# Install-W11SoundScheme
# ---------------------------------------------------------------------------

function Install-W11SoundScheme {
    <#
    .SYNOPSIS
        Installs a custom sound scheme from a theme configuration.
    .DESCRIPTION
        Copies .wav files from the project assets into local app data,
        registers the scheme name in the registry, maps each configured
        event to its .wav file, and optionally activates the scheme.
    .PARAMETER Config
        A PSCustomObject containing a .sounds property with:
          - schemeName    : unique registry name for the scheme
          - setFolder     : subfolder under assets\sounds\ containing .wav files
          - events        : hashtable of { EventName = 'filename.wav' }
          - explorerEvents: (optional) hashtable of { EventName = 'filename.wav' }
    .PARAMETER Activate
        If specified, the scheme is immediately set as the active sound scheme.
    .EXAMPLE
        Install-W11SoundScheme -Config $themeConfig -Activate
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [switch]$Activate
    )

    # ---- Validate configuration ----
    if (-not $Config.sounds) {
        throw 'Config object is missing the .sounds property.'
    }
    $sounds = $Config.sounds

    if (-not $sounds.schemeName) {
        throw 'Config.sounds.schemeName is required.'
    }
    if (-not $sounds.setFolder) {
        throw 'Config.sounds.setFolder is required.'
    }

    $schemeName = $sounds.schemeName
    Write-Verbose "Installing sound scheme: $schemeName"

    # ---- Resolve paths ----
    # Project root is two levels above this module's directory
    $projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $sourceDir   = Join-Path $projectRoot "assets\sounds\$($sounds.setFolder)"
    $destDir     = Join-Path $Script:SoundsStore $schemeName

    if (-not (Test-Path $sourceDir)) {
        throw "Sound source folder not found: $sourceDir"
    }

    # ---- Copy .wav files to local storage ----
    Write-Verbose "Copying .wav files from '$sourceDir' to '$destDir'"
    if (-not (Test-Path $destDir)) {
        New-Item -Path $destDir -ItemType Directory -Force | Out-Null
    }

    $wavFiles = Get-ChildItem -Path $sourceDir -Filter '*.wav' -File
    if ($wavFiles.Count -eq 0) {
        Write-Warning "No .wav files found in source folder: $sourceDir"
    }

    foreach ($wav in $wavFiles) {
        Copy-Item -Path $wav.FullName -Destination $destDir -Force
        Write-Verbose "  Copied: $($wav.Name)"
    }

    # ---- Register the scheme name ----
    $schemeNameKey = Join-Path $Script:NamesRoot $schemeName
    Write-Verbose "Registering scheme at: $schemeNameKey"
    New-Item -Path $schemeNameKey -Force | Out-Null
    Set-ItemProperty -Path $schemeNameKey -Name '(Default)' -Value $schemeName

    # ---- Map events under Apps\.Default ----
    $registeredCount = 0
    if ($sounds.events) {
        foreach ($entry in $sounds.events.GetEnumerator()) {
            $eventName = $entry.Key
            $wavName   = $entry.Value
            $parentKey = Join-Path $Script:AppsDefault $eventName

            # Only register if the parent event key exists in the registry
            if (-not (Test-Path $parentKey)) {
                Write-Warning "Skipping unknown event '$eventName' (key not found: $parentKey)"
                continue
            }

            $eventSchemeKey = Join-Path $parentKey $schemeName
            $wavFullPath   = Join-Path $destDir $wavName

            if (-not (Test-Path $wavFullPath)) {
                Write-Warning "WAV file not found for event '$eventName': $wavFullPath"
                continue
            }

            Write-Verbose "  Mapping event '$eventName' -> $wavFullPath"
            New-Item -Path $eventSchemeKey -Force | Out-Null
            Set-ItemProperty -Path $eventSchemeKey -Name '(Default)' -Value $wavFullPath
            $registeredCount++
        }
    }

    # ---- Map events under Apps\Explorer (optional) ----
    $explorerCount = 0
    if ($sounds.explorerEvents) {
        foreach ($entry in $sounds.explorerEvents.GetEnumerator()) {
            $eventName = $entry.Key
            $wavName   = $entry.Value
            $parentKey = Join-Path $Script:AppsExplorer $eventName

            if (-not (Test-Path $parentKey)) {
                Write-Warning "Skipping unknown Explorer event '$eventName' (key not found: $parentKey)"
                continue
            }

            $eventSchemeKey = Join-Path $parentKey $schemeName
            $wavFullPath   = Join-Path $destDir $wavName

            if (-not (Test-Path $wavFullPath)) {
                Write-Warning "WAV file not found for Explorer event '$eventName': $wavFullPath"
                continue
            }

            Write-Verbose "  Mapping Explorer event '$eventName' -> $wavFullPath"
            New-Item -Path $eventSchemeKey -Force | Out-Null
            Set-ItemProperty -Path $eventSchemeKey -Name '(Default)' -Value $wavFullPath
            $explorerCount++
        }
    }

    # ---- Activate the scheme if requested ----
    if ($Activate) {
        Write-Verbose "Activating scheme '$schemeName'"
        Set-ItemProperty -Path $Script:SchemesRoot -Name '(Default)' -Value $schemeName

        # Copy each registered event's wav path into the .Current subkey
        if ($sounds.events) {
            foreach ($entry in $sounds.events.GetEnumerator()) {
                $eventName = $entry.Key
                $wavName   = $entry.Value
                $parentKey = Join-Path $Script:AppsDefault $eventName
                $currentKey = Join-Path $parentKey '.Current'
                $wavFullPath = Join-Path $destDir $wavName

                if ((Test-Path $parentKey) -and (Test-Path $wavFullPath)) {
                    New-Item -Path $currentKey -Force | Out-Null
                    Set-ItemProperty -Path $currentKey -Name '(Default)' -Value $wavFullPath
                }
            }
        }

        if ($sounds.explorerEvents) {
            foreach ($entry in $sounds.explorerEvents.GetEnumerator()) {
                $eventName = $entry.Key
                $wavName   = $entry.Value
                $parentKey = Join-Path $Script:AppsExplorer $eventName
                $currentKey = Join-Path $parentKey '.Current'
                $wavFullPath = Join-Path $destDir $wavName

                if ((Test-Path $parentKey) -and (Test-Path $wavFullPath)) {
                    New-Item -Path $currentKey -Force | Out-Null
                    Set-ItemProperty -Path $currentKey -Name '(Default)' -Value $wavFullPath
                }
            }
        }
    }

    # ---- Summary output ----
    $status = if ($Activate) { 'Installed and activated' } else { 'Installed' }
    Write-Host "Sound scheme '$schemeName': $status" -ForegroundColor Green
    Write-Host "  .Default events registered : $registeredCount"
    Write-Host "  Explorer events registered : $explorerCount"
    Write-Host "  Sound files stored in      : $destDir"
}

# ---------------------------------------------------------------------------
# Uninstall-W11SoundScheme
# ---------------------------------------------------------------------------

function Uninstall-W11SoundScheme {
    <#
    .SYNOPSIS
        Removes a custom sound scheme from the registry and local storage.
    .DESCRIPTION
        Deletes the scheme registration, removes all per-event registry keys
        for the scheme, cleans up stored .wav files, and reverts to the
        Windows default scheme if this was the active scheme.
    .PARAMETER SchemeName
        The registry name of the scheme to uninstall.
    .EXAMPLE
        Uninstall-W11SoundScheme -SchemeName 'MyCustomScheme'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SchemeName
    )

    Write-Verbose "Uninstalling sound scheme: $SchemeName"

    # ---- Verify the scheme exists ----
    $schemeNameKey = Join-Path $Script:NamesRoot $SchemeName
    if (-not (Test-Path $schemeNameKey)) {
        Write-Error "Sound scheme '$SchemeName' is not installed (key not found: $schemeNameKey)"
        return
    }

    # ---- Check if this is the currently active scheme ----
    $activeScheme = (Get-ItemProperty -Path $Script:SchemesRoot -Name '(Default)').'(Default)'
    $wasActive = ($activeScheme -eq $SchemeName)

    # ---- Remove scheme name registration ----
    Write-Verbose "Removing scheme registration: $schemeNameKey"
    Remove-Item -Path $schemeNameKey -Recurse -Force

    # ---- Remove per-event subkeys under Apps\.Default ----
    Write-Verbose "Cleaning event keys under Apps\.Default"
    foreach ($eventName in $Script:DefaultEvents) {
        $eventSchemeKey = Join-Path $Script:AppsDefault "$eventName\$SchemeName"
        if (Test-Path $eventSchemeKey) {
            Remove-Item -Path $eventSchemeKey -Recurse -Force
            Write-Verbose "  Removed: $eventSchemeKey"
        }
    }

    # ---- Remove per-event subkeys under Apps\Explorer ----
    Write-Verbose "Cleaning event keys under Apps\Explorer"
    foreach ($eventName in $Script:ExplorerEvents) {
        $eventSchemeKey = Join-Path $Script:AppsExplorer "$eventName\$SchemeName"
        if (Test-Path $eventSchemeKey) {
            Remove-Item -Path $eventSchemeKey -Recurse -Force
            Write-Verbose "  Removed: $eventSchemeKey"
        }
    }

    # ---- Delete stored sound files ----
    $soundDir = Join-Path $Script:SoundsStore $SchemeName
    if (Test-Path $soundDir) {
        Write-Verbose "Deleting sound files: $soundDir"
        Remove-Item -Path $soundDir -Recurse -Force
    }

    # ---- Revert to Windows default if this was the active scheme ----
    if ($wasActive) {
        Write-Verbose "Reverting active scheme to '.Default'"
        Set-ItemProperty -Path $Script:SchemesRoot -Name '(Default)' -Value '.Default'

        # Restore .Current subkeys from the .Default scheme values
        foreach ($eventName in $Script:DefaultEvents) {
            $defaultKey = Join-Path $Script:AppsDefault "$eventName\.Default"
            $currentKey = Join-Path $Script:AppsDefault "$eventName\.Current"

            if (Test-Path $defaultKey) {
                $defaultValue = (Get-ItemProperty -Path $defaultKey -Name '(Default)').'(Default)'
                New-Item -Path $currentKey -Force | Out-Null
                Set-ItemProperty -Path $currentKey -Name '(Default)' -Value $defaultValue
            }
        }

        foreach ($eventName in $Script:ExplorerEvents) {
            $defaultKey = Join-Path $Script:AppsExplorer "$eventName\.Default"
            $currentKey = Join-Path $Script:AppsExplorer "$eventName\.Current"

            if (Test-Path $defaultKey) {
                $defaultValue = (Get-ItemProperty -Path $defaultKey -Name '(Default)').'(Default)'
                New-Item -Path $currentKey -Force | Out-Null
                Set-ItemProperty -Path $currentKey -Name '(Default)' -Value $defaultValue
            }
        }

        Write-Host "Active scheme reverted to Windows default." -ForegroundColor Yellow
    }

    Write-Host "Sound scheme '$SchemeName' has been uninstalled." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Get-W11SoundSchemes
# ---------------------------------------------------------------------------

function Get-W11SoundSchemes {
    <#
    .SYNOPSIS
        Lists all registered sound schemes visible in the registry.
    .DESCRIPTION
        Enumerates HKCU:\AppEvents\Schemes\Names and returns each scheme
        with its display name and whether it is currently active.
    .EXAMPLE
        Get-W11SoundSchemes
    .EXAMPLE
        Get-W11SoundSchemes | Where-Object IsActive
    #>
    [CmdletBinding()]
    param()

    # Get the currently active scheme name
    $activeScheme = (Get-ItemProperty -Path $Script:SchemesRoot -Name '(Default)').'(Default)'
    Write-Verbose "Active scheme: $activeScheme"

    # Enumerate all registered scheme names
    if (-not (Test-Path $Script:NamesRoot)) {
        Write-Warning "No sound schemes registry key found at: $Script:NamesRoot"
        return
    }

    $schemeKeys = Get-ChildItem -Path $Script:NamesRoot -ErrorAction SilentlyContinue

    if (-not $schemeKeys -or $schemeKeys.Count -eq 0) {
        Write-Verbose "No sound schemes found under $($Script:NamesRoot)"
        return
    }

    foreach ($key in $schemeKeys) {
        $name = $key.PSChildName
        $displayName = (Get-ItemProperty -Path $key.PSPath -Name '(Default)').'(Default)'

        # If the display name is empty, fall back to the key name
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = $name
        }

        [PSCustomObject]@{
            Name        = $name
            DisplayName = $displayName
            IsActive    = ($name -eq $activeScheme)
        }
    }
}

# ---------------------------------------------------------------------------
# Module exports
# ---------------------------------------------------------------------------
Export-ModuleMember -Function @(
    'Install-W11SoundScheme'
    'Uninstall-W11SoundScheme'
    'Get-W11SoundSchemes'
)
