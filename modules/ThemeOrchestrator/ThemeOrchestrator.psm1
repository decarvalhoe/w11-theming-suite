#Requires -Version 5.1
<#
.SYNOPSIS
    ThemeOrchestrator module for w11-theming-suite.

.DESCRIPTION
    High-level orchestrator that coordinates all other theming modules
    (ConfigManager, CursorManager, SoundManager, WallpaperManager,
    RegistryManager, ThemeFileManager, BackupManager) to install,
    uninstall, and switch Windows 11 themes.

    This module assumes all dependent modules are already loaded by
    the root module.
#>

# ---------------------------------------------------------------------------
# Install-W11Theme
# ---------------------------------------------------------------------------

function Install-W11Theme {
    <#
    .SYNOPSIS
        Installs and applies a Windows 11 theme from a preset or config file.

    .DESCRIPTION
        Orchestrates the full theme installation pipeline:
        1. Loads configuration (preset or custom path)
        2. Creates a backup of the current state (unless -NoBackup)
        3. Applies cursors, sounds, wallpaper, registry settings, and .theme file
        4. Re-applies registry overrides after Windows processes the .theme file

    .PARAMETER PresetName
        Name of a built-in preset to install (resolved via Get-W11ThemeConfig).

    .PARAMETER ConfigPath
        Path to a custom theme configuration JSON file.

    .PARAMETER NoBackup
        Skip creating a backup of the current theme state before installation.

    .PARAMETER Force
        Overwrite existing backup if one with the same name already exists.

    .EXAMPLE
        Install-W11Theme -PresetName "Nordic"

    .EXAMPLE
        Install-W11Theme -ConfigPath "C:\MyThemes\custom.json"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]$PresetName,

        [Parameter()]
        [string]$ConfigPath,

        [Parameter()]
        [switch]$NoBackup,

        [Parameter()]
        [switch]$Force
    )

    try {
        # ------------------------------------------------------------------
        # 1. Load configuration
        # ------------------------------------------------------------------
        if ($PresetName) {
            Write-Host "[1/6] Loading preset '$PresetName'..." -ForegroundColor Cyan
            $config = Get-W11ThemeConfig -PresetName $PresetName
        }
        elseif ($ConfigPath) {
            Write-Host "[1/6] Loading config from '$ConfigPath'..." -ForegroundColor Cyan
            $config = Get-W11ThemeConfig -Path $ConfigPath
        }
        else {
            throw "You must specify either -PresetName or -ConfigPath."
        }

        # ------------------------------------------------------------------
        # 2. Display banner
        # ------------------------------------------------------------------
        $themeName    = $config.meta.name
        $themeVersion = $config.meta.version
        $themeAuthor  = $config.meta.author

        Write-Host ""
        Write-Host "=============================================" -ForegroundColor Cyan
        Write-Host " Installing theme: $themeName v$themeVersion by $themeAuthor" -ForegroundColor Cyan
        Write-Host "=============================================" -ForegroundColor Cyan
        Write-Host ""

        # ------------------------------------------------------------------
        # 3. Backup current state (unless -NoBackup)
        # ------------------------------------------------------------------
        $backupLabel = "pre-$($themeName -replace '\s','_')"

        if (-not $NoBackup) {
            Write-Host "[2/6] Backing up current theme state as '$backupLabel'..." -ForegroundColor Yellow
            Backup-W11ThemeState -Name $backupLabel -Force:$Force
        }
        else {
            Write-Host "[2/6] Skipping backup (-NoBackup specified)." -ForegroundColor Yellow
        }

        # ------------------------------------------------------------------
        # 4. Step-by-step application
        # ------------------------------------------------------------------

        # 4a. Cursors
        if ($config.cursors) {
            Write-Host "[3/6] Installing cursor scheme..." -ForegroundColor Cyan
            Install-W11CursorScheme -Config $config -Activate
        }
        else {
            Write-Host "[3/6] No cursor configuration found, skipping." -ForegroundColor Yellow
        }

        # 4b. Sounds
        if ($config.sounds) {
            Write-Host "[4/6] Installing sound scheme..." -ForegroundColor Cyan
            Install-W11SoundScheme -Config $config -Activate
        }
        else {
            Write-Host "[4/6] No sound configuration found, skipping." -ForegroundColor Yellow
        }

        # 4c. Wallpaper
        if ($config.wallpaper) {
            Write-Host "[5/6] Setting wallpaper..." -ForegroundColor Cyan
            Set-W11Wallpaper -Config $config
        }
        else {
            Write-Host "[5/6] No wallpaper configuration found, skipping." -ForegroundColor Yellow
        }

        # 4d. Apply registry theme settings
        Write-Host "[6/6] Applying registry theme settings..." -ForegroundColor Cyan
        Set-W11RegistryTheme -Config $config

        # 4e. Generate and apply .theme file
        Write-Host "       Generating .theme file..." -ForegroundColor Cyan
        $themeFile = New-W11ThemeFile -Config $config

        Write-Host "       Applying .theme file (Windows will process it)..." -ForegroundColor Cyan
        Start-Process -FilePath $themeFile -Wait

        # 4f. Wait for Windows to settle, then re-apply registry overrides
        #     The .theme file application may reset some registry values,
        #     so we force our desired settings back.
        Write-Host "       Waiting for Windows to settle..." -ForegroundColor Cyan
        Start-Sleep -Seconds 2

        Write-Host "       Re-applying registry overrides..." -ForegroundColor Cyan
        Set-W11RegistryTheme -Config $config -Section @('DarkMode', 'AccentColor', 'DWM', 'Taskbar')

        # ------------------------------------------------------------------
        # 5. Success summary
        # ------------------------------------------------------------------
        Write-Host ""
        Write-Host "Theme '$themeName' applied successfully!" -ForegroundColor Green
        Write-Host ""

        # ------------------------------------------------------------------
        # 6. Backup reminder
        # ------------------------------------------------------------------
        if (-not $NoBackup) {
            Write-Host "Backup saved as '$backupLabel'. Use Uninstall-W11Theme to restore." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "ERROR: Failed to install theme. $_" -ForegroundColor Red
        throw
    }
}


# ---------------------------------------------------------------------------
# Uninstall-W11Theme
# ---------------------------------------------------------------------------

function Uninstall-W11Theme {
    <#
    .SYNOPSIS
        Restores a previous theme state from a backup.

    .DESCRIPTION
        Lists available backups (or uses the specified one) and restores
        the Windows 11 theme state from that backup.

    .PARAMETER BackupName
        Name of the backup to restore. If omitted, the most recent backup
        is used automatically.

    .EXAMPLE
        Uninstall-W11Theme -BackupName "pre-Nordic"

    .EXAMPLE
        Uninstall-W11Theme
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]$BackupName
    )

    try {
        # If no backup name provided, find the most recent one
        if (-not $BackupName) {
            Write-Host "No backup name specified. Retrieving available backups..." -ForegroundColor Cyan

            $backups = Get-W11ThemeBackups

            if (-not $backups -or $backups.Count -eq 0) {
                Write-Host "No backups found. Nothing to restore." -ForegroundColor Yellow
                return
            }

            # Use the most recent backup (last in the list)
            $BackupName = ($backups | Select-Object -Last 1).Name
            Write-Host "Using most recent backup: '$BackupName'" -ForegroundColor Cyan
        }

        # Restore the backup
        Write-Host "Restoring theme state from backup '$BackupName'..." -ForegroundColor Cyan
        Restore-W11ThemeState -Name $BackupName

        Write-Host ""
        Write-Host "Theme restored from backup '$BackupName'." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to restore theme. $_" -ForegroundColor Red
        throw
    }
}


# ---------------------------------------------------------------------------
# Switch-W11Theme
# ---------------------------------------------------------------------------

function Switch-W11Theme {
    <#
    .SYNOPSIS
        Quickly switches to a different theme preset.

    .DESCRIPTION
        Ensures an "original" backup exists (creating one if needed),
        then installs the specified preset without creating an additional
        backup. This allows fast switching between themes while always
        preserving the original system state.

    .PARAMETER PresetName
        Name of the preset theme to switch to.

    .EXAMPLE
        Switch-W11Theme -PresetName "Catppuccin"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$PresetName
    )

    try {
        # ------------------------------------------------------------------
        # 1. Ensure an "original" backup exists
        # ------------------------------------------------------------------
        $backups = Get-W11ThemeBackups

        $hasOriginal = $false
        if ($backups) {
            $hasOriginal = ($backups | Where-Object { $_.Name -eq 'original' }) -ne $null
        }

        if (-not $hasOriginal) {
            Write-Host "Saving current state as 'original' backup..." -ForegroundColor Yellow
            Backup-W11ThemeState -Name 'original'
        }
        else {
            Write-Host "Original backup already exists, preserving it." -ForegroundColor Cyan
        }

        # ------------------------------------------------------------------
        # 2. Install the new theme without creating another backup
        # ------------------------------------------------------------------
        Install-W11Theme -PresetName $PresetName -NoBackup

        # ------------------------------------------------------------------
        # 3. Summary
        # ------------------------------------------------------------------
        Write-Host ""
        Write-Host "Switched to theme '$PresetName'. Original state saved as 'original'." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to switch theme. $_" -ForegroundColor Red
        throw
    }
}


# ---------------------------------------------------------------------------
# Get-W11InstalledThemes
# ---------------------------------------------------------------------------

function Get-W11InstalledThemes {
    <#
    .SYNOPSIS
        Lists all available theme configurations (presets and user themes).

    .DESCRIPTION
        Scans the project's config\presets\ and config\user\ directories
        for .json theme configuration files and returns summary information
        about each one.

    .EXAMPLE
        Get-W11InstalledThemes

    .EXAMPLE
        Get-W11InstalledThemes | Format-Table Name, Version, Source
    #>
    [CmdletBinding()]
    param()

    try {
        # Determine project root (ThemeOrchestrator.psm1 is at modules\ThemeOrchestrator\)
        $ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path

        $results = @()

        # Define the directories to scan and their source labels
        $scanTargets = @(
            @{ Path = Join-Path $ProjectRoot 'config\presets'; Source = 'preset' }
            @{ Path = Join-Path $ProjectRoot 'config\user';    Source = 'user'   }
        )

        foreach ($target in $scanTargets) {
            $dir    = $target.Path
            $source = $target.Source

            if (-not (Test-Path $dir)) {
                Write-Verbose "Directory not found, skipping: $dir"
                continue
            }

            $jsonFiles = Get-ChildItem -Path $dir -Filter '*.json' -File -ErrorAction SilentlyContinue

            foreach ($file in $jsonFiles) {
                try {
                    $raw  = Get-Content -Path $file.FullName -Raw -ErrorAction Stop
                    $json = $raw | ConvertFrom-Json -ErrorAction Stop

                    # Extract meta section; gracefully handle missing fields
                    $meta = $json.meta

                    $entry = [PSCustomObject]@{
                        Name        = if ($meta.name)        { $meta.name }        else { $file.BaseName }
                        Version     = if ($meta.version)     { $meta.version }     else { 'N/A' }
                        Author      = if ($meta.author)      { $meta.author }      else { 'Unknown' }
                        Description = if ($meta.description) { $meta.description } else { '' }
                        Tags        = if ($meta.tags)        { $meta.tags }        else { @() }
                        Source      = $source
                        FileName    = $file.Name
                    }

                    $results += $entry
                }
                catch {
                    Write-Host "WARNING: Could not parse '$($file.Name)': $_" -ForegroundColor Yellow
                }
            }
        }

        if ($results.Count -eq 0) {
            Write-Host "No theme configurations found." -ForegroundColor Yellow
        }
        else {
            Write-Host "Found $($results.Count) theme(s)." -ForegroundColor Cyan
        }

        return $results
    }
    catch {
        Write-Host "ERROR: Failed to enumerate themes. $_" -ForegroundColor Red
        throw
    }
}
