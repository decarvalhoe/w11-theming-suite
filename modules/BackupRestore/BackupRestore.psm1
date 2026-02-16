# BackupRestore.psm1
# Snapshot and restore Windows 11 theme state for w11-theming-suite.
# Captures registry settings, cursors, sounds, wallpaper into a named backup folder
# and can restore them to return the system to a previous theme state.

# ---------------------------------------------------------------------------
# Project root detection: modules\BackupRestore -> go up two levels
# ---------------------------------------------------------------------------
$script:ProjectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

# ---------------------------------------------------------------------------
# P/Invoke: load native methods for cursor refresh, wallpaper, and broadcast
# Only add the type if it hasn't been loaded yet (same pattern as RegistryConfigurator)
# ---------------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'W11ThemeSuite.NativeMethods').Type) {
    Add-Type -Namespace 'W11ThemeSuite' -Name 'NativeMethods' -MemberDefinition @'
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SystemParametersInfo(
            uint uiAction, uint uiParam, string pvParam, uint fWinIni);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(
            IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam,
            uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);

        public const uint SPI_SETCURSORS       = 0x0057;
        public const uint SPI_SETDESKWALLPAPER  = 0x0014;
        public const uint SPIF_UPDATEINIFILE    = 0x01;
        public const uint SPIF_SENDCHANGE       = 0x02;

        public const uint WM_SETTINGCHANGE     = 0x001A;
        public static readonly IntPtr HWND_BROADCAST = new IntPtr(0xFFFF);
        public const uint SMTO_ABORTIFHUNG     = 0x0002;
'@
}

# ---------------------------------------------------------------------------
# Helper: Broadcast WM_SETTINGCHANGE so Explorer picks up registry changes
# ---------------------------------------------------------------------------
function Send-SettingChange {
    $result = [UIntPtr]::Zero
    [W11ThemeSuite.NativeMethods]::SendMessageTimeout(
        [W11ThemeSuite.NativeMethods]::HWND_BROADCAST,
        [W11ThemeSuite.NativeMethods]::WM_SETTINGCHANGE,
        [UIntPtr]::Zero,
        'ImmutableControl',
        [W11ThemeSuite.NativeMethods]::SMTO_ABORTIFHUNG,
        5000,
        [ref]$result
    ) | Out-Null
}

# ---------------------------------------------------------------------------
# Helper: Flatten the nested RegistryMap hashtable into a flat list of entries
# Each entry gets: Section, Key, Path, Name, Type
# ---------------------------------------------------------------------------
function Get-FlatRegistryMap {
    param([hashtable]$Map)

    $entries = @()
    foreach ($section in $Map.Keys) {
        foreach ($key in $Map[$section].Keys) {
            $entry = $Map[$section][$key]
            $entries += [PSCustomObject]@{
                Section = $section
                Key     = $key
                Path    = $entry.Path
                Name    = $entry.Name
                Type    = $entry.Type
            }
        }
    }
    return $entries
}

# ===========================================================================
#  Backup-W11ThemeState
# ===========================================================================
function Backup-W11ThemeState {
    <#
    .SYNOPSIS
        Creates a snapshot of the current Windows 11 theme state.
    .DESCRIPTION
        Reads registry keys defined in RegistryMap, cursor scheme, sound scheme,
        and wallpaper settings, then saves them to a named backup folder.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Name,

        [Parameter()]
        [switch]$Force
    )

    # Auto-generate name if not supplied
    if (-not $Name) {
        $Name = "backup_$(Get-Date -Format 'yyyy-MM-dd_HHmmss')"
    }

    $backupDir = Join-Path $script:ProjectRoot "backups\$Name"

    # Guard against overwriting an existing backup
    if ((Test-Path $backupDir) -and -not $Force) {
        Write-Error "Backup '$Name' already exists at '$backupDir'. Use -Force to overwrite."
        return
    }

    # Create (or recreate) the backup folder
    New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
    Write-Verbose "Backup directory created: $backupDir"

    $timestamp = Get-Date -Format 'o'  # ISO 8601

    # ------------------------------------------------------------------
    # 1. Registry snapshot from RegistryMap
    # ------------------------------------------------------------------
    $registryMapPath = Join-Path $script:ProjectRoot 'modules\RegistryConfigurator\RegistryMap.psd1'
    if (-not (Test-Path $registryMapPath)) {
        Write-Error "RegistryMap.psd1 not found at '$registryMapPath'."
        return
    }

    $registryMap = Import-PowerShellDataFile -Path $registryMapPath
    $flatEntries = Get-FlatRegistryMap -Map $registryMap

    $registrySnapshot = @()
    foreach ($entry in $flatEntries) {
        $value = $null
        $exists = $false
        try {
            $regValue = Get-ItemProperty -Path $entry.Path -Name $entry.Name -ErrorAction Stop
            $value = $regValue.($entry.Name)
            $exists = $true
        }
        catch {
            # Key or value does not exist; record as null
        }

        $registrySnapshot += [PSCustomObject]@{
            Section = $entry.Section
            Key     = $entry.Key
            Path    = $entry.Path
            Name    = $entry.Name
            Type    = $entry.Type
            Value   = if ($exists) { $value } else { $null }
            Exists  = $exists
        }
    }

    $registrySnapshot | ConvertTo-Json -Depth 10 |
        Set-Content -Path (Join-Path $backupDir 'registry-snapshot.json') -Encoding UTF8
    Write-Verbose "Registry snapshot saved ($($registrySnapshot.Count) entries)."

    # ------------------------------------------------------------------
    # 2. Cursor scheme
    # ------------------------------------------------------------------
    $cursorRoles = @(
        'Arrow', 'Help', 'AppStarting', 'Wait', 'Crosshair',
        'IBeam', 'NWPen', 'No', 'SizeNS', 'SizeWE',
        'SizeNWSE', 'SizeNESW', 'SizeAll', 'UpArrow', 'Hand'
    )

    $cursorPath = 'HKCU:\Control Panel\Cursors'
    $cursorData = @{}
    try {
        $cursorProps = Get-ItemProperty -Path $cursorPath -ErrorAction Stop
        foreach ($role in $cursorRoles) {
            $cursorData[$role] = $cursorProps.$role
        }
        # The (Default) value stores the scheme name
        $cursorData['SchemeName'] = $cursorProps.'(default)'
    }
    catch {
        Write-Warning "Could not read cursor scheme: $_"
    }

    $cursorData | ConvertTo-Json -Depth 5 |
        Set-Content -Path (Join-Path $backupDir 'cursors-backup.json') -Encoding UTF8
    Write-Verbose "Cursor scheme saved."

    # ------------------------------------------------------------------
    # 3. Sound scheme
    # ------------------------------------------------------------------
    $soundData = @{}
    try {
        $soundScheme = Get-ItemProperty -Path 'HKCU:\AppEvents\Schemes' -ErrorAction Stop
        $soundData['SchemeName'] = $soundScheme.'(default)'
    }
    catch {
        Write-Warning "Could not read sound scheme: $_"
    }

    $soundData | ConvertTo-Json -Depth 5 |
        Set-Content -Path (Join-Path $backupDir 'sounds-backup.json') -Encoding UTF8
    Write-Verbose "Sound scheme saved."

    # ------------------------------------------------------------------
    # 4. Wallpaper settings
    # ------------------------------------------------------------------
    $wallpaperData = @{}
    try {
        $desktopProps = Get-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -ErrorAction Stop
        $wallpaperData['Wallpaper']      = $desktopProps.Wallpaper
        $wallpaperData['WallpaperStyle'] = $desktopProps.WallpaperStyle
        $wallpaperData['TileWallpaper']  = $desktopProps.TileWallpaper
    }
    catch {
        Write-Warning "Could not read wallpaper settings: $_"
    }

    $wallpaperData | ConvertTo-Json -Depth 5 |
        Set-Content -Path (Join-Path $backupDir 'wallpaper-backup.json') -Encoding UTF8
    Write-Verbose "Wallpaper settings saved."

    # ------------------------------------------------------------------
    # 5. Current theme path
    # ------------------------------------------------------------------
    $currentThemePath = $null
    try {
        $themeProps = Get-ItemProperty `
            -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes' -ErrorAction Stop
        $currentThemePath = $themeProps.CurrentTheme
    }
    catch {
        Write-Warning "Could not read current theme path: $_"
    }

    # ------------------------------------------------------------------
    # 6. OS build information
    # ------------------------------------------------------------------
    $osBuild = $null
    try {
        $ntVersion = Get-ItemProperty `
            -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        $osBuild = $ntVersion.DisplayVersion
    }
    catch {
        Write-Warning "Could not read OS build version: $_"
    }

    # ------------------------------------------------------------------
    # 7. Write backup manifest
    # ------------------------------------------------------------------
    $manifest = [ordered]@{
        Name             = $Name
        Timestamp        = $timestamp
        OSBuild          = $osBuild
        CurrentThemePath = $currentThemePath
    }

    $manifest | ConvertTo-Json -Depth 5 |
        Set-Content -Path (Join-Path $backupDir 'backup-manifest.json') -Encoding UTF8
    Write-Verbose "Backup manifest written."

    Write-Host "Backup '$Name' created successfully at '$backupDir'." -ForegroundColor Green

    # Return a summary object
    [PSCustomObject]@{
        Name      = $Name
        Path      = $backupDir
        Timestamp = $timestamp
    }
}

# ===========================================================================
#  Restore-W11ThemeState
# ===========================================================================
function Restore-W11ThemeState {
    <#
    .SYNOPSIS
        Restores a previously saved Windows 11 theme state.
    .DESCRIPTION
        Reads a named backup (or a direct path) and applies the stored registry
        keys, cursor scheme, and wallpaper settings back to the system.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Name,

        [Parameter()]
        [string]$BackupPath
    )

    # Resolve the backup directory
    if ($Name) {
        $backupDir = Join-Path $script:ProjectRoot "backups\$Name"
    }
    elseif ($BackupPath) {
        $backupDir = $BackupPath
    }
    else {
        Write-Error "You must provide either -Name or -BackupPath."
        return
    }

    if (-not (Test-Path $backupDir)) {
        Write-Error "Backup directory not found: '$backupDir'."
        return
    }

    # ------------------------------------------------------------------
    # Verify backup integrity via manifest
    # ------------------------------------------------------------------
    $manifestPath = Join-Path $backupDir 'backup-manifest.json'
    if (-not (Test-Path $manifestPath)) {
        Write-Error "backup-manifest.json not found in '$backupDir'. The backup may be corrupt."
        return
    }

    $manifest = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json
    Write-Verbose "Restoring backup '$($manifest.Name)' from $($manifest.Timestamp)."

    $restoredCount = 0
    $removedCount  = 0
    $errorCount    = 0

    # ------------------------------------------------------------------
    # 1. Restore registry snapshot
    # ------------------------------------------------------------------
    $snapshotPath = Join-Path $backupDir 'registry-snapshot.json'
    if (Test-Path $snapshotPath) {
        $snapshot = Get-Content -Path $snapshotPath -Raw | ConvertFrom-Json

        foreach ($entry in $snapshot) {
            try {
                if (-not $entry.Exists) {
                    # Value did not exist at backup time; remove it if it now exists
                    if (Test-Path $entry.Path) {
                        $current = Get-ItemProperty -Path $entry.Path -Name $entry.Name -ErrorAction SilentlyContinue
                        if ($null -ne $current) {
                            Remove-ItemProperty -Path $entry.Path -Name $entry.Name -ErrorAction Stop
                            $removedCount++
                            Write-Verbose "Removed: $($entry.Path)\$($entry.Name)"
                        }
                    }
                }
                else {
                    # Ensure the parent registry path exists
                    if (-not (Test-Path $entry.Path)) {
                        New-Item -Path $entry.Path -Force | Out-Null
                    }

                    # Map string type name to the RegistryValueKind enum
                    $regType = switch ($entry.Type) {
                        'DWord'        { [Microsoft.Win32.RegistryValueKind]::DWord }
                        'QWord'        { [Microsoft.Win32.RegistryValueKind]::QWord }
                        'String'       { [Microsoft.Win32.RegistryValueKind]::String }
                        'ExpandString'  { [Microsoft.Win32.RegistryValueKind]::ExpandString }
                        'MultiString'   { [Microsoft.Win32.RegistryValueKind]::MultiString }
                        'Binary'        { [Microsoft.Win32.RegistryValueKind]::Binary }
                        default         { [Microsoft.Win32.RegistryValueKind]::String }
                    }

                    Set-ItemProperty -Path $entry.Path -Name $entry.Name `
                        -Value $entry.Value -Type $regType -ErrorAction Stop
                    $restoredCount++
                    Write-Verbose "Restored: $($entry.Path)\$($entry.Name) = $($entry.Value)"
                }
            }
            catch {
                $errorCount++
                Write-Warning "Failed to restore $($entry.Path)\$($entry.Name): $_"
            }
        }

        Write-Verbose "Registry: $restoredCount restored, $removedCount removed, $errorCount errors."
    }
    else {
        Write-Warning "registry-snapshot.json not found; skipping registry restore."
    }

    # ------------------------------------------------------------------
    # 2. Restore cursor scheme
    # ------------------------------------------------------------------
    $cursorsPath = Join-Path $backupDir 'cursors-backup.json'
    if (Test-Path $cursorsPath) {
        $cursorData = Get-Content -Path $cursorsPath -Raw | ConvertFrom-Json
        $cursorRegPath = 'HKCU:\Control Panel\Cursors'

        # Write the scheme name to the (Default) value
        if ($cursorData.SchemeName) {
            Set-ItemProperty -Path $cursorRegPath -Name '(default)' `
                -Value $cursorData.SchemeName -ErrorAction SilentlyContinue
        }

        # Write each cursor role
        $cursorRoles = @(
            'Arrow', 'Help', 'AppStarting', 'Wait', 'Crosshair',
            'IBeam', 'NWPen', 'No', 'SizeNS', 'SizeWE',
            'SizeNWSE', 'SizeNESW', 'SizeAll', 'UpArrow', 'Hand'
        )
        foreach ($role in $cursorRoles) {
            $val = $cursorData.$role
            if ($null -ne $val) {
                Set-ItemProperty -Path $cursorRegPath -Name $role `
                    -Value $val -ErrorAction SilentlyContinue
            }
        }

        # Apply cursor change immediately via SPI_SETCURSORS
        [W11ThemeSuite.NativeMethods]::SystemParametersInfo(
            [W11ThemeSuite.NativeMethods]::SPI_SETCURSORS, 0, $null,
            [W11ThemeSuite.NativeMethods]::SPIF_UPDATEINIFILE -bor
            [W11ThemeSuite.NativeMethods]::SPIF_SENDCHANGE
        ) | Out-Null

        Write-Verbose "Cursor scheme restored and applied."
    }
    else {
        Write-Warning "cursors-backup.json not found; skipping cursor restore."
    }

    # ------------------------------------------------------------------
    # 3. Restore wallpaper settings
    # ------------------------------------------------------------------
    $wallpaperPath = Join-Path $backupDir 'wallpaper-backup.json'
    if (Test-Path $wallpaperPath) {
        $wallpaperData = Get-Content -Path $wallpaperPath -Raw | ConvertFrom-Json
        $desktopRegPath = 'HKCU:\Control Panel\Desktop'

        if ($null -ne $wallpaperData.WallpaperStyle) {
            Set-ItemProperty -Path $desktopRegPath -Name 'WallpaperStyle' `
                -Value $wallpaperData.WallpaperStyle -ErrorAction SilentlyContinue
        }
        if ($null -ne $wallpaperData.TileWallpaper) {
            Set-ItemProperty -Path $desktopRegPath -Name 'TileWallpaper' `
                -Value $wallpaperData.TileWallpaper -ErrorAction SilentlyContinue
        }
        if ($wallpaperData.Wallpaper) {
            Set-ItemProperty -Path $desktopRegPath -Name 'Wallpaper' `
                -Value $wallpaperData.Wallpaper -ErrorAction SilentlyContinue

            # Apply wallpaper immediately via SPI_SETDESKWALLPAPER
            [W11ThemeSuite.NativeMethods]::SystemParametersInfo(
                [W11ThemeSuite.NativeMethods]::SPI_SETDESKWALLPAPER, 0,
                $wallpaperData.Wallpaper,
                [W11ThemeSuite.NativeMethods]::SPIF_UPDATEINIFILE -bor
                [W11ThemeSuite.NativeMethods]::SPIF_SENDCHANGE
            ) | Out-Null
        }

        Write-Verbose "Wallpaper settings restored."
    }
    else {
        Write-Warning "wallpaper-backup.json not found; skipping wallpaper restore."
    }

    # ------------------------------------------------------------------
    # 4. Broadcast setting change to refresh the shell
    # ------------------------------------------------------------------
    Send-SettingChange
    Write-Verbose "WM_SETTINGCHANGE broadcast sent."

    # ------------------------------------------------------------------
    # Output summary
    # ------------------------------------------------------------------
    Write-Host "Restore complete for backup '$($manifest.Name)'." -ForegroundColor Green
    Write-Host "  Registry values restored: $restoredCount" -ForegroundColor Cyan
    Write-Host "  Registry values removed:  $removedCount" -ForegroundColor Cyan
    Write-Host "  Errors:                   $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { 'Yellow' } else { 'Cyan' })
}

# ===========================================================================
#  Get-W11ThemeBackups
# ===========================================================================
function Get-W11ThemeBackups {
    <#
    .SYNOPSIS
        Lists all available theme backups.
    .DESCRIPTION
        Scans the project backups folder for directories containing a valid
        backup-manifest.json and returns their metadata.
    #>
    [CmdletBinding()]
    param()

    $backupsRoot = Join-Path $script:ProjectRoot 'backups'

    if (-not (Test-Path $backupsRoot)) {
        Write-Warning "Backups directory not found: '$backupsRoot'."
        return @()
    }

    $results = @()
    $dirs = Get-ChildItem -Path $backupsRoot -Directory -ErrorAction SilentlyContinue

    foreach ($dir in $dirs) {
        $manifestFile = Join-Path $dir.FullName 'backup-manifest.json'
        if (Test-Path $manifestFile) {
            try {
                $manifest = Get-Content -Path $manifestFile -Raw | ConvertFrom-Json
                $results += [PSCustomObject]@{
                    Name      = $manifest.Name
                    Timestamp = $manifest.Timestamp
                    OSBuild   = $manifest.OSBuild
                    ThemePath = $manifest.CurrentThemePath
                }
            }
            catch {
                Write-Warning "Could not parse manifest in '$($dir.Name)': $_"
            }
        }
    }

    return $results
}

# ===========================================================================
#  Remove-W11ThemeBackup
# ===========================================================================
function Remove-W11ThemeBackup {
    <#
    .SYNOPSIS
        Deletes a named theme backup.
    .DESCRIPTION
        Removes the backup folder and all its contents after user confirmation
        (unless -Force is specified).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [switch]$Force
    )

    $backupDir = Join-Path $script:ProjectRoot "backups\$Name"

    if (-not (Test-Path $backupDir)) {
        Write-Error "Backup '$Name' not found at '$backupDir'."
        return
    }

    # Confirm with the user unless -Force is set
    if (-not $Force) {
        $confirm = $PSCmdlet.ShouldProcess(
            "Backup '$Name' at '$backupDir'",
            "Permanently delete"
        )
        if (-not $confirm) {
            Write-Host "Removal cancelled." -ForegroundColor Yellow
            return
        }
    }

    try {
        Remove-Item -Path $backupDir -Recurse -Force -ErrorAction Stop
        Write-Host "Backup '$Name' has been removed." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to remove backup '$Name': $_"
    }
}
