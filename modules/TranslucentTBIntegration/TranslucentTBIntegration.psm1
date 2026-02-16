#Requires -Version 5.1
<#
.SYNOPSIS
    TranslucentTB integration module for w11-theming-suite.

.DESCRIPTION
    Detects TranslucentTB installation (Store or standalone), reads and writes
    its settings.json config, and allows theme presets to include TranslucentTB
    appearance settings that are applied automatically.
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Internal: Find TranslucentTB config path
# ---------------------------------------------------------------------------
function Find-TranslucentTBConfigPath {
    [CmdletBinding()]
    [OutputType([string])]
    param()

    # 1. Microsoft Store version (most common on Windows 11)
    $storePattern = Join-Path $env:LOCALAPPDATA 'Packages\*TranslucentTB*'
    $storeDirs = Get-ChildItem -Path $storePattern -Directory -ErrorAction SilentlyContinue
    foreach ($dir in $storeDirs) {
        $configPath = Join-Path $dir.FullName 'RoamingState\settings.json'
        if (Test-Path $configPath) {
            Write-Verbose "Found Store TranslucentTB config: $configPath"
            return $configPath
        }
    }

    # 2. Standalone / GitHub release version
    $standalonePaths = @(
        (Join-Path $env:LOCALAPPDATA 'TranslucentTB\settings.json'),
        (Join-Path $env:LOCALAPPDATA 'translucenttb\settings.json'),
        (Join-Path $env:APPDATA 'TranslucentTB\settings.json')
    )
    foreach ($p in $standalonePaths) {
        if (Test-Path $p) {
            Write-Verbose "Found standalone TranslucentTB config: $p"
            return $p
        }
    }

    return $null
}

# ---------------------------------------------------------------------------
# Internal: Check if TranslucentTB process is running
# ---------------------------------------------------------------------------
function Test-TranslucentTBRunning {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $proc = Get-Process -Name 'TranslucentTB' -ErrorAction SilentlyContinue
    return ($null -ne $proc)
}

# ===========================================================================
#  Public: Test-W11TranslucentTBInstalled
# ===========================================================================
function Test-W11TranslucentTBInstalled {
    <#
    .SYNOPSIS
        Checks if TranslucentTB is installed and returns installation details.

    .DESCRIPTION
        Detects TranslucentTB by looking for its config file in known locations
        (Microsoft Store and standalone). Also checks if it is currently running.

    .OUTPUTS
        PSCustomObject with Installed, Running, ConfigPath, InstallType properties.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()

    $configPath = Find-TranslucentTBConfigPath
    $installed = $null -ne $configPath
    $running = Test-TranslucentTBRunning

    $installType = if (-not $installed) {
        'NotFound'
    } elseif ($configPath -match 'Packages.*TranslucentTB') {
        'MicrosoftStore'
    } else {
        'Standalone'
    }

    return [PSCustomObject]@{
        Installed   = $installed
        Running     = $running
        ConfigPath  = $configPath
        InstallType = $installType
    }
}

# ===========================================================================
#  Public: Get-W11TranslucentTBConfig
# ===========================================================================
function Get-W11TranslucentTBConfig {
    <#
    .SYNOPSIS
        Reads the current TranslucentTB configuration.

    .DESCRIPTION
        Parses the TranslucentTB settings.json and returns it as a PSCustomObject.
        Returns $null if TranslucentTB is not installed.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()

    $configPath = Find-TranslucentTBConfigPath
    if (-not $configPath) {
        Write-Warning 'TranslucentTB is not installed or config not found.'
        return $null
    }

    try {
        # Read and strip // comments (TranslucentTB uses JS-style comments)
        $rawContent = Get-Content -Path $configPath -Raw -ErrorAction Stop
        $cleanContent = ($rawContent -split "`n" | ForEach-Object {
            # Strip single-line // comments but not inside strings (e.g. URLs)
            # Only strip // at the start of a line (with optional whitespace)
            if ($_ -match '^\s*//') { '' } else { $_ }
        }) -join "`n"
        $config = $cleanContent | ConvertFrom-Json -ErrorAction Stop
        Write-Verbose "Loaded TranslucentTB config from: $configPath"
        return $config
    }
    catch {
        Write-Error "Failed to parse TranslucentTB config at '$configPath': $_"
        return $null
    }
}

# ===========================================================================
#  Public: Set-W11TranslucentTBConfig
# ===========================================================================
function Set-W11TranslucentTBConfig {
    <#
    .SYNOPSIS
        Applies TranslucentTB settings from a theme config.

    .DESCRIPTION
        Takes the translucentTB section from a theme config and merges it
        into the existing TranslucentTB settings.json. TranslucentTB auto-
        reloads when its config file changes on disk.

    .PARAMETER Config
        The full theme PSCustomObject containing an advanced.translucentTB section.

    .PARAMETER Settings
        Direct TranslucentTB settings hashtable to apply (alternative to Config).

    .PARAMETER Force
        Overwrite even if TranslucentTB is not running.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $false)]
        [PSCustomObject]$Config,

        [Parameter(Mandatory = $false)]
        [hashtable]$Settings,

        [switch]$Force
    )

    # Determine the settings to apply
    $ttbSettings = $null
    if ($Settings) {
        $ttbSettings = $Settings
    }
    elseif ($Config -and $Config.advanced -and $Config.advanced.translucentTB -and $Config.advanced.translucentTB.config) {
        $ttbSettings = $Config.advanced.translucentTB.config
    }

    if (-not $ttbSettings) {
        Write-Verbose 'No TranslucentTB settings found in config. Skipping.'
        return
    }

    # Find config path
    $configPath = Find-TranslucentTBConfigPath
    if (-not $configPath) {
        Write-Warning 'TranslucentTB is not installed. Cannot apply TranslucentTB settings.'
        return
    }

    # Check if running
    $running = Test-TranslucentTBRunning
    if (-not $running -and -not $Force) {
        Write-Warning 'TranslucentTB is not running. Settings will be applied but take effect on next launch. Use -Force to suppress this warning.'
    }

    if (-not $PSCmdlet.ShouldProcess($configPath, 'Update TranslucentTB settings')) {
        return
    }

    try {
        # Load current config
        $rawContent = Get-Content -Path $configPath -Raw -ErrorAction Stop
        $cleanContent = ($rawContent -split "`n" | ForEach-Object {
            # Strip single-line // comments but not inside strings (e.g. URLs)
            # Only strip // at the start of a line (with optional whitespace)
            if ($_ -match '^\s*//') { '' } else { $_ }
        }) -join "`n"
        $currentConfig = $cleanContent | ConvertFrom-Json -ErrorAction Stop

        # Merge settings: iterate provided keys and overwrite in current config
        $ttbSettingsObj = if ($ttbSettings -is [hashtable]) {
            # Convert hashtable to PSCustomObject for uniform handling
            [PSCustomObject]$ttbSettings
        } else {
            $ttbSettings
        }

        foreach ($prop in $ttbSettingsObj.PSObject.Properties) {
            $sectionName = $prop.Name
            $sectionValue = $prop.Value

            if ($null -ne $currentConfig.PSObject.Properties[$sectionName]) {
                # Merge sub-properties
                if ($sectionValue -is [PSCustomObject] -or $sectionValue -is [hashtable]) {
                    $sectionObj = if ($sectionValue -is [hashtable]) { [PSCustomObject]$sectionValue } else { $sectionValue }
                    foreach ($subProp in $sectionObj.PSObject.Properties) {
                        $currentConfig.$sectionName | Add-Member -NotePropertyName $subProp.Name -NotePropertyValue $subProp.Value -Force
                    }
                }
                else {
                    $currentConfig | Add-Member -NotePropertyName $sectionName -NotePropertyValue $sectionValue -Force
                }
            }
            else {
                $currentConfig | Add-Member -NotePropertyName $sectionName -NotePropertyValue $sectionValue -Force
            }
        }

        # Write back with schema comment header
        $jsonOutput = "// See https://TranslucentTB.github.io/config for more information`r`n"
        $jsonOutput += ($currentConfig | ConvertTo-Json -Depth 10)

        Set-Content -Path $configPath -Value $jsonOutput -Encoding UTF8 -ErrorAction Stop
        Write-Host "  TranslucentTB config updated: $configPath" -ForegroundColor Green

        if ($running) {
            Write-Host '  TranslucentTB will auto-reload the new settings.' -ForegroundColor Gray
        }
    }
    catch {
        Write-Error "Failed to update TranslucentTB config: $_"
    }
}

# ---------------------------------------------------------------------------
# Export public functions
# ---------------------------------------------------------------------------
Export-ModuleMember -Function @(
    'Get-W11TranslucentTBConfig',
    'Set-W11TranslucentTBConfig',
    'Test-W11TranslucentTBInstalled'
)
