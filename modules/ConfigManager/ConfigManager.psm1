#Requires -Version 5.1

<#
.SYNOPSIS
    ConfigManager module for w11-theming-suite.
.DESCRIPTION
    Provides JSON-based configuration loading, validation, preset inheritance
    (via meta.basedOn), interactive config creation, and deep-merge utilities.
#>

# ---------------------------------------------------------------------------
# Module-scoped variables
# ---------------------------------------------------------------------------

# Project root is two levels up from this module directory
# $PSScriptRoot = modules\ConfigManager  ->  ..\..\  = project root
$script:ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path

# ---------------------------------------------------------------------------
# Internal helper: Test-W11ThemeConfigSchema
# ---------------------------------------------------------------------------

function Test-W11ThemeConfigSchema {
    <#
    .SYNOPSIS
        Validates that a config object contains all required fields.
    .DESCRIPTION
        Checks for the presence of meta.name, meta.version, and mode.
        Returns $true when valid; writes a non-terminating error and
        returns $false when a required field is missing.
    .PARAMETER Config
        The PSCustomObject loaded from a theme JSON file.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config
    )

    $valid = $true

    # --- meta block ---
    if (-not $Config.PSObject.Properties['meta']) {
        Write-Error 'Config validation failed: missing top-level "meta" object.'
        return $false
    }

    if (-not $Config.meta.PSObject.Properties['name'] -or [string]::IsNullOrWhiteSpace($Config.meta.name)) {
        Write-Error 'Config validation failed: "meta.name" is required.'
        $valid = $false
    }

    if (-not $Config.meta.PSObject.Properties['version'] -or [string]::IsNullOrWhiteSpace($Config.meta.version)) {
        Write-Error 'Config validation failed: "meta.version" is required.'
        $valid = $false
    }

    # --- mode ---
    if (-not $Config.PSObject.Properties['mode'] -or [string]::IsNullOrWhiteSpace($Config.mode)) {
        Write-Error 'Config validation failed: "mode" is required (dark or light).'
        $valid = $false
    }

    return $valid
}

# ---------------------------------------------------------------------------
# Internal helper: Resolve-W11AssetPath
# ---------------------------------------------------------------------------

function Resolve-W11AssetPath {
    <#
    .SYNOPSIS
        Converts a relative asset path to an absolute path.
    .DESCRIPTION
        If the supplied path is already rooted it is returned unchanged.
        Otherwise the path is resolved relative to the given BasePath
        (defaults to $script:ProjectRoot).
    .PARAMETER RelativePath
        The path value read from the JSON config (may be relative or absolute).
    .PARAMETER BasePath
        The base directory to resolve relative paths against.
        Defaults to $script:ProjectRoot.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$RelativePath,

        [Parameter()]
        [string]$BasePath = $script:ProjectRoot
    )

    # Nothing to resolve when the value is empty or null
    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        return $RelativePath
    }

    # Already an absolute path – return as-is
    if ([System.IO.Path]::IsPathRooted($RelativePath)) {
        return $RelativePath
    }

    # Build absolute path from the base path
    $absolute = Join-Path $BasePath $RelativePath
    return $absolute
}

# ---------------------------------------------------------------------------
# Private helper: Resolve asset paths inside a loaded config object
# ---------------------------------------------------------------------------

function Resolve-AllAssetPaths {
    <#
    .SYNOPSIS
        Walks known asset-path properties and resolves them to absolute paths.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config
    )

    $assetsRoot = Join-Path $script:ProjectRoot 'assets'

    # cursors.setFolder – relative to assets/cursors/
    if ($Config.PSObject.Properties['cursors'] -and
        $null -ne $Config.cursors -and
        $Config.cursors.PSObject.Properties['setFolder'] -and
        -not [string]::IsNullOrWhiteSpace($Config.cursors.setFolder)) {
        $Config.cursors.setFolder = Resolve-W11AssetPath -RelativePath $Config.cursors.setFolder -BasePath (Join-Path $assetsRoot 'cursors')
    }

    # sounds.setFolder – relative to assets/sounds/
    if ($Config.PSObject.Properties['sounds'] -and
        $null -ne $Config.sounds -and
        $Config.sounds.PSObject.Properties['setFolder'] -and
        -not [string]::IsNullOrWhiteSpace($Config.sounds.setFolder)) {
        $Config.sounds.setFolder = Resolve-W11AssetPath -RelativePath $Config.sounds.setFolder -BasePath (Join-Path $assetsRoot 'sounds')
    }

    # wallpaper.path – relative to assets/wallpapers/
    if ($Config.PSObject.Properties['wallpaper'] -and
        $null -ne $Config.wallpaper -and
        $Config.wallpaper.PSObject.Properties['path'] -and
        -not [string]::IsNullOrWhiteSpace($Config.wallpaper.path)) {
        $Config.wallpaper.path = Resolve-W11AssetPath -RelativePath $Config.wallpaper.path -BasePath (Join-Path $assetsRoot 'wallpapers')
    }

    return $Config
}

# ---------------------------------------------------------------------------
# Public: Merge-W11ThemeConfig
# ---------------------------------------------------------------------------

function Merge-W11ThemeConfig {
    <#
    .SYNOPSIS
        Deep-merges two PSCustomObjects (Base + Override).
    .DESCRIPTION
        Recursively walks every property. Rules:
          - Non-null Override values replace Base values.
          - Null Override values are skipped (Base value kept).
          - Arrays in Override replace Base arrays entirely.
          - Nested PSCustomObjects are merged recursively.
    .PARAMETER Base
        The base / parent configuration object.
    .PARAMETER Override
        The override / child configuration object whose values take precedence.
    .OUTPUTS
        [PSCustomObject] The merged result.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Base,

        [Parameter(Mandatory)]
        [PSCustomObject]$Override
    )

    # Start with a shallow clone of Base so we don't mutate the original
    $merged = [PSCustomObject]@{}
    foreach ($prop in $Base.PSObject.Properties) {
        $merged | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value
    }

    # Walk every property in Override
    foreach ($prop in $Override.PSObject.Properties) {
        $overrideValue = $prop.Value

        # Skip null override values – keep base
        if ($null -eq $overrideValue) {
            continue
        }

        $baseHasProp = $merged.PSObject.Properties[$prop.Name]

        if ($null -ne $baseHasProp -and
            $baseHasProp.Value -is [PSCustomObject] -and
            $overrideValue -is [PSCustomObject]) {
            # Both sides are objects – recurse
            $merged.($prop.Name) = Merge-W11ThemeConfig -Base $baseHasProp.Value -Override $overrideValue
        }
        elseif ($null -ne $baseHasProp) {
            # Override replaces base (including arrays)
            $merged.($prop.Name) = $overrideValue
        }
        else {
            # Property only exists in Override – add it
            $merged | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $overrideValue
        }
    }

    return $merged
}

# ---------------------------------------------------------------------------
# Public: Get-W11ThemeConfig
# ---------------------------------------------------------------------------

function Get-W11ThemeConfig {
    <#
    .SYNOPSIS
        Loads and validates a w11-theming-suite JSON configuration file.
    .DESCRIPTION
        Accepts either a direct -Path to a JSON file or a -PresetName that is
        resolved to config\presets\<PresetName>.json under the project root.

        When the loaded config contains meta.basedOn, the parent preset is
        loaded recursively and deep-merged (child overrides parent).

        Relative asset paths (cursors.setFolder, sounds.setFolder,
        wallpaper.path) are resolved to absolute paths.
    .PARAMETER Path
        Full or relative path to a JSON config file.
    .PARAMETER PresetName
        Name of a preset (without extension). Resolves to
        <ProjectRoot>\config\presets\<PresetName>.json.
    .OUTPUTS
        [PSCustomObject] The fully resolved theme configuration.
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByPath')]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'ByPath', Position = 0)]
        [string]$Path,

        [Parameter(Mandatory, ParameterSetName = 'ByPreset')]
        [string]$PresetName
    )

    # Resolve the JSON file path
    if ($PSCmdlet.ParameterSetName -eq 'ByPreset') {
        $Path = Join-Path $script:ProjectRoot "config\presets\$PresetName.json"
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Error "Configuration file not found: $Path"
        return $null
    }

    # Load and parse JSON
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $config = $raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to parse JSON from '$Path': $_"
        return $null
    }

    # Validate required schema
    if (-not (Test-W11ThemeConfigSchema -Config $config)) {
        return $null
    }

    # Handle preset inheritance via meta.basedOn
    if ($config.PSObject.Properties['meta'] -and
        $config.meta.PSObject.Properties['basedOn'] -and
        -not [string]::IsNullOrWhiteSpace($config.meta.basedOn)) {

        $parentName = $config.meta.basedOn
        Write-Verbose "Inheriting from parent preset: $parentName"

        $parentConfig = Get-W11ThemeConfig -PresetName $parentName

        if ($null -eq $parentConfig) {
            Write-Error "Failed to load parent preset '$parentName' referenced by basedOn in '$Path'."
            return $null
        }

        # Deep-merge: child overrides parent
        $config = Merge-W11ThemeConfig -Base $parentConfig -Override $config
    }

    # Resolve relative asset paths to absolute
    $config = Resolve-AllAssetPaths -Config $config

    return $config
}

# ---------------------------------------------------------------------------
# Public: New-W11ThemeConfig
# ---------------------------------------------------------------------------

function New-W11ThemeConfig {
    <#
    .SYNOPSIS
        Interactive wizard that creates a new w11-theming-suite configuration file.
    .DESCRIPTION
        Prompts the user for basic theme properties (name, author, mode,
        accent color, wallpaper path) and generates a valid JSON config
        with all standard sections.  Unset sections default to $null so
        they can be populated later.
    .PARAMETER OutputPath
        Optional. Full path where the JSON file will be saved. When omitted
        the file is saved to config\user\<ThemeName>.json under the project root.
    .OUTPUTS
        [string] The path to the created configuration file.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter()]
        [string]$OutputPath
    )

    # ---- Interactive prompts ----
    $themeName = Read-Host -Prompt 'Theme name'
    if ([string]::IsNullOrWhiteSpace($themeName)) {
        Write-Error 'Theme name cannot be empty.'
        return $null
    }

    $author     = Read-Host -Prompt 'Author'
    $modeInput  = Read-Host -Prompt 'Mode (dark/light)'
    $mode       = if ($modeInput -match '^(dark|light)$') { $modeInput.ToLower() } else { 'dark' }
    $accentHex  = Read-Host -Prompt 'Accent color hex (e.g. #0078D4)'
    $wallpaper  = Read-Host -Prompt 'Wallpaper path (relative or absolute, leave blank to skip)'

    # ---- Build config object ----
    $config = [ordered]@{
        meta      = [ordered]@{
            name    = $themeName
            version = '1.0.0'
            author  = $author
            basedOn = $null
        }
        mode      = $mode
        accent    = [ordered]@{
            color = $accentHex
        }
        wallpaper = [ordered]@{
            path = if ([string]::IsNullOrWhiteSpace($wallpaper)) { $null } else { $wallpaper }
        }
        cursors   = [ordered]@{
            setFolder = $null
            size      = $null
        }
        sounds    = [ordered]@{
            setFolder = $null
        }
        taskbar   = $null
        explorer  = $null
    }

    # ---- Determine output path ----
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        # Sanitize the theme name for use as a filename
        $safeName   = $themeName -replace '[\\/:*?"<>|]', '_'
        $OutputPath = Join-Path $script:ProjectRoot "config\user\$safeName.json"
    }

    # Ensure the parent directory exists
    $parentDir = Split-Path -Path $OutputPath -Parent
    if (-not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    # ---- Write JSON ----
    try {
        $json = $config | ConvertTo-Json -Depth 10
        Set-Content -LiteralPath $OutputPath -Value $json -Encoding UTF8 -ErrorAction Stop
        Write-Verbose "Configuration saved to $OutputPath"
    }
    catch {
        Write-Error "Failed to write configuration to '$OutputPath': $_"
        return $null
    }

    return $OutputPath
}

# ---------------------------------------------------------------------------
# Export public functions
# ---------------------------------------------------------------------------

Export-ModuleMember -Function @(
    'Get-W11ThemeConfig',
    'New-W11ThemeConfig',
    'Merge-W11ThemeConfig'
)
