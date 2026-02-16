#Requires -Version 5.1
<#
.SYNOPSIS
    ThemeFileBuilder module — generates valid Windows 11 .theme files (INI format)
    from a theme configuration PSCustomObject.
#>

# ---------------------------------------------------------------------------
# Internal Helper: Format-IniSection
# ---------------------------------------------------------------------------
function Format-IniSection {
    <#
    .SYNOPSIS
        Formats a single INI section as a string block.
    .PARAMETER SectionName
        The name that appears inside square brackets, e.g. "Theme".
    .PARAMETER Entries
        An ordered hashtable of Key=Value pairs for this section.
        Entries whose value is $null are silently skipped.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SectionName,

        [Parameter(Mandatory)]
        [System.Collections.Specialized.OrderedDictionary]$Entries
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("[$SectionName]")

    foreach ($key in $Entries.Keys) {
        $value = $Entries[$key]
        if ($null -ne $value) {
            $lines.Add("$key=$value")
        }
    }

    return ($lines -join "`r`n")
}


# ---------------------------------------------------------------------------
# Internal Helper: Get-DefaultVisualStylesPath
# ---------------------------------------------------------------------------
function Get-DefaultVisualStylesPath {
    <#
    .SYNOPSIS
        Returns the default Windows 11 Aero visual style path.
    #>
    return '%SystemRoot%\resources\Themes\Aero\Aero.msstyles'
}

# ---------------------------------------------------------------------------
# Exported Function: New-W11ThemeFile
# ---------------------------------------------------------------------------
function New-W11ThemeFile {
    <#
    .SYNOPSIS
        Generates a valid Windows 11 .theme file from a configuration object.
    .DESCRIPTION
        Accepts a PSCustomObject with theme metadata, colors, wallpaper,
        cursors, and visual-style settings, then writes a properly formatted
        INI-based .theme file that Windows 11 can apply directly.
    .PARAMETER Config
        A PSCustomObject containing the theme definition.  Expected members:
          .meta.name          — display name of the theme
          .colors             — hashtable/object of colour triplets (R G B)
          .wallpaper          — wallpaper settings (path, mode, style, etc.)
          .cursors            — cursor set definition
          .visualStyles       — visual-style overrides
    .PARAMETER OutputPath
        Optional explicit output file path.  When omitted the file is written
        to  <ProjectRoot>\output\<ThemeName>.theme
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [Parameter()]
        [string]$OutputPath
    )

    # -----------------------------------------------------------------------
    # Resolve project root (two levels up from this module's directory)
    # -----------------------------------------------------------------------
    $ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..') -ErrorAction Stop).Path

    # -----------------------------------------------------------------------
    # Determine output path
    # -----------------------------------------------------------------------
    if (-not $OutputPath) {
        $safeName = $Config.meta.name -replace '\s', '_'
        $outputDir = Join-Path $ProjectRoot 'output'
        if (-not (Test-Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        }
        $OutputPath = Join-Path $outputDir "$safeName.theme"
    }

    # Collect all sections — each element is a formatted INI string
    $sections = [System.Collections.Generic.List[string]]::new()

    # -----------------------------------------------------------------------
    # [Theme] section
    # -----------------------------------------------------------------------
    $themeEntries = [ordered]@{
        DisplayName = $Config.meta.name
    }
    $sections.Add((Format-IniSection -SectionName 'Theme' -Entries $themeEntries))

    # -----------------------------------------------------------------------
    # [Control Panel\Colors] section
    # -----------------------------------------------------------------------
    # Mapping from config property names to INI key names
    $colorKeyMap = [ordered]@{
        activeTitle          = 'ActiveTitle'
        background           = 'Background'
        hilight              = 'Hilight'
        hilightText          = 'HilightText'
        titleText            = 'TitleText'
        window               = 'Window'
        windowText           = 'WindowText'
        scrollbar            = 'Scrollbar'
        inactiveTitle        = 'InactiveTitle'
        menu                 = 'Menu'
        menuText             = 'MenuText'
        activeBorder         = 'ActiveBorder'
        inactiveBorder       = 'InactiveBorder'
        appWorkspace         = 'AppWorkSpace'
        buttonFace           = 'ButtonFace'
        buttonShadow         = 'ButtonShadow'
        grayText             = 'GrayText'
        buttonText           = 'ButtonText'
        inactiveTitleText    = 'InactiveTitleText'
        buttonHilight        = 'ButtonHilight'
        buttonDkShadow       = 'ButtonDkShadow'
        buttonLight          = 'ButtonLight'
        infoText             = 'InfoText'
        infoWindow           = 'InfoWindow'
        gradientActiveTitle  = 'GradientActiveTitle'
        gradientInactiveTitle = 'GradientInactiveTitle'
    }

    # Sensible dark-theme defaults (RGB triplets as strings)
    $darkDefaults = [ordered]@{
        ActiveTitle          = '36 36 36'
        Background           = '30 30 30'
        Hilight              = '0 120 215'
        HilightText          = '255 255 255'
        TitleText            = '255 255 255'
        Window               = '32 32 32'
        WindowText           = '255 255 255'
        Scrollbar            = '50 50 50'
        InactiveTitle        = '44 44 44'
        Menu                 = '36 36 36'
        MenuText             = '255 255 255'
        ActiveBorder         = '48 48 48'
        InactiveBorder       = '44 44 44'
        AppWorkSpace         = '30 30 30'
        ButtonFace           = '51 51 51'
        ButtonShadow         = '25 25 25'
        GrayText             = '128 128 128'
        ButtonText           = '255 255 255'
        InactiveTitleText    = '160 160 160'
        ButtonHilight        = '70 70 70'
        ButtonDkShadow       = '15 15 15'
        ButtonLight          = '60 60 60'
        InfoText             = '255 255 255'
        InfoWindow           = '45 45 45'
        GradientActiveTitle  = '36 36 36'
        GradientInactiveTitle = '44 44 44'
    }

    $colorEntries = [ordered]@{}

    if ($null -ne $Config.colors) {
        # Build entries from config, using the key map for proper INI names
        foreach ($configKey in $colorKeyMap.Keys) {
            $iniKey = $colorKeyMap[$configKey]
            $value  = $Config.colors.$configKey
            if ($null -ne $value) {
                $colorEntries[$iniKey] = $value
            } else {
                # Fall back to dark default for missing keys
                $colorEntries[$iniKey] = $darkDefaults[$iniKey]
            }
        }
    } else {
        # No colors in config — use all dark defaults
        $colorEntries = $darkDefaults
    }

    $sections.Add((Format-IniSection -SectionName 'Control Panel\Colors' -Entries $colorEntries))

    # -----------------------------------------------------------------------
    # [Control Panel\Desktop] section (wallpaper)
    # -----------------------------------------------------------------------
    $wpPath  = ''
    $wpTile  = '0'
    $wpStyle = '10'  # Default: Fill

    if ($null -ne $Config.wallpaper) {
        $rawPath = $Config.wallpaper.path
        if ($rawPath) {
            # Resolve relative paths against the project assets folder
            if ($rawPath -notmatch '^[A-Za-z]:\\') {
                $wpPath = Join-Path $ProjectRoot "assets\wallpapers\$rawPath"
            } else {
                $wpPath = $rawPath
            }
        }

        # TileWallpaper flag
        if ($Config.wallpaper.tile) {
            $wpTile = '1'
        }

        # WallpaperStyle: Center=0, Stretch=2, Fit=6, Fill=10
        if ($Config.wallpaper.style) {
            $wpStyle = switch ($Config.wallpaper.style) {
                'center'  { '0' }
                'stretch' { '2' }
                'fit'     { '6' }
                'fill'    { '10' }
                default   { $Config.wallpaper.style }
            }
        }
    }

    $desktopEntries = [ordered]@{
        Wallpaper      = $wpPath
        TileWallpaper  = $wpTile
        WallpaperStyle = $wpStyle
    }
    $sections.Add((Format-IniSection -SectionName 'Control Panel\Desktop' -Entries $desktopEntries))

    # -----------------------------------------------------------------------
    # [Slideshow] section — only when wallpaper mode is 'slideshow'
    # -----------------------------------------------------------------------
    if ($Config.wallpaper -and $Config.wallpaper.mode -eq 'slideshow') {
        $slideshowEntries = [ordered]@{}

        if ($Config.wallpaper.intervalMs) {
            $slideshowEntries['Interval'] = $Config.wallpaper.intervalMs
        }

        $slideshowEntries['Shuffle'] = if ($Config.wallpaper.shuffle) { '1' } else { '0' }

        # Resolve the images root path
        $imagesRoot = $Config.wallpaper.imagesRootPath
        if ($imagesRoot -and $imagesRoot -notmatch '^[A-Za-z]:\\') {
            $imagesRoot = Join-Path $ProjectRoot "assets\wallpapers\$imagesRoot"
        }
        if ($imagesRoot) {
            $slideshowEntries['ImagesRootPath'] = $imagesRoot
        }

        # Add individual item paths (Item0Path, Item1Path, ...)
        if ($Config.wallpaper.items) {
            for ($i = 0; $i -lt $Config.wallpaper.items.Count; $i++) {
                $itemPath = $Config.wallpaper.items[$i]
                if ($itemPath -notmatch '^[A-Za-z]:\\') {
                    $itemPath = Join-Path $ProjectRoot "assets\wallpapers\$itemPath"
                }
                $slideshowEntries["Item${i}Path"] = $itemPath
            }
        }

        $sections.Add((Format-IniSection -SectionName 'Slideshow' -Entries $slideshowEntries))
    }

    # -----------------------------------------------------------------------
    # [Control Panel\Cursors] section
    # -----------------------------------------------------------------------
    if ($null -ne $Config.cursors) {
        # Mapping from config cursor role names to INI key names
        $cursorKeyMap = [ordered]@{
            Arrow       = 'Arrow'
            Help        = 'Help'
            AppStarting = 'AppStarting'
            Wait        = 'Wait'
            Crosshair   = 'Crosshair'
            IBeam       = 'IBeam'
            NWPen       = 'NWPen'
            No          = 'No'
            SizeNS      = 'SizeNS'
            SizeWE      = 'SizeWE'
            SizeNWSE    = 'SizeNWSE'
            SizeNESW    = 'SizeNESW'
            SizeAll     = 'SizeAll'
            UpArrow     = 'UpArrow'
            Hand        = 'Hand'
        }

        $cursorEntries = [ordered]@{}

        # Scheme name metadata
        $schemeName = if ($Config.cursors.schemeName) { $Config.cursors.schemeName } else { $Config.meta.name }
        $cursorEntries['DefaultValue']     = $schemeName
        $cursorEntries['DefaultValue.MUI'] = $schemeName

        # Resolve each cursor file path
        $setFolder = $Config.cursors.setFolder
        foreach ($role in $cursorKeyMap.Keys) {
            $iniKey   = $cursorKeyMap[$role]
            $fileName = $Config.cursors.$role
            if ($fileName) {
                $cursorEntries[$iniKey] = Join-Path $ProjectRoot "assets\cursors\$setFolder\$fileName"
            }
        }

        $sections.Add((Format-IniSection -SectionName 'Control Panel\Cursors' -Entries $cursorEntries))
    }

    # -----------------------------------------------------------------------
    # [VisualStyles] section
    # -----------------------------------------------------------------------
    $vsPath = if ($Config.visualStyles -and $Config.visualStyles.path) {
        $Config.visualStyles.path
    } else {
        Get-DefaultVisualStylesPath
    }

    $colorizationColor = if ($Config.visualStyles -and $Config.visualStyles.colorizationColor) {
        $Config.visualStyles.colorizationColor
    } else {
        '0XC40078D7'  # Default Windows 11 accent blue
    }

    $colorStyle = if ($Config.visualStyles -and $Config.visualStyles.colorStyle) {
        $Config.visualStyles.colorStyle
    } else {
        'NormalColor'
    }

    $size = if ($Config.visualStyles -and $Config.visualStyles.size) {
        $Config.visualStyles.size
    } else {
        'NormalSize'
    }

    $transparency = if ($Config.visualStyles -and $null -ne $Config.visualStyles.transparency) {
        $Config.visualStyles.transparency
    } else {
        '1'
    }

    $composition = if ($Config.visualStyles -and $null -ne $Config.visualStyles.composition) {
        $Config.visualStyles.composition
    } else {
        '1'
    }

    $vsEntries = [ordered]@{
        Path              = $vsPath
        ColorStyle        = $colorStyle
        Size              = $size
        ColorizationColor = $colorizationColor
        Transparency      = $transparency
        Composition       = $composition
    }
    $sections.Add((Format-IniSection -SectionName 'VisualStyles' -Entries $vsEntries))

    # -----------------------------------------------------------------------
    # [boot] section
    # -----------------------------------------------------------------------
    $bootEntries = [ordered]@{
        'SCRNSAVE.EXE' = ''
    }
    $sections.Add((Format-IniSection -SectionName 'boot' -Entries $bootEntries))

    # -----------------------------------------------------------------------
    # [MasterThemeSelector] section
    # -----------------------------------------------------------------------
    $mtsEntries = [ordered]@{
        MTSM = 'DABJDKT'
    }
    $sections.Add((Format-IniSection -SectionName 'MasterThemeSelector' -Entries $mtsEntries))

    # -----------------------------------------------------------------------
    # Assemble final content and write to disk
    # -----------------------------------------------------------------------
    $iniContent = $sections -join "`r`n`r`n"

    try {
        # Ensure the output directory exists
        $outDir = Split-Path $OutputPath -Parent
        if (-not (Test-Path $outDir)) {
            New-Item -Path $outDir -ItemType Directory -Force | Out-Null
        }

        # Windows .theme files require Unicode (UTF-16 LE) encoding
        Set-Content -Path $OutputPath -Value $iniContent -Encoding Unicode -Force
        Write-Verbose "Theme file written to: $OutputPath"
    }
    catch {
        Write-Error "Failed to write theme file to '$OutputPath': $_"
        return $null
    }

    # Return the full path of the generated file
    return $OutputPath
}

# ---------------------------------------------------------------------------
# Module exports
# ---------------------------------------------------------------------------
Export-ModuleMember -Function 'New-W11ThemeFile'
