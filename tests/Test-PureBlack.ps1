# Test-PureBlack.ps1 - Integration tests for pure-black preset + TranslucentTB
$ErrorActionPreference = 'Continue'
$passed = 0; $failed = 0; $total = 0

function Test-Assert {
    param([string]$Name, [scriptblock]$Test)
    $script:total++
    try {
        $result = & $Test
        if ($result) {
            Write-Host "  PASS: $Name" -ForegroundColor Green
            $script:passed++
        } else {
            Write-Host "  FAIL: $Name" -ForegroundColor Red
            $script:failed++
        }
    } catch {
        Write-Host "  FAIL: $Name - $_" -ForegroundColor Red
        $script:failed++
    }
}

Write-Host "=== Pure Black + TranslucentTB Integration Tests ===" -ForegroundColor Cyan
Write-Host ""

# Import module
$modulePath = 'C:\Dev\w11-theming-suite\w11-theming-suite.psd1'
Import-Module $modulePath -Force -DisableNameChecking

# Test 1: Module loads TranslucentTB functions
Test-Assert "Test-W11TranslucentTBInstalled exported" {
    $null -ne (Get-Command Test-W11TranslucentTBInstalled -EA SilentlyContinue)
}

Test-Assert "Get-W11TranslucentTBConfig exported" {
    $null -ne (Get-Command Get-W11TranslucentTBConfig -EA SilentlyContinue)
}

Test-Assert "Set-W11TranslucentTBConfig exported" {
    $null -ne (Get-Command Set-W11TranslucentTBConfig -EA SilentlyContinue)
}

# Test 2: Load pure-black config
Test-Assert "Pure-black config loads successfully" {
    $cfg = Get-W11ThemeConfig -PresetName 'pure-black'
    $cfg.meta.name -eq 'Pure Black'
}

# Test 3: Pure-black has TranslucentTB section
Test-Assert "Pure-black contains translucentTB config" {
    $cfg = Get-W11ThemeConfig -PresetName 'pure-black'
    $null -ne $cfg.advanced.translucentTB.config
}

# Test 4: TranslucentTB config has clear accent
Test-Assert "TranslucentTB desktop_appearance.accent is 'clear'" {
    $cfg = Get-W11ThemeConfig -PresetName 'pure-black'
    $cfg.advanced.translucentTB.config.desktop_appearance.accent -eq 'clear'
}

# Test 5: Pure-black dark mode settings correct
Test-Assert "Dark mode apps=0, system=0" {
    $cfg = Get-W11ThemeConfig -PresetName 'pure-black'
    $cfg.mode.appsUseLightTheme -eq 0 -and $cfg.mode.systemUsesLightTheme -eq 0
}

# Test 6: Pure-black accent color is black
Test-Assert "Accent color is #000000" {
    $cfg = Get-W11ThemeConfig -PresetName 'pure-black'
    $cfg.accentColor.color -eq '#000000'
}

# Test 7: OLED transparency override present
Test-Assert "UseOLEDTaskbarTransparency registry override present" {
    $cfg = Get-W11ThemeConfig -PresetName 'pure-black'
    $override = $cfg.advanced.registryOverrides | Where-Object { $_.name -eq 'UseOLEDTaskbarTransparency' }
    $null -ne $override -and $override.value -eq 1
}

# Test 8: Wallpaper path resolves
Test-Assert "Wallpaper path resolves to existing file" {
    $cfg = Get-W11ThemeConfig -PresetName 'pure-black'
    $wpRelative = $cfg.wallpaper.path
    # Try absolute first (ConfigManager may resolve it), then relative
    if (Test-Path $wpRelative) { $true }
    else {
        $wpFull = Join-Path 'C:\Dev\w11-theming-suite\assets\wallpapers' $wpRelative
        Test-Path $wpFull
    }
}

# Test 9: TranslucentTB detection works
Test-Assert "Test-W11TranslucentTBInstalled runs without error" {
    $status = Test-W11TranslucentTBInstalled
    $null -ne $status.Installed -and $null -ne $status.Running
}

# Test 10: Get-W11TranslucentTBConfig reads current settings
Test-Assert "Get-W11TranslucentTBConfig reads config" {
    $ttbCfg = Get-W11TranslucentTBConfig
    # Should return non-null if TTB is installed
    $status = Test-W11TranslucentTBInstalled
    if ($status.Installed) {
        $null -ne $ttbCfg
    } else {
        # If not installed, should return null gracefully
        $true
    }
}

# Test 11: .theme file generation works with pure-black
Test-Assert ".theme file generates from pure-black config" {
    $cfg = Get-W11ThemeConfig -PresetName 'pure-black'
    $themeFile = New-W11ThemeFile -Config $cfg
    (Test-Path $themeFile) -and ((Get-Content $themeFile -Raw) -match '\[Theme\]')
}

# Test 12: Get-W11InstalledThemes lists pure-black
Test-Assert "Get-W11InstalledThemes includes pure-black" {
    $themes = Get-W11InstalledThemes
    ($themes | Where-Object { $_.Name -eq 'Pure Black' }) -ne $null
}

Write-Host ""
Write-Host "=== Results: $passed/$total passed, $failed failed ===" -ForegroundColor $(if ($failed -eq 0) {'Green'} else {'Red'})
