# Test-DwmDiagnostic.ps1 — Full DWM diagnostic for w11-theming-suite
# Checks system prerequisites, GPU/driver info, and tests each DWM attribute
$ErrorActionPreference = 'Continue'

# PS 5.1 safe uint32 constants (hex literals > 0x7FFFFFFF fail in PS 5.1)
$COLOR_DEFAULT = [uint32]4294967295  # 0xFFFFFFFF
$COLOR_NONE    = [uint32]4294967294  # 0xFFFFFFFE

Import-Module 'C:\Dev\w11-theming-suite\w11-theming-suite.psd1' -Force -DisableNameChecking

Write-Host '=============================================' -ForegroundColor Cyan
Write-Host ' DWM DIAGNOSTIC — w11-theming-suite'          -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan

# ── SYSTEM INFO ──
$ver = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
Write-Host ''
Write-Host '=== SYSTEM ===' -ForegroundColor Yellow
Write-Host "  Windows 11 $($ver.DisplayVersion) Build $($ver.CurrentBuildNumber).$($ver.UBR)"
Write-Host "  PowerShell $($PSVersionTable.PSVersion)"

# GPU
try {
    Get-CimInstance Win32_VideoController | ForEach-Object {
        Write-Host "  GPU: $($_.Name) ($($_.DriverVersion))" -ForegroundColor Gray
        Write-Host "    Resolution: $($_.VideoModeDescription)" -ForegroundColor DarkGray
    }
} catch {}

# ── PREREQUISITES ──
Write-Host ''
Write-Host '=== PREREQUISITES ===' -ForegroundColor Yellow

$persKey = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
$dwmKey  = 'HKCU:\SOFTWARE\Microsoft\Windows\DWM'

$et = (Get-ItemProperty $persKey -EA SilentlyContinue).EnableTransparency
if ($et -eq 1) { Write-Host '  [OK] Transparency Effects: ON' -ForegroundColor Green }
elseif ($et -eq 0) { Write-Host '  [BLOCKER] Transparency Effects: OFF' -ForegroundColor Red }
else { Write-Host "  [?] Transparency Effects: $et" -ForegroundColor Yellow }

$dwm = Get-ItemProperty $dwmKey -EA SilentlyContinue
Write-Host "  ColorPrevalence=$($dwm.ColorPrevalence) EnableWindowColorization=$($dwm.EnableWindowColorization)" -ForegroundColor Gray

$vfx = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' -EA SilentlyContinue).VisualFXSetting
Write-Host "  VisualFX=$vfx (0=Auto 1=BestAppearance 2=BestPerformance 3=Custom)" -ForegroundColor Gray
if ($vfx -eq 2) { Write-Host '  [WARNING] Best Performance may disable effects' -ForegroundColor Yellow }

try {
    $pp = powercfg /getactivescheme 2>&1
    Write-Host "  Power: $pp" -ForegroundColor DarkGray
} catch {}

# ── DWM API TESTS ──
Write-Host ''
Write-Host '=== DWM API TESTS ===' -ForegroundColor Yellow

$wins = [W11ThemeSuite.DwmHelper]::GetVisibleWindows()
Write-Host "  $($wins.Count) visible windows:" -ForegroundColor Gray
foreach ($w in $wins) {
    $cn = [W11ThemeSuite.DwmHelper]::GetWindowClassName($w)
    Write-Host "    $cn (0x$($w.ToString('X')))" -ForegroundColor DarkGray
}

# Apply colors to ALL windows and report per-window results
Write-Host ''
Write-Host '>>> Applying RED BORDER + CYAN CAPTION + YELLOW TEXT to all windows...' -ForegroundColor Magenta

$red    = [uint32]255        # 0x000000FF BGR = red
$cyan   = [uint32]16776960   # 0x00FFFF00 BGR = cyan
$yellow = [uint32]65535      # 0x0000FFFF BGR = yellow

foreach ($w in $wins) {
    $cn = [W11ThemeSuite.DwmHelper]::GetWindowClassName($w)
    $hr1 = [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 34, [ref]$red, 4)
    $hr2 = [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 35, [ref]$cyan, 4)
    $hr3 = [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 36, [ref]$yellow, 4)
    $status = if ($hr1 -eq 0 -and $hr2 -eq 0 -and $hr3 -eq 0) { 'OK' } else { "FAIL(0x$($hr1.ToString('X8')))" }
    Write-Host "    $cn => $status" -ForegroundColor $(if($status -eq 'OK'){'Green'}else{'Red'})
}

Write-Host ''
Write-Host '  >>> REGARDE: bordures rouges + titre cyan + texte jaune? (15s) <<<' -ForegroundColor Yellow
Start-Sleep -Seconds 15

# Reset colors
foreach ($w in $wins) {
    [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 34, [ref]$COLOR_DEFAULT, 4) | Out-Null
    [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 35, [ref]$COLOR_DEFAULT, 4) | Out-Null
    [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 36, [ref]$COLOR_DEFAULT, 4) | Out-Null
}
Write-Host '  Colors reset.' -ForegroundColor Gray

# Backdrop test on all windows
Write-Host ''
Write-Host '>>> Applying MICA backdrop + Dark Mode to all windows...' -ForegroundColor Magenta

foreach ($w in $wins) {
    $cn = [W11ThemeSuite.DwmHelper]::GetWindowClassName($w)
    $hr = [W11ThemeSuite.DwmHelper]::SetBackdropType($w, 2)  # Mica
    $one = 1
    [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 20, [ref]$one, 4) | Out-Null
    $status = if ($hr -eq 0) { 'OK' } else { "FAIL(0x$($hr.ToString('X8')))" }
    Write-Host "    $cn => $status" -ForegroundColor $(if($status -eq 'OK'){'Green'}else{'Red'})
}

Write-Host ''
Write-Host '  >>> REGARDE: Mica backdrop (teinte subtile du bureau)? (15s) <<<' -ForegroundColor Yellow
Start-Sleep -Seconds 15

# Reset all
Write-Host ''
Write-Host '>>> Reset...' -ForegroundColor Magenta
$auto = 0
$zero = 0
foreach ($w in $wins) {
    [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 38, [ref]$auto, 4) | Out-Null
    [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 34, [ref]$COLOR_DEFAULT, 4) | Out-Null
    [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 35, [ref]$COLOR_DEFAULT, 4) | Out-Null
    [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 36, [ref]$COLOR_DEFAULT, 4) | Out-Null
    [W11ThemeSuite.DwmHelper]::DwmSetWindowAttribute($w, 20, [ref]$zero, 4) | Out-Null
    [W11ThemeSuite.DwmHelper]::ResetFrame($w) | Out-Null
}
Write-Host '  All reset.' -ForegroundColor Green

Write-Host ''
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host ' DONE — Report which windows showed effects.' -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan
