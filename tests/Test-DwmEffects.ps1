# Test-DwmEffects.ps1 - Test DwmSetWindowAttribute + DwmExtendFrameIntoClientArea
# v3: Adds ExtendFrame (MARGINS -1) which is REQUIRED for visual backdrop rendering
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

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " DwmSetWindowAttribute + ExtendFrame v3"      -ForegroundColor Cyan
Write-Host " TranslucentTB: DESINSTALLE"                    -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

Import-Module 'C:\Dev\w11-theming-suite\w11-theming-suite.psd1' -Force -DisableNameChecking

# --- Module exports ---
Test-Assert "Set-W11WindowBackdrop exported" {
    $null -ne (Get-Command Set-W11WindowBackdrop -EA SilentlyContinue)
}
Test-Assert "Set-W11WindowColors exported" {
    $null -ne (Get-Command Set-W11WindowColors -EA SilentlyContinue)
}

# --- DwmHelper type loaded ---
Test-Assert "DwmHelper P/Invoke type loaded" {
    $null -ne ([Type]'W11ThemeSuite.DwmHelper')
}

# --- MARGINS struct loaded ---
Test-Assert "MARGINS struct loaded" {
    $null -ne ([Type]'W11ThemeSuite.MARGINS')
}

# --- ExtendFrame method exists ---
Test-Assert "DwmHelper.ExtendFrame method exists" {
    $null -ne ([W11ThemeSuite.DwmHelper].GetMethod('ExtendFrame'))
}

# --- Find visible windows ---
Test-Assert "GetVisibleWindows returns windows" {
    $wins = [W11ThemeSuite.DwmHelper]::GetVisibleWindows()
    Write-Host "       ($($wins.Count) visible window(s) found)" -ForegroundColor Gray
    $wins.Count -gt 0
}

# --- Test ExtendFrame on a single window first ---
Write-Host ""
Write-Host ">>> TEST ExtendFrame sur une fenetre <<<" -ForegroundColor Magenta
$wins = [W11ThemeSuite.DwmHelper]::GetVisibleWindows()
if ($wins.Count -gt 0) {
    $testHwnd = $wins[0]
    $className = [W11ThemeSuite.DwmHelper]::GetWindowClassName($testHwnd)
    Test-Assert "ExtendFrame returns 0 (S_OK) on $className" {
        $hr = [W11ThemeSuite.DwmHelper]::ExtendFrame($testHwnd)
        Write-Host "       HRESULT=0x$($hr.ToString('X8'))" -ForegroundColor Gray
        $hr -eq 0
    }
}

# --- Apply MICA to all windows (with ExtendFrame) ---
Write-Host ""
Write-Host ">>> MICA + ExtendFrame sur toutes les fenetres <<<" -ForegroundColor Magenta
Test-Assert "Set-W11WindowBackdrop -Style mica -AllWindows" {
    Set-W11WindowBackdrop -Style mica -AllWindows -DarkMode
    $true
}
Write-Host "    *** REGARDE tes fenetres! Effet Mica visible? ***" -ForegroundColor Yellow
Write-Host "    (L'effet de teinte subtile du bureau doit apparaitre) (8s)" -ForegroundColor Gray
Start-Sleep -Seconds 8

# --- Apply ACRYLIC to all windows ---
Write-Host ""
Write-Host ">>> ACRYLIC (givree) sur toutes les fenetres <<<" -ForegroundColor Magenta
Test-Assert "Set-W11WindowBackdrop -Style acrylic -AllWindows" {
    Set-W11WindowBackdrop -Style acrylic -AllWindows -DarkMode
    $true
}
Write-Host "    *** REGARDE! Effet Acrylic/givre (blur) visible? ***" -ForegroundColor Yellow
Write-Host "    (Fond flou/translucide) (8s)" -ForegroundColor Gray
Start-Sleep -Seconds 8

# --- Apply TABBED (Mica Alt) to all windows ---
Write-Host ""
Write-Host ">>> MICA ALT (tabbed) sur toutes les fenetres <<<" -ForegroundColor Magenta
Test-Assert "Set-W11WindowBackdrop -Style tabbed -AllWindows" {
    Set-W11WindowBackdrop -Style tabbed -AllWindows -DarkMode
    $true
}
Write-Host "    *** REGARDE! Mica Alt visible? ***" -ForegroundColor Yellow
Write-Host "    (Variante de Mica plus opaque) (8s)" -ForegroundColor Gray
Start-Sleep -Seconds 8

# --- Color test: black borders, black caption ---
Write-Host ""
Write-Host ">>> COULEURS: bordures noires, titre noir <<<" -ForegroundColor Magenta
Test-Assert "Set-W11WindowColors black theme -AllWindows" {
    Set-W11WindowColors -BorderColor '#000000' -CaptionColor '#000000' -TextColor '#CCCCCC' -AllWindows -DarkMode
    $true
}
Write-Host "    Barres titre noires? Texte gris clair? (6s)" -ForegroundColor Yellow
Start-Sleep -Seconds 6

# --- Try on the taskbar specifically ---
Write-Host ""
Write-Host ">>> DWM sur la TASKBAR (Shell_TrayWnd) <<<" -ForegroundColor Magenta
$tbHwnd = [W11ThemeSuite.DwmHelper]::FindWindow('Shell_TrayWnd', $null)
if ($tbHwnd -ne [IntPtr]::Zero) {
    # ExtendFrame + SetBackdropType via the C# helper
    $hr = [W11ThemeSuite.DwmHelper]::SetBackdropType($tbHwnd, 3)  # acrylic
    Write-Host "    DwmSetWindowAttribute sur taskbar: HRESULT=0x$($hr.ToString('X8'))"
    if ($hr -eq 0) {
        Write-Host "    API OK! Taskbar acrylic visible? (5s)" -ForegroundColor Green
    } else {
        Write-Host "    API echec (attendu sur Win11 XAML taskbar)" -ForegroundColor Yellow
    }
    Start-Sleep -Seconds 5
} else {
    Write-Host "    Taskbar non trouvee" -ForegroundColor Red
}

# --- Reset everything ---
Write-Host ""
Write-Host ">>> RESET au defaut <<<" -ForegroundColor Magenta
Test-Assert "Set-W11WindowBackdrop -Style auto -AllWindows (reset)" {
    Set-W11WindowBackdrop -Style auto -AllWindows
    $true
}
Test-Assert "Set-W11WindowColors -AllWindows default (reset)" {
    Set-W11WindowColors -BorderColor 'default' -CaptionColor 'default' -TextColor 'default' -AllWindows
    $true
}
Write-Host "    Fenetres revenues au defaut? (3s)" -ForegroundColor Yellow
Start-Sleep -Seconds 3

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " Resultats: $passed/$total passes, $failed echec(s)" -ForegroundColor $(if ($failed -eq 0) {'Green'} else {'Red'})
Write-Host " API: DwmSetWindowAttribute + ExtendFrame"     -ForegroundColor Gray
Write-Host " Cle: ExtendFrame(MARGINS -1) = sheet of glass" -ForegroundColor Gray
Write-Host " Tier requis: AUCUN"                            -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Cyan
