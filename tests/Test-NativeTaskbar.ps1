# Test-NativeTaskbar.ps1 - Tests for NativeTaskbarTransparency module
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

Write-Host "=== NativeTaskbarTransparency Module Tests ===" -ForegroundColor Cyan
Write-Host ""

# Import module
Import-Module 'C:\Dev\w11-theming-suite\w11-theming-suite.psd1' -Force -DisableNameChecking

# Test 1-5: All native functions exported
Test-Assert "Set-W11NativeTaskbarTransparency exported" {
    $null -ne (Get-Command Set-W11NativeTaskbarTransparency -EA SilentlyContinue)
}
Test-Assert "Get-W11NativeTaskbarTransparency exported" {
    $null -ne (Get-Command Get-W11NativeTaskbarTransparency -EA SilentlyContinue)
}
Test-Assert "Remove-W11NativeTaskbarTransparency exported" {
    $null -ne (Get-Command Remove-W11NativeTaskbarTransparency -EA SilentlyContinue)
}
Test-Assert "Register-W11TaskbarTransparencyStartup exported" {
    $null -ne (Get-Command Register-W11TaskbarTransparencyStartup -EA SilentlyContinue)
}
Test-Assert "Unregister-W11TaskbarTransparencyStartup exported" {
    $null -ne (Get-Command Unregister-W11TaskbarTransparencyStartup -EA SilentlyContinue)
}

# Test 6: Total module exports count (26 original + 5 native = 31)
Test-Assert "Module exports 31 functions total" {
    $cmds = Get-Command -Module 'w11-theming-suite'
    $cmds.Count -eq 31
}

# Test 7: P/Invoke types loaded
Test-Assert "P/Invoke AccentPolicy type available" {
    $null -ne ([Type]'W11ThemeSuite.AccentPolicy')
}
Test-Assert "P/Invoke TaskbarTransparency type available" {
    $null -ne ([Type]'W11ThemeSuite.TaskbarTransparency')
}

# Test 8: FindWindow can locate the taskbar
Test-Assert "FindWindow locates Shell_TrayWnd" {
    $hwnd = [W11ThemeSuite.TaskbarTransparency]::FindWindow('Shell_TrayWnd', $null)
    $hwnd -ne [IntPtr]::Zero
}

# Test 9: Apply clear transparency to taskbar (live test!)
Test-Assert "Apply clear transparency succeeds" {
    $hwnd = [W11ThemeSuite.TaskbarTransparency]::FindWindow('Shell_TrayWnd', $null)
    [W11ThemeSuite.TaskbarTransparency]::Apply($hwnd, 2, 0x00000000)
}

# Test 10: Set via high-level function
Test-Assert "Set-W11NativeTaskbarTransparency -Style clear works" {
    Set-W11NativeTaskbarTransparency -Style clear
    $true  # If no exception, it passed
}

# Test 11: Get returns config after Set
Test-Assert "Get-W11NativeTaskbarTransparency returns saved config" {
    $cfg = Get-W11NativeTaskbarTransparency
    $cfg.Style -eq 'clear' -and $cfg.Enabled -eq $true
}

# Test 12: pure-black config has nativeTaskbar section
Test-Assert "pure-black preset has nativeTaskbar config" {
    $cfg = Get-W11ThemeConfig -PresetName 'pure-black'
    $null -ne $cfg.advanced.nativeTaskbar -and $cfg.advanced.nativeTaskbar.style -eq 'clear'
}

# Test 13: Reset to normal
Test-Assert "Reset to normal (ACCENT_DISABLED) works" {
    Set-W11NativeTaskbarTransparency -Style normal
    $true
}

# Test 14: Clean up - remove config
Test-Assert "Remove-W11NativeTaskbarTransparency cleans up" {
    Remove-W11NativeTaskbarTransparency
    $cfg = Get-W11NativeTaskbarTransparency
    $cfg.Enabled -eq $false
}

Write-Host ""
Write-Host "=== Results: $passed/$total passed, $failed failed ===" -ForegroundColor $(if ($failed -eq 0) {'Green'} else {'Red'})
