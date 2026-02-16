# Test-TaskbarTAP.ps1 --Test TAP DLL injection for taskbar transparency
$ErrorActionPreference = 'Continue'

Import-Module 'C:\Dev\w11-theming-suite\w11-theming-suite.psd1' -Force -DisableNameChecking

Write-Host '=============================================' -ForegroundColor Cyan
Write-Host ' TASKBAR TAP INJECTION TEST'                  -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan

# Check DLL exists
$tapDll = 'C:\Dev\w11-theming-suite\native\bin\TaskbarTAP.dll'
if (Test-Path $tapDll) {
    $info = Get-Item $tapDll
    Write-Host "[OK] TaskbarTAP.dll: $($info.Length) bytes" -ForegroundColor Green
} else {
    Write-Host "[FAIL] TaskbarTAP.dll not found!" -ForegroundColor Red
    exit 1
}

# Find explorer PID
$explorerPid = Get-TaskbarExplorerPid
Write-Host "Explorer PID (taskbar owner): $explorerPid" -ForegroundColor Gray

# Inject
Write-Host ''
Write-Host '>>> Injecting TAP DLL (Transparent mode)...' -ForegroundColor Magenta
$result = Invoke-TaskbarTAPInject -Mode Transparent -Verbose

if ($result) {
    Write-Host ''
    Write-Host '=============================================' -ForegroundColor Green
    Write-Host ' INJECTION SUCCEEDED --Check your taskbar!'   -ForegroundColor Green
    Write-Host '=============================================' -ForegroundColor Green
    Write-Host ''
    Write-Host 'The taskbar should now be transparent.' -ForegroundColor Yellow
    Write-Host 'Waiting 20 seconds before testing mode switch...' -ForegroundColor Gray
    Start-Sleep -Seconds 20

    Write-Host ''
    Write-Host '>>> Switching to Acrylic mode...' -ForegroundColor Magenta
    Set-TaskbarTAPMode -Mode Acrylic
    Write-Host 'Waiting 10 seconds...' -ForegroundColor Gray
    Start-Sleep -Seconds 10

    Write-Host ''
    Write-Host '>>> Switching back to Default...' -ForegroundColor Magenta
    Set-TaskbarTAPMode -Mode Default
    Write-Host 'Waiting 5 seconds...' -ForegroundColor Gray
    Start-Sleep -Seconds 5

    Write-Host ''
    Write-Host 'Test complete.' -ForegroundColor Cyan
} else {
    Write-Host ''
    Write-Host '=============================================' -ForegroundColor Red
    Write-Host ' INJECTION FAILED'                             -ForegroundColor Red
    Write-Host '=============================================' -ForegroundColor Red
}
