#Requires -Version 5.1
<#
.SYNOPSIS
    One-click theme restoration script.
.DESCRIPTION
    Restores Windows theme state from a backup created by Install-Theme.
.PARAMETER BackupName
    Name of the backup to restore. If not provided, shows interactive selection.
.EXAMPLE
    .\Uninstall-Theme.ps1
.EXAMPLE
    .\Uninstall-Theme.ps1 -BackupName "pre-Deep_Dark"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BackupName
)

$ErrorActionPreference = 'Stop'

# Import the root module
$ModulePath = Join-Path (Split-Path $PSScriptRoot) 'w11-theming-suite.psd1'
Import-Module $ModulePath -Force

# If no backup name, show interactive selection
if (-not $BackupName) {
    Write-Host "`n=== w11-theming-suite - Theme Uninstaller ===" -ForegroundColor Cyan
    Write-Host ""

    $backups = Get-W11ThemeBackups
    if ($backups.Count -eq 0) {
        Write-Host "No backups found. Nothing to restore." -ForegroundColor Yellow
        exit 0
    }

    Write-Host "Available backups:" -ForegroundColor White
    for ($i = 0; $i -lt $backups.Count; $i++) {
        $b = $backups[$i]
        Write-Host "  [$($i + 1)] $($b.Name) - Created: $($b.Timestamp)" -ForegroundColor Gray
    }
    Write-Host ""

    $selection = Read-Host "Select a backup to restore (1-$($backups.Count))"
    $idx = [int]$selection - 1
    if ($idx -lt 0 -or $idx -ge $backups.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        exit 1
    }

    $BackupName = $backups[$idx].Name
}

Uninstall-W11Theme -BackupName $BackupName
