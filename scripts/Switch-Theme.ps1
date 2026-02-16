#Requires -Version 5.1
<#
.SYNOPSIS
    Quick theme switcher.
.DESCRIPTION
    Switches between installed theme presets. Saves original state on first switch.
.PARAMETER PresetName
    Name of the preset to switch to. If not provided, shows interactive selection.
.EXAMPLE
    .\Switch-Theme.ps1 -PresetName deep-dark
.EXAMPLE
    .\Switch-Theme.ps1
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$PresetName
)

$ErrorActionPreference = 'Stop'

# Import the root module
$ModulePath = Join-Path (Split-Path $PSScriptRoot) 'w11-theming-suite.psd1'
Import-Module $ModulePath -Force

# If no preset name, show interactive selection
if (-not $PresetName) {
    Write-Host "`n=== w11-theming-suite - Theme Switcher ===" -ForegroundColor Cyan
    Write-Host ""

    $themes = Get-W11InstalledThemes
    if ($themes.Count -eq 0) {
        Write-Host "No themes found." -ForegroundColor Yellow
        exit 1
    }

    Write-Host "Available themes:" -ForegroundColor White
    for ($i = 0; $i -lt $themes.Count; $i++) {
        $t = $themes[$i]
        $source = if ($t.Source -eq 'preset') { '[Preset]' } else { '[User]' }
        Write-Host "  [$($i + 1)] $source $($t.Name) - $($t.Description)" -ForegroundColor Gray
    }
    Write-Host ""

    $selection = Read-Host "Select a theme (1-$($themes.Count))"
    $idx = [int]$selection - 1
    if ($idx -lt 0 -or $idx -ge $themes.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        exit 1
    }

    $PresetName = $themes[$idx].FileName -replace '\.json$', ''
}

Switch-W11Theme -PresetName $PresetName
