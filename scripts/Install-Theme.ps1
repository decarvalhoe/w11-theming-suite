#Requires -Version 5.1
<#
.SYNOPSIS
    One-click theme installation script.
.DESCRIPTION
    Loads the w11-theming-suite module and applies a theme preset.
    Automatically backs up the current state before applying.
.PARAMETER PresetName
    Name of the preset to install (e.g. 'deep-dark', 'macos-monterey').
.PARAMETER ConfigPath
    Path to a custom theme JSON config file.
.PARAMETER NoBackup
    Skip creating a backup before applying.
.EXAMPLE
    .\Install-Theme.ps1 -PresetName deep-dark
.EXAMPLE
    .\Install-Theme.ps1 -ConfigPath C:\MyThemes\custom.json
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$PresetName,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'

# Import the root module
$ModulePath = Join-Path (Split-Path $PSScriptRoot) 'w11-theming-suite.psd1'
Import-Module $ModulePath -Force

# If no parameters provided, show interactive selection
if (-not $PresetName -and -not $ConfigPath) {
    Write-Host "`n=== w11-theming-suite - Theme Installer ===" -ForegroundColor Cyan
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

# Install the theme
$params = @{}
if ($PresetName) { $params['PresetName'] = $PresetName }
if ($ConfigPath) { $params['ConfigPath'] = $ConfigPath }
if ($NoBackup)   { $params['NoBackup'] = $true }

Install-W11Theme @params
