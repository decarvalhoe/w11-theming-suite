#Requires -Version 5.1
<#
.SYNOPSIS
    Export a .theme file from a config preset.
.DESCRIPTION
    Generates a Windows .theme file from a JSON config without applying it.
.PARAMETER PresetName
    Name of the preset to export.
.PARAMETER ConfigPath
    Path to a custom config JSON file.
.PARAMETER OutputPath
    Where to save the .theme file. Defaults to output\ folder.
.EXAMPLE
    .\Export-ThemeFile.ps1 -PresetName deep-dark
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$PresetName,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

# Import the root module
$ModulePath = Join-Path (Split-Path $PSScriptRoot) 'w11-theming-suite.psd1'
Import-Module $ModulePath -Force

# Load config
if ($PresetName) {
    $config = Get-W11ThemeConfig -PresetName $PresetName
}
elseif ($ConfigPath) {
    $config = Get-W11ThemeConfig -Path $ConfigPath
}
else {
    Write-Host "Please specify -PresetName or -ConfigPath." -ForegroundColor Yellow
    exit 1
}

# Generate theme file
$params = @{ Config = $config }
if ($OutputPath) { $params['OutputPath'] = $OutputPath }

$result = New-W11ThemeFile @params
Write-Host "Theme file exported to: $result" -ForegroundColor Green
