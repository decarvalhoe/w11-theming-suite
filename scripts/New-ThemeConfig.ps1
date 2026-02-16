#Requires -Version 5.1
<#
.SYNOPSIS
    Interactive theme configuration wizard.
.DESCRIPTION
    Creates a new theme JSON config through an interactive wizard.
.EXAMPLE
    .\New-ThemeConfig.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# Import the root module
$ModulePath = Join-Path (Split-Path $PSScriptRoot) 'w11-theming-suite.psd1'
Import-Module $ModulePath -Force

Write-Host "`n=== w11-theming-suite - Theme Config Wizard ===" -ForegroundColor Cyan
Write-Host ""

New-W11ThemeConfig
