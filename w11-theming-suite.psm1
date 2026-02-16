# w11-theming-suite - Windows 11 Native Theming Toolkit
# Root module loader - imports all submodules

$script:ProjectRoot = $PSScriptRoot
$script:ModulePath = Join-Path $PSScriptRoot 'modules'

# Import submodules in dependency order
$Modules = @(
    'ConfigManager',
    'BackupRestore',
    'RegistryConfigurator',
    'ThemeFileBuilder',
    'CursorSchemeBuilder',
    'SoundSchemeBuilder',
    'WallpaperManager',
    'ThemeOrchestrator'
)

foreach ($mod in $Modules) {
    $modFile = Join-Path $ModulePath "$mod\$mod.psm1"
    if (Test-Path $modFile) {
        try {
            Import-Module $modFile -Force -DisableNameChecking
        }
        catch {
            Write-Warning "Failed to load module '$mod': $_"
        }
    }
    else {
        Write-Warning "Module file not found: $modFile"
    }
}

# Export project root path for submodules
function Get-W11ProjectRoot {
    return $script:ProjectRoot
}

Export-ModuleMember -Function @(
    # ConfigManager
    'Get-W11ThemeConfig',
    'New-W11ThemeConfig',
    'Merge-W11ThemeConfig',
    # BackupRestore
    'Backup-W11ThemeState',
    'Restore-W11ThemeState',
    'Get-W11ThemeBackups',
    'Remove-W11ThemeBackup',
    # RegistryConfigurator
    'Set-W11RegistryTheme',
    'Get-W11RegistryTheme',
    # ThemeFileBuilder
    'New-W11ThemeFile',
    # CursorSchemeBuilder
    'Install-W11CursorScheme',
    'Uninstall-W11CursorScheme',
    'Get-W11CursorSchemes',
    # SoundSchemeBuilder
    'Install-W11SoundScheme',
    'Uninstall-W11SoundScheme',
    'Get-W11SoundSchemes',
    # WallpaperManager
    'Set-W11Wallpaper',
    'Get-W11Wallpaper',
    # ThemeOrchestrator
    'Install-W11Theme',
    'Uninstall-W11Theme',
    'Switch-W11Theme',
    'Get-W11InstalledThemes',
    # Utility
    'Get-W11ProjectRoot'
)
