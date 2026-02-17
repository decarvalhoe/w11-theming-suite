@{
    RootModule        = 'w11-theming-suite.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'
    Author            = 'w11-theming-suite'
    CompanyName       = 'Community'
    Copyright         = '(c) 2026. All rights reserved.'
    Description       = 'Windows 11 Native Theming Toolkit - Complete theme management via PowerShell'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Get-W11ThemeConfig',
        'New-W11ThemeConfig',
        'Merge-W11ThemeConfig',
        'Backup-W11ThemeState',
        'Restore-W11ThemeState',
        'Get-W11ThemeBackups',
        'Remove-W11ThemeBackup',
        'Set-W11RegistryTheme',
        'Get-W11RegistryTheme',
        'New-W11ThemeFile',
        'Install-W11CursorScheme',
        'Uninstall-W11CursorScheme',
        'Get-W11CursorSchemes',
        'Install-W11SoundScheme',
        'Uninstall-W11SoundScheme',
        'Get-W11SoundSchemes',
        'Set-W11Wallpaper',
        'Get-W11Wallpaper',
        'Test-W11TranslucentTBInstalled',
        'Get-W11TranslucentTBConfig',
        'Set-W11TranslucentTBConfig',
        'Set-W11WindowBackdrop',
        'Set-W11WindowColors',
        'Set-W11NativeTaskbarTransparency',
        'Get-W11NativeTaskbarTransparency',
        'Remove-W11NativeTaskbarTransparency',
        'Register-W11TaskbarTransparencyStartup',
        'Unregister-W11TaskbarTransparencyStartup',
        'Invoke-TaskbarTAPInject',
        'Set-TaskbarTAPMode',
        'Get-TaskbarExplorerPid',
        'Invoke-ShellTAPInject',
        'Set-ShellTAPMode',
        'Invoke-StartMenuDiscovery',
        'Invoke-StartMenuTransparency',
        'Invoke-ActionCenterDiscovery',
        'Invoke-ActionCenterTransparency',
        'Start-W11BackdropWatcher',
        'Stop-W11BackdropWatcher',
        'Register-W11BackdropWatcherStartup',
        'Unregister-W11BackdropWatcherStartup',
        'Install-W11Theme',
        'Uninstall-W11Theme',
        'Switch-W11Theme',
        'Get-W11InstalledThemes',
        'Get-W11ProjectRoot'
    )
    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
    PrivateData        = @{
        PSData = @{
            Tags       = @('Windows11', 'Theme', 'Customization', 'DarkMode', 'Registry')
            LicenseUri = ''
            ProjectUri = 'https://github.com/decarvalhoe/w11-theming-suite'
        }
    }
}
