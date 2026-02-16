@{
    # Module manifest for WallpaperManager
    # Part of the w11-theming-suite project

    RootModule        = 'WallpaperManager.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a3b2c1d0-4e5f-6a7b-8c9d-0e1f2a3b4c5d'
    Author            = 'w11-theming-suite'
    Description       = 'Manages Windows 11 desktop wallpaper settings via registry and P/Invoke.'

    FunctionsToExport = @(
        'Set-W11Wallpaper',
        'Get-W11Wallpaper'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
