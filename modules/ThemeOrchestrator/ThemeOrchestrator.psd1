@{
    # Module manifest for ThemeOrchestrator
    # High-level orchestrator that coordinates all theming modules

    RootModule        = 'ThemeOrchestrator.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a3b7c9d1-4e5f-6a7b-8c9d-0e1f2a3b4c5d'
    Author            = 'w11-theming-suite'
    Description       = 'High-level orchestrator for Windows 11 theme installation, removal, and switching.'

    # Functions to export from this module
    FunctionsToExport = @(
        'Install-W11Theme',
        'Uninstall-W11Theme',
        'Switch-W11Theme',
        'Get-W11InstalledThemes'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
