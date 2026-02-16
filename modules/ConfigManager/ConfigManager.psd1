@{
    # Module manifest for ConfigManager

    # Script module associated with this manifest
    RootModule        = 'ConfigManager.psm1'

    # Version number of this module
    ModuleVersion     = '1.0.0'

    # ID used to uniquely identify this module
    GUID              = 'a3b8f7e2-1d4c-4a9e-b6f0-8c2e5d7a1b3f'

    # Author of this module
    Author            = 'w11-theming-suite'

    # Description of the functionality provided by this module
    Description       = 'Configuration manager for w11-theming-suite - JSON parsing, validation, and inheritance'

    # Minimum version of PowerShell required by this module
    PowerShellVersion = '5.1'

    # Functions to export from this module
    FunctionsToExport = @(
        'Get-W11ThemeConfig',
        'New-W11ThemeConfig',
        'Merge-W11ThemeConfig'
    )

    # Cmdlets to export from this module
    CmdletsToExport   = @()

    # Variables to export from this module
    VariablesToExport  = @()

    # Aliases to export from this module
    AliasesToExport    = @()
}
