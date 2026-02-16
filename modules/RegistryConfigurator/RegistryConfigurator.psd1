@{
    # Module manifest for RegistryConfigurator
    # Part of w11-theming-suite - deep Windows 11 registry customizations

    RootModule        = 'RegistryConfigurator.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a3b2c1d0-4e5f-6a7b-8c9d-0e1f2a3b4c5d'
    Author            = 'w11-theming-suite'
    Description       = 'Registry configurator for w11-theming-suite - deep Windows 11 registry customizations'

    FunctionsToExport = @(
        'Set-W11RegistryTheme',
        'Get-W11RegistryTheme'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()

    PrivateData = @{
        PSData = @{
            Tags       = @('Windows11', 'Theming', 'Registry')
            ProjectUri = ''
        }
    }
}
