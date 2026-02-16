@{
    # Module manifest for CursorSchemeBuilder
    # Part of the w11-theming-suite project

    RootModule        = 'CursorSchemeBuilder.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a3c7e2f1-8b4d-4e6a-9f01-3d5c7a9b2e4f'
    Author            = 'w11-theming-suite'
    Description       = 'Manages custom Windows 11 cursor schemes natively via the registry.'

    # Functions to export from this module
    FunctionsToExport = @(
        'Install-W11CursorScheme',
        'Uninstall-W11CursorScheme',
        'Get-W11CursorSchemes'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()

    PrivateData = @{
        PSData = @{
            Tags       = @('Windows11', 'Cursors', 'Theming', 'Customization')
            ProjectUri = ''
        }
    }
}
