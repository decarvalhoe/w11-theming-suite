@{
    # Module manifest for SoundSchemeBuilder
    # Part of w11-theming-suite

    RootModule        = 'SoundSchemeBuilder.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a3b2c1d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d'
    Author            = 'w11-theming-suite'
    Description       = 'Manages custom Windows 11 sound schemes via native registry operations.'

    FunctionsToExport = @(
        'Install-W11SoundScheme'
        'Uninstall-W11SoundScheme'
        'Get-W11SoundSchemes'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()

    PrivateData = @{
        PSData = @{
            Tags       = @('Windows11', 'Sound', 'Theme', 'Registry')
            ProjectUri = ''
        }
    }
}
