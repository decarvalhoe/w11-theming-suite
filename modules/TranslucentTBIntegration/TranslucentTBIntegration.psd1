@{
    RootModule        = 'TranslucentTBIntegration.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b2c3d4e5-f6a7-8901-bcde-f23456789012'
    Author            = 'w11-theming-suite'
    Description       = 'TranslucentTB integration for w11-theming-suite - read, write, and sync TranslucentTB config'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Get-W11TranslucentTBConfig',
        'Set-W11TranslucentTBConfig',
        'Test-W11TranslucentTBInstalled'
    )
    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
