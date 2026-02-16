@{
    RootModule        = 'NativeTaskbarTransparency.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'c3d4e5f6-a7b8-9012-cdef-345678901234'
    Author            = 'w11-theming-suite'
    CompanyName       = 'w11-theming-suite'
    Copyright         = '(c) w11-theming-suite. All rights reserved.'
    Description       = 'Native taskbar transparency for Windows 11 via SetWindowCompositionAttribute - no third-party software required'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Set-W11NativeTaskbarTransparency',
        'Get-W11NativeTaskbarTransparency',
        'Remove-W11NativeTaskbarTransparency',
        'Register-W11TaskbarTransparencyStartup',
        'Unregister-W11TaskbarTransparencyStartup'
    )
    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
    PrivateData        = @{
        PSData = @{
            Tags       = @('Windows11', 'Taskbar', 'Transparency', 'Theming', 'Native')
            ProjectUri = ''
        }
    }
}
