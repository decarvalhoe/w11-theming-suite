@{
    RootModule        = 'NativeTaskbarTransparency.psm1'
    ModuleVersion     = '2.0.0'
    GUID              = 'c3d4e5f6-a7b8-9012-cdef-345678901234'
    Author            = 'w11-theming-suite'
    CompanyName       = 'w11-theming-suite'
    Copyright         = '(c) w11-theming-suite. All rights reserved.'
    Description       = 'Windows 11 window backdrop and color theming via the official DwmSetWindowAttribute API. Supports Mica, Acrylic, and Mica Alt backdrops plus border/caption/text color customization on all application windows. Includes backward-compatible taskbar transparency with SWCA fallback.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Set-W11WindowBackdrop',
        'Set-W11WindowColors',
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
            Tags       = @('Windows11', 'Taskbar', 'Transparency', 'Theming', 'Native', 'DWM', 'Mica', 'Acrylic', 'Backdrop')
            ProjectUri = ''
            ReleaseNotes = 'v2.0.0 - Complete rewrite from SWCA to DwmSetWindowAttribute API. Added Set-W11WindowBackdrop and Set-W11WindowColors for all application windows. Old SWCA approach kept as taskbar fallback only.'
        }
    }
}
