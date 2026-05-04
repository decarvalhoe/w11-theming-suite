@{
    RootModule        = 'TaskbarGrouping.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'd9f8a7b6-c5e4-4d3a-b2c1-9f8e7d6c5b4a'
    Author            = 'w11-theming-suite'
    CompanyName       = 'w11-theming-suite'
    Copyright         = '(c) w11-theming-suite. MIT License.'
    Description       = 'Force taskbar grouping of multiple instances of an application by profile, on Windows 11 (incl. build 26200+ where AUMID/window class alone are insufficient). Combines three native techniques in cascade: per-profile EXE hardlinks, Win32 window class override, and PKEY_AppUserModel_ID. Designed for terminal multiplexers (WezTerm, Windows Terminal) but works for any executable. Companion to NativeTaskbarTransparency.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        # High-level configuration
        'Set-W11TaskbarGrouping',
        'Get-W11TaskbarGrouping',
        'Remove-W11TaskbarGrouping',

        # EXE alias management (per-profile hardlinks)
        'New-W11TaskbarExeAlias',
        'Get-W11TaskbarExeAlias',
        'Remove-W11TaskbarExeAlias',

        # AUMID management (window-level)
        'Set-W11WindowAumid',
        'Get-W11WindowAumid',

        # Helper for downstream launcher integrations
        'Get-W11TaskbarGroupingLaunchSpec'
    )
    CmdletsToExport    = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
    PrivateData        = @{
        PSData = @{
            Tags         = @('Windows11', 'Taskbar', 'Grouping', 'AUMID', 'WezTerm', 'Theming', 'Native', 'No-Third-Party')
            ProjectUri   = 'https://github.com/decarvalhoe/w11-theming-suite'
            LicenseUri   = 'https://github.com/decarvalhoe/w11-theming-suite/blob/main/LICENSE'
            ReleaseNotes = @'
v1.0.0 — Initial release.
- Set-W11TaskbarGrouping: high-level config (profile -> EXE alias + class + AUMID).
- New-W11TaskbarExeAlias: create per-profile hardlink so Windows groups by image name.
- Set-W11WindowAumid: P/Invoke SHGetPropertyStoreForWindow + PKEY_AppUserModel_ID.
- Get-W11TaskbarGroupingLaunchSpec: returns the exact arg list a launcher should
  pass to spawn an instance with grouping applied (used by universal-project-launcher).
'@
        }
    }
}
