@{
    # Module manifest for BackupRestore
    # Part of w11-theming-suite - snapshot and restore Windows theme state

    RootModule        = 'BackupRestore.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'f4a7c3e1-9b2d-4f8a-a1c6-3e5d7b9f0a2c'
    Author            = 'w11-theming-suite'
    Description       = 'Backup and restore for w11-theming-suite - snapshot and restore Windows theme state'

    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        'Backup-W11ThemeState',
        'Restore-W11ThemeState',
        'Get-W11ThemeBackups',
        'Remove-W11ThemeBackup'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()

    PrivateData = @{
        PSData = @{
            Tags       = @('Windows11', 'Theming', 'Backup', 'Restore')
            ProjectUri = ''
        }
    }
}
