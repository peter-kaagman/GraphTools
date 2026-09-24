@{
    # Script module associated with this manifest
    RootModule = 'GraphTools.psm1'

    # Version number of this module
    ModuleVersion = '0.1.0'

    # Unique identifier for this module
    GUID = '18d3cc1b-7c44-4d8f-b381-382090b46299'

    # Module author
    Author = 'Peter Kaagman'

    # Module description
    Description = 'Lightweight PowerShell helpers for Microsoft Graph REST API operations.'

    # Minimum PowerShell version supported by this module
    PowerShellVersion = '5.1'

    # Functions exported by this module
    FunctionsToExport = @(
        'Get-GraphHeaders'
        'Get-GraphCollection'
        'Get-Groups'
        'Get-Teams'
        'Get-GroupByMailNickname'
        'Get-GroupById'
        'Test-TeamArchived'
        'Set-TeamArchived'
        'Clear-TeamArchived'
        'Remove-Group'
    )

    # No cmdlets, variables or aliases are exported
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData = @{
        PSData = @{
            Tags = @(
                'MicrosoftGraph'
                'Microsoft365'
                'Entra'
                'Teams'
                'PowerShell'
            )
        }
    }
}

