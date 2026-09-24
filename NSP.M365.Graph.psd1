@{
    RootModule = 'NSP.M365.Graph.psm1'
    ModuleVersion = '0.1.0'
    GUID = '5608504a-bd61-4a57-9c87-d6af884a2aed'
    Author = 'Network Systems Plus'
    CompanyName = 'Network Systems Plus'
    Copyright = '(c) 2026 Network Systems Plus. All rights reserved.'
    Description = 'Shared Microsoft Graph connection, paging, and swappable request transport for NSP M365 tools.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Connect-NSPGraph'
        'Get-NSPGraphTransport'
        'Invoke-NSPGraphCollection'
        'Invoke-NSPGraphRequest'
        'Set-NSPGraphTransport'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('MicrosoftGraph', 'M365', 'Entra', 'NSP')
            ReleaseNotes = 'Initial module: Connect-NSPGraph extracted from NSP-IntuneApps, request/paging transport layer with a fake transport for tests.'
        }
    }
}
