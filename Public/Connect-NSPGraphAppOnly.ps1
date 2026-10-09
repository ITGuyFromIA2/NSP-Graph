function Connect-NSPGraphAppOnly {
    <#
    .SYNOPSIS
        Unattended sign-in: connects as an app with a certificate from the local store, and returns the context.
    .DESCRIPTION
        For scheduled jobs only. The certificate is referenced by thumbprint and must be in the
        store of the account the job runs as. Refuses a context that is not app-only for that
        client and tenant.
    .EXAMPLE
        Connect-NSPGraphAppOnly -ClientId $record.ClientId -TenantId $record.TenantId -CertificateThumbprint $record.CertificateThumbprint
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$CertificateThumbprint
    )

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop | Out-Null
    $context = Get-MgContext
    if (-not $context -or [string]$context.AuthType -ne 'AppOnly' -or $context.ClientId -ne $ClientId -or $context.TenantId -ne $TenantId) {
        throw "App-only sign-in did not produce the expected context (client $ClientId, tenant $TenantId)."
    }
    $context
}
