function Get-NSPGraphAppCertificate {
    <#
    .SYNOPSIS
        Unexpired certificates with a private key for an app-only registration, newest expiry first.
    .DESCRIPTION
        Certificates are found by subject (CN=<AppName>) in one store. Wrapped so tests never
        touch a real certificate store.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$Store
    )
    if (-not (Test-Path -LiteralPath $Store)) { return }
    @(Get-ChildItem -LiteralPath $Store | Where-Object {
            $_.Subject -eq $Subject -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date)
        } | Sort-Object NotAfter -Descending)
}
