function New-NSPGraphAppCertificate {
    <#
    .SYNOPSIS
        Creates a self-signed, non-exportable certificate for app-only sign-in.
    .DESCRIPTION
        The private key never leaves this host: the store is the only copy, and only the public
        key is uploaded to the app. Wrapped so tests never create real certificates.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$Store,
        [ValidateRange(1, 24)][int]$ValidMonths = 12
    )
    if (-not $PSCmdlet.ShouldProcess($Store, "Create certificate $Subject")) { return }
    New-SelfSignedCertificate -Subject $Subject -CertStoreLocation $Store -KeyExportPolicy NonExportable -KeySpec Signature `
        -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -NotAfter (Get-Date).AddMonths($ValidMonths)
}
