function Invoke-NSPGraphRequest {
    <#
    .SYNOPSIS
        Sends one Microsoft Graph REST request through the active transport and returns hashtables.
    .DESCRIPTION
        The canonical shape for NSP M365 tools is the Graph REST JSON shape: camelCase keys in
        hashtables. Request bodies are hashtables in that same shape.

        -Uri may be relative ('identity/conditionalAccess/policies'), in which case -ApiVersion
        is prefixed, or absolute (an @odata.nextLink), which is used unchanged.

        This function does not ask for confirmation. Callers that write to a tenant are expected
        to gate the call behind their own approved plan.
    .EXAMPLE
        Invoke-NSPGraphRequest -Uri 'identity/conditionalAccess/policies/00000000-0000-0000-0000-000000000000'
    .EXAMPLE
        Invoke-NSPGraphRequest -Method PATCH -Uri "identity/conditionalAccess/policies/$id" -Body @{ state = 'enabled' }
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')][string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('v1.0', 'beta')][string]$ApiVersion = 'v1.0',
        [object]$Body
    )

    $resolvedUri = if ($Uri -match '^https://') { $Uri } else { "$ApiVersion/$($Uri.TrimStart('/'))" }
    $transport = $script:NSPGraphTransport

    if ($transport.Name -eq 'Fake') {
        $transport.CallLog.Add([pscustomobject]@{ Method = $Method; Uri = $resolvedUri; Body = $Body })
        return & $transport.Responder $Method $resolvedUri $Body
    }

    $requestArgs = @{
        Method = $Method
        Uri = $resolvedUri
        OutputType = 'HashTable'
        ErrorAction = 'Stop'
    }
    if ($PSBoundParameters.ContainsKey('Body')) {
        $requestArgs.Body = if ($Body -is [string]) { $Body } else { ConvertTo-Json -InputObject $Body -Depth 32 -Compress }
        $requestArgs.ContentType = 'application/json'
    }
    Invoke-MgGraphRequest @requestArgs
}
