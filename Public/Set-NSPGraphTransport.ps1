function Set-NSPGraphTransport {
    <#
    .SYNOPSIS
        Selects how Invoke-NSPGraphRequest reaches Microsoft Graph for this session.
    .DESCRIPTION
        MgGraphRequest (default) sends requests with Invoke-MgGraphRequest on the current
        Connect-MgGraph context.

        Fake sends nothing. Each request is recorded in the call log and answered by -Responder,
        a scriptblock receiving (Method, Uri, Body) and returning a hashtable shaped like the
        Graph JSON response. Use it to test planners and executors without a tenant.

        Resource-specific cmdlet backends (for example Get-MgIdentityConditionalAccessPolicy)
        belong in the consuming module, because they differ per resource.
    .EXAMPLE
        Set-NSPGraphTransport -Name Fake -Responder { param($Method, $Uri, $Body) @{ value = @() } }
    .EXAMPLE
        Set-NSPGraphTransport -Name MgGraphRequest
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateSet('MgGraphRequest', 'Fake')][string]$Name,
        [scriptblock]$Responder
    )

    if ($Name -eq 'Fake' -and -not $Responder) {
        throw 'The Fake transport requires -Responder.'
    }
    if ($PSCmdlet.ShouldProcess("Graph transport", "Use $Name")) {
        $script:NSPGraphTransport = @{
            Name = $Name
            Responder = if ($Name -eq 'Fake') { $Responder } else { $null }
            CallLog = [System.Collections.Generic.List[object]]::new()
        }
    }
}
