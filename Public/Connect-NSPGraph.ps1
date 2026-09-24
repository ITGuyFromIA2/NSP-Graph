function Connect-NSPGraph {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph, or validates an existing delegated context, with the required scopes.
    .DESCRIPTION
        Extracted from NSP-IntuneApps (Private\Connect-NSPGraph.ps1) so every NSP M365 tool shares
        one implementation.

        Pass -ClientId/-TenantId to authenticate as a registered app instead of the Microsoft Graph
        PowerShell SDK's default app. A pre-consented registration can sign in silently.

        -Connect only calls Connect-MgGraph when the existing session is insufficient: missing a
        scope, or authenticated as a different app or tenant than requested. Reconnecting
        unconditionally causes repeated sign-in prompts even when a good session already exists.
        Callers should always pass their tool's single, complete scope list.

        -UseDeviceCode requests the device-code flow instead of the Windows broker (WAM). The WAM
        popup can open behind other windows and time out, which surfaces as
        "User canceled authentication".

        Connect-MgGraph keeps one context per process, so this serves one tenant at a time.
    .EXAMPLE
        Connect-NSPGraph -Scopes 'Policy.Read.All', 'Group.Read.All' -Connect
    .EXAMPLE
        $context = Connect-NSPGraph -Scopes 'Policy.Read.All' -Optional
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Scopes,
        [switch]$Connect,
        [switch]$Optional,
        [string]$ClientId,
        [string]$TenantId,
        [switch]$UseDeviceCode
    )

    if ($Connect) {
        if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
            if (-not (Get-Module -ListAvailable NSP.Bootstrap)) {
                throw 'Microsoft.Graph.Authentication is not installed. Install it, or install NSP.Bootstrap so it can be installed for you.'
            }
            Import-Module NSP.Bootstrap -ErrorAction Stop
            Install-NSPModule -Name Microsoft.Graph.Authentication -Scope CurrentUser
        }
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

        $existingContext = if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) { Get-MgContext } else { $null }
        $hasRequiredScopes = $existingContext -and (@($Scopes | Where-Object { $_ -notin @($existingContext.Scopes) })).Count -eq 0
        $matchesRequestedApp = (-not $ClientId -or $existingContext.ClientId -eq $ClientId) -and (-not $TenantId -or $existingContext.TenantId -eq $TenantId)

        if (-not ($existingContext -and $hasRequiredScopes -and $matchesRequestedApp)) {
            $connectArgs = @{ Scopes = $Scopes; NoWelcome = $true }
            if ($ClientId) { $connectArgs.ClientId = $ClientId }
            if ($TenantId) { $connectArgs.TenantId = $TenantId }
            if ($UseDeviceCode) { $connectArgs.UseDeviceCode = $true }
            Connect-MgGraph @connectArgs | Out-Null
        }
    }

    $context = if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) { Get-MgContext } else { $null }
    if (-not $context) {
        if ($Optional) { return $null }
        throw 'No Microsoft Graph context is available. Use -Connect for delegated interactive login.'
    }
    if (-not $Optional) {
        $missingScopes = @($Scopes | Where-Object { $_ -notin @($context.Scopes) })
        if (@($missingScopes).Count -gt 0) {
            throw "The current Graph context lacks required scope(s): $($missingScopes -join ', '). Reconnect with -Connect."
        }
    }
    $context
}
