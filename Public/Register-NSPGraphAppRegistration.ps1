function Register-NSPGraphAppRegistration {
    <#
    .SYNOPSIS
        Tenant bootstrap: ensures a single-tenant, delegated-only public-client app exists with
        admin consent for a tool's scope list, and records it locally. Plan-only by default.
    .DESCRIPTION
        Generalized from NSP-IntuneApps' Register-NSPIntuneWin32AppRegistration, over REST.

        Without -Execute this reports what it would do and changes nothing. With -Execute and
        ShouldProcess approval it creates or repairs, in order:
          - the application (single tenant, public client, no secret);
          - redirect URIs: http://localhost and the Windows broker (WAM) URI
            ms-appx-web://Microsoft.AAD.BrokerPlugin/<ClientId>, without which delegated sign-in as
            the app fails with AADSTS50011;
          - the required delegated permissions on the application;
          - its service principal;
          - an org-wide (AllPrincipals) delegated consent grant covering every scope in
            -DelegatedScopes. An existing grant is topped up, never narrowed.

        Identity comes from the record file <RecordDirectory>\<TenantId>.json. When there is no
        record but an app with the same display name exists, nothing is created or adopted
        automatically. Pass -AdoptClientId to take one over deliberately.

        The bootstrap itself connects with the SDK's default app and the scopes needed to create
        applications and grants. That is the one sign-in that does not use the tool's own app.
    .EXAMPLE
        Register-NSPGraphAppRegistration -AppName 'NSP-M365-ConditionalAccess' -DelegatedScopes $scopes -RecordDirectory .\Config\Local\AppRegistrations
    .EXAMPLE
        Register-NSPGraphAppRegistration -AppName 'NSP-M365-ConditionalAccess' -DelegatedScopes $scopes -RecordDirectory $dir -Execute
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string[]]$DelegatedScopes,
        [Parameter(Mandatory)][string]$RecordDirectory,
        [string]$TenantId,
        [string]$AdoptClientId,
        [switch]$Execute,
        [switch]$UseDeviceCode
    )

    $graphAppId = '00000003-0000-0000-c000-000000000000'
    $bootstrapScopes = @('Application.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All', 'Organization.Read.All')
    $connectArgs = @{ Scopes = $bootstrapScopes; Connect = $true }
    if ($UseDeviceCode) { $connectArgs.UseDeviceCode = $true }
    $context = Connect-NSPGraph @connectArgs
    if ($TenantId -and $context.TenantId -ne $TenantId) {
        throw "Connected to tenant $($context.TenantId), which does not match the requested TenantId $TenantId. Reconnect against the correct tenant."
    }
    $TenantId = $context.TenantId
    $recordPath = Join-Path $RecordDirectory "$TenantId.json"

    function Get-NSPGraphItem {
        # Null-safe: indexing or piping $null would otherwise throw or yield a phantom item.
        param($Map, [string]$Key)
        if ($null -eq $Map) { return }
        foreach ($item in @($Map[$Key])) { if ($null -ne $item) { $item } }
    }
    function Get-Filtered {
        param([string]$Collection, [string]$Filter, [string]$Select)
        @(Invoke-NSPGraphCollection -Uri "$Collection`?`$filter=$([uri]::EscapeDataString($Filter))&`$select=$Select")
    }

    $tenantDomain = try {
        $organization = @((Invoke-NSPGraphRequest -Uri 'organization?$select=verifiedDomains')['value'])[0]
        @(Get-NSPGraphItem $organization 'verifiedDomains' | Where-Object { $_['isDefault'] })[0]['name']
    } catch { $null }

    $graphSp = @(Get-Filtered 'servicePrincipals' "appId eq '$graphAppId'" 'id,oauth2PermissionScopes')[0]
    if (-not $graphSp) { throw 'The Microsoft Graph service principal could not be resolved in this tenant.' }
    $resolvedScopes = foreach ($name in $DelegatedScopes) {
        $scope = @(Get-NSPGraphItem $graphSp 'oauth2PermissionScopes' | Where-Object { $_['value'] -ceq $name })[0]
        if (-not $scope) { throw "Delegated permission '$name' was not found on the Microsoft Graph service principal." }
        @{ id = $scope['id']; type = 'Scope' }
    }
    $requiredResourceAccess = @(@{ resourceAppId = $graphAppId; resourceAccess = @($resolvedScopes) })

    $appSelect = 'id,appId,displayName,publicClient,requiredResourceAccess'
    $record = if (Test-Path -LiteralPath $recordPath) { Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json } else { $null }
    $application = $null
    $adopting = $false
    if ($record) {
        $application = @(Get-Filtered 'applications' "appId eq '$($record.ClientId)'" $appSelect)[0]
        if (-not $application) {
            Write-Warning "The recorded app $($record.ClientId) no longer exists in tenant $TenantId. A new registration is planned."
        }
    }
    if (-not $application) {
        $sameName = @(Get-Filtered 'applications' "displayName eq '$($AppName -replace "'", "''")'" $appSelect)
        if ($AdoptClientId) {
            $application = @($sameName | Where-Object { $_['appId'] -eq $AdoptClientId })[0]
            if (-not $application) { throw "-AdoptClientId $AdoptClientId is not an application named '$AppName' in this tenant." }
            $adopting = $true
        } elseif ($sameName.Count -gt 0) {
            return [pscustomobject]@{
                Status = 'ExistingUnrecorded'
                TenantId = $TenantId
                TenantDomain = $tenantDomain
                AppName = $AppName
                ExistingClientIds = @($sameName | ForEach-Object { $_['appId'] })
                RecordPath = $recordPath
                Message = "An app named '$AppName' already exists but is not recorded locally. Nothing was created. Review it, then rerun with -AdoptClientId <appId> to use it."
            }
        }
    }

    # Work out every change needed before changing anything.
    $actions = [System.Collections.Generic.List[string]]::new()
    $servicePrincipal = $null
    $grant = $null
    $missingScopes = @($DelegatedScopes)
    $desiredRedirects = @('http://localhost')
    if ($application) {
        $clientId = $application['appId']
        $desiredRedirects += "ms-appx-web://Microsoft.AAD.BrokerPlugin/$clientId"
        $currentRedirects = @(Get-NSPGraphItem (Get-NSPGraphItem $application 'publicClient') 'redirectUris')
        if (@($desiredRedirects | Where-Object { $_ -notin $currentRedirects }).Count) { $actions.Add('AddRedirectUris') }
        $currentScopeIds = @(Get-NSPGraphItem $application 'requiredResourceAccess' | Where-Object { $_['resourceAppId'] -eq $graphAppId } |
                ForEach-Object { Get-NSPGraphItem $_ 'resourceAccess' } | ForEach-Object { $_['id'] })
        if (@($resolvedScopes | Where-Object { $_.id -notin $currentScopeIds }).Count) { $actions.Add('UpdateRequiredPermissions') }

        $servicePrincipal = @(Get-Filtered 'servicePrincipals' "appId eq '$clientId'" 'id,appId')[0]
        if (-not $servicePrincipal) {
            $actions.Add('CreateServicePrincipal')
        } else {
            # Consent grants are keyed by the service principal's object ID, not the application's.
            $grant = @(Get-Filtered 'oauth2PermissionGrants' "clientId eq '$($servicePrincipal['id'])' and resourceId eq '$($graphSp['id'])' and consentType eq 'AllPrincipals'" 'id,scope')[0]
            $granted = if ($grant) { @([string]$grant['scope'] -split '\s+' | Where-Object { $_ }) } else { @() }
            $missingScopes = @($DelegatedScopes | Where-Object { $_ -notin $granted })
        }
        if ($missingScopes.Count) { $actions.Add('GrantConsent') }
    } else {
        $clientId = $null
        foreach ($action in 'CreateApplication', 'AddRedirectUris', 'CreateServicePrincipal', 'GrantConsent') { $actions.Add($action) }
    }
    if ($adopting) { $actions.Add('RecordAdoption') }

    $result = [ordered]@{
        Status = $null
        TenantId = $TenantId
        TenantDomain = $tenantDomain
        AppName = $AppName
        ClientId = $clientId
        Actions = @($actions)
        MissingScopes = $missingScopes
        RecordPath = $recordPath
    }
    if ($actions.Count -eq 0) {
        $result.Status = 'AlreadyRegistered'
        return [pscustomobject]$result
    }
    if (-not $Execute) {
        $result.Status = 'PlanOnly'
        $result.Message = "Run again with -Execute to apply: $($actions -join ', ')."
        return [pscustomobject]$result
    }
    if (-not $PSCmdlet.ShouldProcess("tenant $TenantId ($tenantDomain)", "App registration '$AppName': $($actions -join ', ')")) { return }

    if ('CreateApplication' -in $actions) {
        $application = Invoke-NSPGraphRequest -Method POST -Uri 'applications' -Body @{
            displayName = $AppName
            signInAudience = 'AzureADMyOrg'
            isFallbackPublicClient = $true
            publicClient = @{ redirectUris = @('http://localhost') }
            requiredResourceAccess = $requiredResourceAccess
        }
        $clientId = $application['appId']
        $desiredRedirects = @('http://localhost', "ms-appx-web://Microsoft.AAD.BrokerPlugin/$clientId")
    }
    $patch = @{}
    if ('AddRedirectUris' -in $actions) {
        $patch.publicClient = @{ redirectUris = @(@(Get-NSPGraphItem (Get-NSPGraphItem $application 'publicClient') 'redirectUris') + $desiredRedirects | Select-Object -Unique) }
    }
    if ('UpdateRequiredPermissions' -in $actions) {
        $otherResources = @(Get-NSPGraphItem $application 'requiredResourceAccess' | Where-Object { $_['resourceAppId'] -ne $graphAppId })
        $patch.requiredResourceAccess = @($otherResources) + $requiredResourceAccess
    }
    if ($patch.Count) {
        Invoke-NSPGraphRequest -Method PATCH -Uri "applications/$($application['id'])" -Body $patch | Out-Null
    }
    if ('CreateServicePrincipal' -in $actions) {
        $servicePrincipal = Invoke-NSPGraphRequest -Method POST -Uri 'servicePrincipals' -Body @{ appId = $clientId }
    }
    if ('GrantConsent' -in $actions) {
        if ($grant) {
            $merged = (@([string]$grant['scope'] -split '\s+' | Where-Object { $_ }) + $missingScopes | Select-Object -Unique) -join ' '
            Invoke-NSPGraphRequest -Method PATCH -Uri "oauth2PermissionGrants/$($grant['id'])" -Body @{ scope = $merged } | Out-Null
        } else {
            Invoke-NSPGraphRequest -Method POST -Uri 'oauth2PermissionGrants' -Body @{
                clientId = $servicePrincipal['id']
                consentType = 'AllPrincipals'
                resourceId = $graphSp['id']
                scope = ($DelegatedScopes -join ' ')
            } | Out-Null
        }
    }

    New-Item -ItemType Directory -Path $RecordDirectory -Force | Out-Null
    [ordered]@{
        TenantId = $TenantId
        TenantDomain = $tenantDomain
        ClientId = $clientId
        AppName = $AppName
        RecordedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        GrantedScopes = @($DelegatedScopes)
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $recordPath -Encoding UTF8

    $result.Status = if ('CreateApplication' -in $actions) { 'Created' } elseif ($adopting) { 'Adopted' } else { 'Repaired' }
    $result.ClientId = $clientId
    [pscustomobject]$result
}
