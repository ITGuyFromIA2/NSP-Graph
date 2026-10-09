function Register-NSPGraphAppOnlyRegistration {
    <#
    .SYNOPSIS
        Tenant bootstrap for unattended jobs: an app with application permissions and a certificate
        created on this host. Plan-only by default.
    .DESCRIPTION
        For scheduled jobs that run without a signed-in user. Keep -ApplicationPermissions to the
        minimum the job needs: an app-only token is not limited by any user's roles.

        Without -Execute this reports what it would do and changes nothing. With -Execute and
        ShouldProcess approval it creates or repairs, in order:
          - the application (single tenant, no secret) and its required application permissions;
          - its service principal;
          - an app role assignment (admin consent) for each missing permission;
          - a certificate CN=<AppName> in -CertificateStore, when none there is valid for more than
            -RenewWithinDays. It is self-signed, non-exportable, and valid for 12 months;
          - the app's certificate list: every valid CN=<AppName> certificate in the store. Graph
            replaces the whole list on update and never returns key material, so this host's
            store is the source of truth. One job host per tenant.
        Re-running within -RenewWithinDays of expiry renews the certificate. The old one stays
        valid until it expires.

        The certificate store must belong to the account the job runs as: Cert:\LocalMachine\My
        for SYSTEM (creating it needs an elevated session), or Cert:\CurrentUser\My for a user
        account. The record (<RecordDirectory>\<TenantId>.json) holds the client ID and thumbprint,
        never a key.
    .EXAMPLE
        Register-NSPGraphAppOnlyRegistration -AppName 'NSP-M365-ConditionalAccess-Coverage' -ApplicationPermissions 'User.Read.All', 'GroupMember.ReadWrite.All' -RecordDirectory $dir
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string[]]$ApplicationPermissions,
        [Parameter(Mandatory)][string]$RecordDirectory,
        [string]$CertificateStore = 'Cert:\LocalMachine\My',
        [ValidateRange(1, 180)][int]$RenewWithinDays = 30,
        [string]$TenantId,
        [switch]$Execute,
        [switch]$UseDeviceCode
    )

    $graphAppId = '00000003-0000-0000-c000-000000000000'
    $connectArgs = @{ Scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'Organization.Read.All'); Connect = $true }
    if ($UseDeviceCode) { $connectArgs.UseDeviceCode = $true }
    $context = Connect-NSPGraph @connectArgs
    if ($TenantId -and $context.TenantId -ne $TenantId) {
        throw "Connected to tenant $($context.TenantId), which does not match the requested TenantId $TenantId. Reconnect against the correct tenant."
    }
    $TenantId = $context.TenantId
    $recordPath = Join-Path $RecordDirectory "$TenantId.json"
    $subject = "CN=$AppName"

    function Get-NSPGraphItem {
        param($Map, [string]$Key)
        if ($null -eq $Map) { return }
        foreach ($item in @($Map[$Key])) { if ($null -ne $item) { $item } }
    }
    function Get-Filtered {
        param([string]$Collection, [string]$Filter, [string]$Select)
        @(Invoke-NSPGraphCollection -Uri "$Collection`?`$filter=$([uri]::EscapeDataString($Filter))&`$select=$Select")
    }
    function ConvertTo-KeyCredential {
        param($Certificate)
        @{
            type = 'AsymmetricX509Cert'; usage = 'Verify'
            key = [Convert]::ToBase64String($Certificate.RawData)
            displayName = "$subject $($Certificate.Thumbprint)"
        }
    }

    $tenantDomain = try {
        $organization = @((Invoke-NSPGraphRequest -Uri 'organization?$select=verifiedDomains')['value'])[0]
        @(Get-NSPGraphItem $organization 'verifiedDomains' | Where-Object { $_['isDefault'] })[0]['name']
    } catch { $null }

    $graphSp = @(Get-Filtered 'servicePrincipals' "appId eq '$graphAppId'" 'id,appRoles')[0]
    if (-not $graphSp) { throw 'The Microsoft Graph service principal could not be resolved in this tenant.' }
    $roles = foreach ($name in $ApplicationPermissions) {
        $role = @(Get-NSPGraphItem $graphSp 'appRoles' | Where-Object { $_['value'] -ceq $name })[0]
        if (-not $role) { throw "Application permission '$name' was not found on the Microsoft Graph service principal." }
        @{ Name = $name; Id = $role['id'] }
    }
    $requiredResourceAccess = @(@{ resourceAppId = $graphAppId; resourceAccess = @($roles | ForEach-Object { @{ id = $_.Id; type = 'Role' } }) })

    $record = if (Test-Path -LiteralPath $recordPath) { Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json } else { $null }
    $appSelect = 'id,appId,displayName,requiredResourceAccess,keyCredentials'
    $application = if ($record) { @(Get-Filtered 'applications' "appId eq '$($record.ClientId)'" $appSelect)[0] } else { $null }
    if (-not $application -and @(Get-Filtered 'applications' "displayName eq '$($AppName -replace "'", "''")'" 'appId').Count) {
        return [pscustomobject]@{
            Status = 'ExistingUnrecorded'; TenantId = $TenantId; TenantDomain = $tenantDomain; AppName = $AppName; RecordPath = $recordPath
            Message = "An app named '$AppName' already exists but is not recorded on this host. Nothing was created. Delete it, or register from the host that holds its record."
        }
    }

    $actions = [System.Collections.Generic.List[string]]::new()
    $servicePrincipal = $null
    $missingRoles = @($roles)
    $certificates = @(Get-NSPGraphAppCertificate -Subject $subject -Store $CertificateStore)
    $current = @($certificates | Where-Object { $_.NotAfter -gt (Get-Date).AddDays($RenewWithinDays) })[0]
    if (-not $current) { $actions.Add('CreateCertificate') }
    if ($application) {
        $clientId = $application['appId']
        $currentRoleIds = @(Get-NSPGraphItem $application 'requiredResourceAccess' | Where-Object { $_['resourceAppId'] -eq $graphAppId } |
                ForEach-Object { Get-NSPGraphItem $_ 'resourceAccess' } | ForEach-Object { $_['id'] })
        if (@($roles | Where-Object { $_.Id -notin $currentRoleIds }).Count) { $actions.Add('UpdateRequiredPermissions') }
        $servicePrincipal = @(Get-Filtered 'servicePrincipals' "appId eq '$clientId'" 'id,appId')[0]
        if (-not $servicePrincipal) { $actions.Add('CreateServicePrincipal') }
        else {
            $assigned = @(Invoke-NSPGraphCollection -Uri "servicePrincipals/$($servicePrincipal['id'])/appRoleAssignments" |
                    Where-Object { $_['resourceId'] -eq $graphSp['id'] } | ForEach-Object { $_['appRoleId'] })
            $missingRoles = @($roles | Where-Object { $_.Id -notin $assigned })
        }
        $uploaded = @(Get-NSPGraphItem $application 'keyCredentials' | ForEach-Object { $_['displayName'] })
        if ($actions -contains 'CreateCertificate' -or @($certificates | Where-Object { "$subject $($_.Thumbprint)" -notin $uploaded }).Count) { $actions.Add('UploadCertificates') }
    } else {
        $clientId = $null
        foreach ($action in 'CreateApplication', 'CreateServicePrincipal', 'UploadCertificates') { $actions.Add($action) }
    }
    if ($missingRoles.Count) { $actions.Add('AssignAppRoles') }

    $result = [ordered]@{
        Status = $null; TenantId = $TenantId; TenantDomain = $tenantDomain; AppName = $AppName; ClientId = $clientId
        Actions = @($actions); MissingPermissions = @($missingRoles | ForEach-Object Name)
        CertificateThumbprint = if ($current) { $current.Thumbprint } else { $null }
        CertificateNotAfter = if ($current) { $current.NotAfter.ToUniversalTime().ToString('o') } else { $null }
        CertificateStore = $CertificateStore; RecordPath = $recordPath
    }
    if ($actions.Count -eq 0) { $result.Status = 'AlreadyRegistered'; return [pscustomobject]$result }
    if (-not $Execute) {
        $result.Status = 'PlanOnly'
        $result.Message = "Run again with -Execute to apply: $($actions -join ', ')."
        return [pscustomobject]$result
    }
    if (-not $PSCmdlet.ShouldProcess("tenant $TenantId ($tenantDomain)", "App-only registration '$AppName': $($actions -join ', ')")) { return }

    if ('CreateCertificate' -in $actions) {
        $current = New-NSPGraphAppCertificate -Subject $subject -Store $CertificateStore -Confirm:$false
        $certificates = @(@($current) + @($certificates) | Sort-Object NotAfter -Descending)
    }
    if ('CreateApplication' -in $actions) {
        $application = Invoke-NSPGraphRequest -Method POST -Uri 'applications' -Body @{
            displayName = $AppName; signInAudience = 'AzureADMyOrg'; requiredResourceAccess = $requiredResourceAccess
        }
        $clientId = $application['appId']
    }
    $patch = @{}
    if ('UpdateRequiredPermissions' -in $actions) {
        $otherResources = @(Get-NSPGraphItem $application 'requiredResourceAccess' | Where-Object { $_['resourceAppId'] -ne $graphAppId })
        $patch.requiredResourceAccess = @($otherResources) + $requiredResourceAccess
    }
    if ('UploadCertificates' -in $actions) { $patch.keyCredentials = @($certificates | ForEach-Object { ConvertTo-KeyCredential $_ }) }
    if ($patch.Count) { Invoke-NSPGraphRequest -Method PATCH -Uri "applications/$($application['id'])" -Body $patch | Out-Null }
    if ('CreateServicePrincipal' -in $actions) {
        $servicePrincipal = Invoke-NSPGraphRequest -Method POST -Uri 'servicePrincipals' -Body @{ appId = $clientId }
    }
    foreach ($role in $missingRoles) {
        Invoke-NSPGraphRequest -Method POST -Uri "servicePrincipals/$($servicePrincipal['id'])/appRoleAssignments" -Body @{
            principalId = $servicePrincipal['id']; resourceId = $graphSp['id']; appRoleId = $role.Id
        } | Out-Null
    }

    New-Item -ItemType Directory -Path $RecordDirectory -Force | Out-Null
    [ordered]@{
        TenantId = $TenantId; TenantDomain = $tenantDomain; ClientId = $clientId; AppName = $AppName
        ServicePrincipalId = $servicePrincipal['id']
        ApplicationPermissions = @($ApplicationPermissions)
        CertificateThumbprint = $current.Thumbprint
        CertificateNotAfter = $current.NotAfter.ToUniversalTime().ToString('o')
        CertificateStore = $CertificateStore
        RecordedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $recordPath -Encoding UTF8

    $result.Status = if ('CreateApplication' -in $actions) { 'Created' } elseif ('CreateCertificate' -in $actions) { 'Renewed' } else { 'Repaired' }
    $result.ClientId = $clientId
    $result.CertificateThumbprint = $current.Thumbprint
    $result.CertificateNotAfter = $current.NotAfter.ToUniversalTime().ToString('o')
    [pscustomobject]$result
}
