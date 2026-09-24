BeforeAll {
    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'NSP.M365.Graph.psd1') -Force -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'Support\Get-FakeTenant.ps1')

    $script:appName = 'NSP-M365-Test'
    $script:scopes = @('Policy.Read.All', 'Policy.ReadWrite.ConditionalAccess')

    function Use-FakeTenant {
        param([string[]]$GraphScopes = @('Policy.Read.All', 'Policy.ReadWrite.ConditionalAccess', 'Group.Read.All'))
        $tenant = Get-FakeTenant -GraphScopes $GraphScopes
        Set-NSPGraphTransport -Name Fake -Responder $tenant.Responder
        $tenant.State
    }

    function Get-Write {
        @((Get-NSPGraphTransport).CallLog | Where-Object Method -ne 'GET')
    }

    function Invoke-Register {
        param([string[]]$Scopes = $script:scopes, [switch]$Execute, [string]$AdoptClientId)
        $registerArgs = @{
            AppName = $script:appName
            DelegatedScopes = $Scopes
            RecordDirectory = Join-Path $TestDrive 'records'
            Confirm = $false
        }
        if ($Execute) { $registerArgs.Execute = $true }
        if ($AdoptClientId) { $registerArgs.AdoptClientId = $AdoptClientId }
        Register-NSPGraphAppRegistration @registerArgs
    }
}

AfterAll {
    Remove-Module NSP.M365.Graph -Force -ErrorAction SilentlyContinue
}

Describe 'Register-NSPGraphAppRegistration' {
    BeforeEach {
        Mock Connect-NSPGraph { [pscustomobject]@{ TenantId = 'tenant-1' } } -ModuleName NSP.M365.Graph
        Remove-Item (Join-Path $TestDrive 'records') -Recurse -Force -ErrorAction SilentlyContinue
    }

    AfterEach {
        Set-NSPGraphTransport -Name MgGraphRequest
    }

    It 'plans without writing anything when -Execute is not passed' {
        Use-FakeTenant | Out-Null
        $result = Invoke-Register

        $result.Status | Should -Be 'PlanOnly'
        $result.Actions | Should -Be @('CreateApplication', 'AddRedirectUris', 'CreateServicePrincipal', 'GrantConsent')
        $result.TenantDomain | Should -Be 'fixture.onmicrosoft.com'
        Get-Write | Should -BeNullOrEmpty
        Test-Path (Join-Path $TestDrive 'records\tenant-1.json') | Should -BeFalse
    }

    It 'creates the app, redirect URIs, service principal, and a consent grant keyed to the service principal' {
        $state = Use-FakeTenant
        $result = Invoke-Register -Execute

        $result.Status | Should -Be 'Created'
        $app = $state.Applications[0]
        $app.signInAudience | Should -Be 'AzureADMyOrg'
        $app.isFallbackPublicClient | Should -BeTrue
        $app.publicClient.redirectUris | Should -Be @('http://localhost', "ms-appx-web://Microsoft.AAD.BrokerPlugin/$($app.appId)")

        $sp = @($state.ServicePrincipals | Where-Object { $_.appId -eq $app.appId })[0]
        $sp | Should -Not -BeNullOrEmpty
        $grant = $state.Grants[0]
        $grant.clientId | Should -Be $sp.id -Because 'grants reference the service principal, not the application object'
        $grant.clientId | Should -Not -Be $app.id
        $grant.resourceId | Should -Be 'graph-sp'
        $grant.consentType | Should -Be 'AllPrincipals'
        $grant.scope | Should -Be 'Policy.Read.All Policy.ReadWrite.ConditionalAccess'

        $record = Get-Content (Join-Path $TestDrive 'records\tenant-1.json') -Raw | ConvertFrom-Json
        $record.ClientId | Should -Be $app.appId
        $record.TenantDomain | Should -Be 'fixture.onmicrosoft.com'
    }

    It 'reports AlreadyRegistered on a rerun and writes nothing' {
        Use-FakeTenant | Out-Null
        Invoke-Register -Execute | Out-Null
        $transport = Get-NSPGraphTransport
        $writesBefore = @($transport.CallLog | Where-Object Method -ne 'GET').Count

        (Invoke-Register -Execute).Status | Should -Be 'AlreadyRegistered'
        (Get-Write).Count | Should -Be $writesBefore
    }

    It 'tops up an existing grant when the scope list grows, without replacing it' {
        $state = Use-FakeTenant
        Invoke-Register -Execute | Out-Null

        $plan = Invoke-Register -Scopes ($script:scopes + 'Group.Read.All')
        $plan.Status | Should -Be 'PlanOnly'
        $plan.Actions | Should -Be @('UpdateRequiredPermissions', 'GrantConsent')
        $plan.MissingScopes | Should -Be @('Group.Read.All')

        (Invoke-Register -Scopes ($script:scopes + 'Group.Read.All') -Execute).Status | Should -Be 'Repaired'
        $state.Grants.Count | Should -Be 1
        $state.Grants[0].scope | Should -Be 'Policy.Read.All Policy.ReadWrite.ConditionalAccess Group.Read.All'
        @($state.Applications[0].requiredResourceAccess[0].resourceAccess).Count | Should -Be 3
    }

    It 'restores a missing broker redirect URI and keeps existing ones' {
        $state = Use-FakeTenant
        Invoke-Register -Execute | Out-Null
        $state.Applications[0].publicClient = @{ redirectUris = @('http://localhost') }

        (Invoke-Register).Actions | Should -Be @('AddRedirectUris')
        Invoke-Register -Execute | Out-Null
        $state.Applications[0].publicClient.redirectUris | Should -Contain 'http://localhost'
        $state.Applications[0].publicClient.redirectUris | Should -Contain "ms-appx-web://Microsoft.AAD.BrokerPlugin/$($state.Applications[0].appId)"
    }

    It 'does not create or adopt an unrecorded app with the same name unless told to' {
        $state = Use-FakeTenant
        $state.Applications.Add(@{ id = 'app-object-x'; appId = 'client-x'; displayName = $script:appName; publicClient = @{ redirectUris = @() }; requiredResourceAccess = @() })

        $result = Invoke-Register -Execute
        $result.Status | Should -Be 'ExistingUnrecorded'
        $result.ExistingClientIds | Should -Be @('client-x')
        Get-Write | Should -BeNullOrEmpty

        $adopted = Invoke-Register -Execute -AdoptClientId 'client-x'
        $adopted.Status | Should -Be 'Adopted'
        $adopted.ClientId | Should -Be 'client-x'
        $state.Applications.Count | Should -Be 1
        (Get-Content (Join-Path $TestDrive 'records\tenant-1.json') -Raw | ConvertFrom-Json).ClientId | Should -Be 'client-x'
    }

    It 'rejects -AdoptClientId that is not an app with that name' {
        Use-FakeTenant | Out-Null
        { Invoke-Register -Execute -AdoptClientId 'nope' } | Should -Throw '*is not an application named*'
    }

    It 'throws when connected to a different tenant than requested' {
        Use-FakeTenant | Out-Null
        { Register-NSPGraphAppRegistration -AppName 'x' -DelegatedScopes 'Policy.Read.All' -RecordDirectory $TestDrive -TenantId 'tenant-2' } |
            Should -Throw '*does not match*'
    }

    It 'throws for a scope Graph does not publish' {
        Use-FakeTenant | Out-Null
        { Invoke-Register -Scopes 'Not.A.Scope' } | Should -Throw "*'Not.A.Scope' was not found*"
    }
}
