BeforeAll {
    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'NSP.M365.Graph.psd1') -Force -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'Support\Get-FakeTenant.ps1')

    $script:appName = 'NSP-M365-Test-Job'
    $script:permissions = @('User.Read.All', 'GroupMember.ReadWrite.All')

    function Use-FakeTenant {
        $tenant = Get-FakeTenant
        Set-NSPGraphTransport -Name Fake -Responder $tenant.Responder
        $tenant.State
    }
    function Get-Write { @((Get-NSPGraphTransport).CallLog | Where-Object Method -ne 'GET') }
    function Get-TestCertificate {
        param([string]$Thumbprint, [int]$Days)
        [pscustomobject]@{ Subject = "CN=$script:appName"; Thumbprint = $Thumbprint; NotAfter = (Get-Date).AddDays($Days); RawData = [Text.Encoding]::ASCII.GetBytes($Thumbprint); HasPrivateKey = $true }
    }
    function Invoke-Register {
        param([string[]]$Permissions = $script:permissions, [switch]$Execute)
        $registerArgs = @{
            AppName = $script:appName; ApplicationPermissions = $Permissions
            RecordDirectory = Join-Path $TestDrive 'records'; CertificateStore = 'Cert:\CurrentUser\My'; Confirm = $false
        }
        if ($Execute) { $registerArgs.Execute = $true }
        Register-NSPGraphAppOnlyRegistration @registerArgs
    }
}

AfterAll {
    Remove-Module NSP.M365.Graph -Force -ErrorAction SilentlyContinue
}

Describe 'Register-NSPGraphAppOnlyRegistration' {
    BeforeEach {
        Mock Connect-NSPGraph { [pscustomobject]@{ TenantId = 'tenant-1' } } -ModuleName NSP.M365.Graph
        Mock Get-NSPGraphAppCertificate { @() } -ModuleName NSP.M365.Graph
        Mock New-NSPGraphAppCertificate { Get-TestCertificate 'NEW0000000000000000000000000000000000001' 365 } -ModuleName NSP.M365.Graph
        Remove-Item (Join-Path $TestDrive 'records') -Recurse -Force -ErrorAction SilentlyContinue
    }
    AfterEach { Set-NSPGraphTransport -Name MgGraphRequest }

    It 'plans without creating a certificate or writing anything' {
        Use-FakeTenant | Out-Null
        $result = Invoke-Register
        $result.Status | Should -Be 'PlanOnly'
        $result.Actions | Should -Be @('CreateCertificate', 'CreateApplication', 'CreateServicePrincipal', 'UploadCertificates', 'AssignAppRoles')
        Get-Write | Should -BeNullOrEmpty
        Should -Invoke New-NSPGraphAppCertificate -Times 0 -ModuleName NSP.M365.Graph
    }

    It 'creates the app with application permissions, uploads only the public key, and consents each role' {
        $state = Use-FakeTenant
        $result = Invoke-Register -Execute
        $result.Status | Should -Be 'Created'
        $app = $state.Applications[0]
        $app.requiredResourceAccess[0].resourceAccess.type | Sort-Object -Unique | Should -Be 'Role'
        $app.keyCredentials.Count | Should -Be 1
        $app.keyCredentials[0].key | Should -Be ([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes('NEW0000000000000000000000000000000000001')))
        $app.keyCredentials[0].type | Should -Be 'AsymmetricX509Cert'
        $sp = @($state.ServicePrincipals | Where-Object appId -eq $app.appId)[0]
        @($state.RoleAssignments | ForEach-Object appRoleId) | Should -Be @('role-id-User.Read.All', 'role-id-GroupMember.ReadWrite.All')
        @($state.RoleAssignments | ForEach-Object principalId | Sort-Object -Unique) | Should -Be @($sp.id)
        $record = Get-Content (Join-Path $TestDrive 'records\tenant-1.json') -Raw | ConvertFrom-Json
        $record.CertificateThumbprint | Should -Be 'NEW0000000000000000000000000000000000001'
        $record.ServicePrincipalId | Should -Be $sp.id
        ($record | ConvertTo-Json) | Should -Not -Match 'key'
    }

    It 'is idempotent once registered with a current certificate' {
        Use-FakeTenant | Out-Null
        Invoke-Register -Execute | Out-Null
        Mock Get-NSPGraphAppCertificate { @(Get-TestCertificate 'NEW0000000000000000000000000000000000001' 365) } -ModuleName NSP.M365.Graph
        (Invoke-Register).Status | Should -Be 'AlreadyRegistered'
    }

    It 'renews a certificate near expiry and keeps the old one uploaded until it expires' {
        $state = Use-FakeTenant
        Invoke-Register -Execute | Out-Null
        Mock Get-NSPGraphAppCertificate { @(Get-TestCertificate 'OLD0000000000000000000000000000000000002' 20) } -ModuleName NSP.M365.Graph
        Mock New-NSPGraphAppCertificate { Get-TestCertificate 'NEW0000000000000000000000000000000000003' 365 } -ModuleName NSP.M365.Graph
        $result = Invoke-Register -Execute
        $result.Status | Should -Be 'Renewed'
        $result.CertificateThumbprint | Should -Be 'NEW0000000000000000000000000000000000003'
        @($state.Applications[0].keyCredentials | ForEach-Object displayName) | Should -Be @(
            "CN=$script:appName NEW0000000000000000000000000000000000003", "CN=$script:appName OLD0000000000000000000000000000000000002")
    }

    It 'refuses a permission Graph does not offer, and an unrecorded app of the same name' {
        Use-FakeTenant | Out-Null
        { Invoke-Register -Permissions 'Directory.ReadWrite.Everything' } | Should -Throw '*not found*'
        $state = Use-FakeTenant
        $state.Applications.Add(@{ id = 'app-x'; appId = 'client-x'; displayName = $script:appName })
        (Invoke-Register).Status | Should -Be 'ExistingUnrecorded'
    }
}
