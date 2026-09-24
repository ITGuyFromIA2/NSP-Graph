BeforeAll {
    $script:manifestPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'NSP.M365.Graph.psd1'
    Import-Module $script:manifestPath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module NSP.M365.Graph -Force -ErrorAction SilentlyContinue
}

Describe 'Module manifest' {
    It 'exports exactly the Public folder functions, each with a synopsis and example' {
        $publicPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Public'
        $expected = @(Get-ChildItem -LiteralPath $publicPath -Filter '*.ps1' | Select-Object -ExpandProperty BaseName)
        $declared = @((Import-PowerShellDataFile -Path $script:manifestPath).FunctionsToExport)
        $exported = @(Get-Command -Module NSP.M365.Graph | Select-Object -ExpandProperty Name)
        @(Compare-Object $expected $declared).Count | Should -Be 0
        @(Compare-Object $expected $exported).Count | Should -Be 0
        foreach ($name in $exported) {
            $help = Get-Help $name -Full
            $help.Synopsis | Should -Not -BeNullOrEmpty -Because "$name needs .SYNOPSIS"
            @($help.Examples.Example).Count | Should -BeGreaterThan 0 -Because "$name needs .EXAMPLE"
        }
    }
}

Describe 'Graph transport' {
    AfterEach {
        Set-NSPGraphTransport -Name MgGraphRequest
    }

    It 'defaults to MgGraphRequest' {
        Import-Module $script:manifestPath -Force
        (Get-NSPGraphTransport).Name | Should -Be 'MgGraphRequest'
    }

    It 'requires a responder for the Fake transport' {
        { Set-NSPGraphTransport -Name Fake } | Should -Throw '*requires -Responder*'
    }

    It 'answers from the responder, records the call, and never calls Invoke-MgGraphRequest' {
        Mock Invoke-MgGraphRequest { throw 'must not be called' } -ModuleName NSP.M365.Graph
        Set-NSPGraphTransport -Name Fake -Responder { param($Method, $Uri, $Body) @{ id = 'new-1'; method = $Method; echoed = $Body.state; uri = $Uri } }

        $result = Invoke-NSPGraphRequest -Method POST -Uri '/identity/conditionalAccess/policies' -Body @{ state = 'disabled' }

        $result.id | Should -Be 'new-1'
        $result.echoed | Should -Be 'disabled'
        $result.uri | Should -Be 'v1.0/identity/conditionalAccess/policies'
        $log = (Get-NSPGraphTransport).CallLog
        @($log).Count | Should -Be 1
        $log[0].Method | Should -Be 'POST'
        Should -Invoke Invoke-MgGraphRequest -Times 0 -ModuleName NSP.M365.Graph
    }

    It 'resets the call log when the transport is set again' {
        Set-NSPGraphTransport -Name Fake -Responder { @{} }
        Invoke-NSPGraphRequest -Uri 'organization' | Out-Null
        Set-NSPGraphTransport -Name Fake -Responder { @{} }
        @((Get-NSPGraphTransport).CallLog).Count | Should -Be 0
    }

    It 'prefixes beta for relative URIs and leaves absolute URIs unchanged' {
        Set-NSPGraphTransport -Name Fake -Responder { @{ uri = $args[1] } }
        (Invoke-NSPGraphRequest -Uri 'identity/conditionalAccess/policies' -ApiVersion beta).uri | Should -Be 'beta/identity/conditionalAccess/policies'
        $absolute = 'https://graph.microsoft.com/v1.0/groups?$skiptoken=abc'
        (Invoke-NSPGraphRequest -Uri $absolute).uri | Should -Be $absolute
    }
}

Describe 'Invoke-MgGraphRequest transport' {
    It 'sends a JSON body with hashtable output' {
        Mock Invoke-MgGraphRequest { @{ ok = $true } } -ModuleName NSP.M365.Graph
        Invoke-NSPGraphRequest -Method PATCH -Uri 'identity/conditionalAccess/policies/p1' -Body @{ state = 'enabled' } | Out-Null
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ModuleName NSP.M365.Graph -ParameterFilter {
            $Method -eq 'PATCH' -and
            $Uri -eq 'v1.0/identity/conditionalAccess/policies/p1' -and
            $OutputType -eq 'HashTable' -and
            $ContentType -eq 'application/json' -and
            (ConvertFrom-Json $Body).state -eq 'enabled'
        }
    }

    It 'sends no body for a GET' {
        Mock Invoke-MgGraphRequest { @{ value = @() } } -ModuleName NSP.M365.Graph
        Invoke-NSPGraphRequest -Uri 'organization' | Out-Null
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ModuleName NSP.M365.Graph -ParameterFilter { $null -eq $Body }
    }
}

Describe 'Invoke-NSPGraphCollection' {
    AfterEach {
        Set-NSPGraphTransport -Name MgGraphRequest
    }

    It 'follows nextLink across pages' {
        Set-NSPGraphTransport -Name Fake -Responder {
            $Uri = $args[1]
            if ($Uri -eq 'v1.0/groups') {
                @{ value = @(@{ id = 'g1' }, @{ id = 'g2' }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/groups?page=2' }
            } else {
                @{ value = @(@{ id = 'g3' }) }
            }
        }
        $items = @(Invoke-NSPGraphCollection -Uri 'groups')
        $items.Count | Should -Be 3
        $items[2].id | Should -Be 'g3'
        @((Get-NSPGraphTransport).CallLog).Count | Should -Be 2
    }

    It 'returns nothing for an empty collection' {
        Set-NSPGraphTransport -Name Fake -Responder { @{ value = @() } }
        @(Invoke-NSPGraphCollection -Uri 'groups').Count | Should -Be 0
    }
}
