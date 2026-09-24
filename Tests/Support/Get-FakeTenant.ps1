# In-memory tenant for the Fake transport: just enough of applications, servicePrincipals, and
# oauth2PermissionGrants for app-registration tests. Returns @{ State; Responder }.
function Get-FakeTenant {
    param([string[]]$GraphScopes = @('Policy.Read.All', 'Policy.ReadWrite.ConditionalAccess', 'Group.Read.All'))

    $state = @{
        Counter = 0
        Applications = [System.Collections.Generic.List[hashtable]]::new()
        ServicePrincipals = [System.Collections.Generic.List[hashtable]]::new()
        Grants = [System.Collections.Generic.List[hashtable]]::new()
    }
    $scopeObjects = @($GraphScopes | ForEach-Object { @{ value = $_; id = "scope-id-$_" } })
    $state.ServicePrincipals.Add(@{ id = 'graph-sp'; appId = '00000003-0000-0000-c000-000000000000'; oauth2PermissionScopes = $scopeObjects })

    $responder = {
        param($Method, $Uri, $Body)
        $decoded = [uri]::UnescapeDataString($Uri)
        $path = ($decoded -replace '^v1\.0/', '') -replace '\?.*$', ''
        $filter = if ($decoded -match '\$filter=([^&]+)') { $Matches[1] } else { $null }
        $state.Counter++
        $n = $state.Counter

        function Test-Filter {
            param([hashtable]$Item, [string]$Filter)
            if (-not $Filter) { return $true }
            foreach ($clause in $Filter -split ' and ') {
                if ($clause -notmatch "^(\w+) eq '(.*)'$") { throw "Fake tenant cannot evaluate filter clause: $clause" }
                if ([string]$Item[$Matches[1]] -ne ($Matches[2] -replace "''", "'")) { return $false }
            }
            $true
        }

        $collections = @{ applications = $state.Applications; servicePrincipals = $state.ServicePrincipals; oauth2PermissionGrants = $state.Grants }
        switch ($Method) {
            'GET' {
                if ($path -eq 'organization') { return @{ value = @(@{ verifiedDomains = @(@{ name = 'fixture.onmicrosoft.com'; isDefault = $true }) }) } }
                return @{ value = @($collections[$path] | Where-Object { Test-Filter $_ $filter }) }
            }
            'POST' {
                $item = @{} + $Body
                switch ($path) {
                    'applications' { $item.id = "app-object-$n"; $item.appId = "client-$n" }
                    'servicePrincipals' { $item.id = "sp-$n" }
                    'oauth2PermissionGrants' { $item.id = "grant-$n" }
                }
                $collections[$path].Add($item)
                return $item
            }
            'PATCH' {
                $collection, $id = $path -split '/', 2
                $item = @($collections[$collection] | Where-Object { $_.id -eq $id })[0]
                if (-not $item) { throw "Fake tenant: $collection/$id not found" }
                foreach ($key in $Body.Keys) { $item[$key] = $Body[$key] }
                return $null
            }
        }
        throw "Fake tenant cannot handle $Method $path"
    }.GetNewClosure()

    @{ State = $state; Responder = $responder }
}
