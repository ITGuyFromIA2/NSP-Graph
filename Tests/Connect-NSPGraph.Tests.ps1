Describe 'Connect-NSPGraph' {
    BeforeAll {
        Import-Module (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'NSP.M365.Graph.psd1') -Force
    }

    Context 'without an existing Graph context' {
        It 'throws when a context is required' {
            InModuleScope NSP.M365.Graph {
                { Connect-NSPGraph -Scopes 'DeviceManagementApps.Read.All' } | Should -Throw '*No Microsoft Graph context is available*'
            }
        }

        It 'returns $null when the context is optional' {
            InModuleScope NSP.M365.Graph {
                Connect-NSPGraph -Scopes 'DeviceManagementApps.Read.All' -Optional | Should -Be $null
            }
        }
    }

    Context 'with an existing Graph context' {
        It 'returns the context when all required scopes are present' {
            InModuleScope NSP.M365.Graph {
                function Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; Account = 'operator'; Scopes = @('DeviceManagementApps.Read.All', 'Group.Read.All') } }
                $context = Connect-NSPGraph -Scopes 'DeviceManagementApps.Read.All'
                $context.TenantId | Should -Be 'tenant-1'
            }
        }

        It 'throws when a required scope is missing and the context is required' {
            InModuleScope NSP.M365.Graph {
                function Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; Account = 'operator'; Scopes = @('Group.Read.All') } }
                { Connect-NSPGraph -Scopes 'DeviceManagementApps.Read.All' } | Should -Throw '*lacks required scope*'
            }
        }

        It 'returns the context without validating scopes when optional' {
            InModuleScope NSP.M365.Graph {
                function Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; Account = 'operator'; Scopes = @() } }
                $context = Connect-NSPGraph -Scopes 'DeviceManagementApps.Read.All' -Optional
                $context.TenantId | Should -Be 'tenant-1'
            }
        }
    }

    Context 'connecting with a registered app instead of the SDK default' {
        It 'passes ClientId and TenantId through to Connect-MgGraph when there is no existing context' {
            InModuleScope NSP.M365.Graph {
                Mock Connect-MgGraph { } -ModuleName NSP.M365.Graph
                function Get-MgContext { $null }

                Connect-NSPGraph -Scopes 'DeviceManagementApps.Read.All' -Connect -ClientId 'client-1' -TenantId 'tenant-1' -Optional | Out-Null

                Should -Invoke Connect-MgGraph -Times 1 -ModuleName NSP.M365.Graph -ParameterFilter {
                    $ClientId -eq 'client-1' -and $TenantId -eq 'tenant-1'
                }
            }
        }

        It 'omits ClientId and TenantId when not provided, falling back to the SDK default app' {
            InModuleScope NSP.M365.Graph {
                Mock Connect-MgGraph { } -ModuleName NSP.M365.Graph
                function Get-MgContext { $null }

                Connect-NSPGraph -Scopes 'DeviceManagementApps.Read.All' -Connect -Optional | Out-Null

                Should -Invoke Connect-MgGraph -Times 1 -ModuleName NSP.M365.Graph -ParameterFilter {
                    $null -eq $ClientId -and $null -eq $TenantId
                }
            }
        }

        It 'omits UseDeviceCode by default, leaving Connect-MgGraph on its native broker flow' {
            InModuleScope NSP.M365.Graph {
                Mock Connect-MgGraph { } -ModuleName NSP.M365.Graph
                function Get-MgContext { $null }

                Connect-NSPGraph -Scopes 'DeviceManagementApps.Read.All' -Connect -Optional | Out-Null

                Should -Invoke Connect-MgGraph -Times 1 -ModuleName NSP.M365.Graph -ParameterFilter {
                    -not $PSBoundParameters.ContainsKey('UseDeviceCode')
                }
            }
        }

        It 'passes UseDeviceCode through when requested, to avoid the WAM broker popup' {
            InModuleScope NSP.M365.Graph {
                Mock Connect-MgGraph { } -ModuleName NSP.M365.Graph
                function Get-MgContext { $null }

                Connect-NSPGraph -Scopes 'DeviceManagementApps.Read.All' -Connect -UseDeviceCode -Optional | Out-Null

                Should -Invoke Connect-MgGraph -Times 1 -ModuleName NSP.M365.Graph -ParameterFilter {
                    $UseDeviceCode -eq $true
                }
            }
        }
    }

    Context 'avoiding a redundant reconnect' {
        It 'does not call Connect-MgGraph again when the existing session already matches client, tenant, and scopes' {
            InModuleScope NSP.M365.Graph {
                Mock Connect-MgGraph { throw 'should not be called' } -ModuleName NSP.M365.Graph
                function Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; ClientId = 'client-1'; Account = 'operator'; Scopes = @('DeviceManagementApps.ReadWrite.All') } }

                $context = Connect-NSPGraph -Scopes 'DeviceManagementApps.ReadWrite.All' -Connect -ClientId 'client-1' -TenantId 'tenant-1'

                $context.TenantId | Should -Be 'tenant-1'
                Should -Invoke Connect-MgGraph -Times 0 -ModuleName NSP.M365.Graph
            }
        }

        It 'reconnects when the existing session is authenticated as a different app' {
            InModuleScope NSP.M365.Graph {
                Mock Connect-MgGraph { } -ModuleName NSP.M365.Graph
                function Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; ClientId = 'some-other-client'; Account = 'operator'; Scopes = @('DeviceManagementApps.ReadWrite.All') } }

                Connect-NSPGraph -Scopes 'DeviceManagementApps.ReadWrite.All' -Connect -ClientId 'client-1' -TenantId 'tenant-1' -Optional | Out-Null

                Should -Invoke Connect-MgGraph -Times 1 -ModuleName NSP.M365.Graph
            }
        }

        It 'reconnects when the existing session is missing a required scope' {
            InModuleScope NSP.M365.Graph {
                Mock Connect-MgGraph { } -ModuleName NSP.M365.Graph
                function Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; ClientId = 'client-1'; Account = 'operator'; Scopes = @('Group.Read.All') } }

                Connect-NSPGraph -Scopes 'DeviceManagementApps.ReadWrite.All' -Connect -ClientId 'client-1' -TenantId 'tenant-1' -Optional | Out-Null

                Should -Invoke Connect-MgGraph -Times 1 -ModuleName NSP.M365.Graph
            }
        }
    }
}
