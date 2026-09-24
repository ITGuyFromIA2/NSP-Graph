function Invoke-NSPGraphCollection {
    <#
    .SYNOPSIS
        Reads every page of a Graph collection and returns the items.
    .DESCRIPTION
        Follows @odata.nextLink until the collection is exhausted. Each page goes through
        Invoke-NSPGraphRequest, so the active transport applies.
    .EXAMPLE
        Invoke-NSPGraphCollection -Uri 'identity/conditionalAccess/namedLocations'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('v1.0', 'beta')][string]$ApiVersion = 'v1.0'
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $page = Invoke-NSPGraphRequest -Method GET -Uri $next -ApiVersion $ApiVersion
        foreach ($item in @($page['value'])) {
            if ($null -ne $item) { $items.Add($item) }
        }
        $next = $page['@odata.nextLink']
    }
    $items.ToArray()
}
