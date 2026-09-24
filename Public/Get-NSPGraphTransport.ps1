function Get-NSPGraphTransport {
    <#
    .SYNOPSIS
        Returns the active Graph transport and, for the Fake transport, the requests it received.
    .EXAMPLE
        (Get-NSPGraphTransport).CallLog | Where-Object Method -ne 'GET'
    #>
    [CmdletBinding()]
    param()

    [pscustomobject]@{
        Name = $script:NSPGraphTransport.Name
        CallLog = @($script:NSPGraphTransport.CallLog)
    }
}
