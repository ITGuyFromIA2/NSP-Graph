# One function per file. The manifest's FunctionsToExport is the exported surface.
foreach ($folder in 'Private', 'Public') {
    $path = Join-Path $PSScriptRoot $folder
    if (-not (Test-Path -LiteralPath $path)) { continue }
    foreach ($file in Get-ChildItem -LiteralPath $path -Filter '*.ps1' -File | Sort-Object Name) {
        try { . $file.FullName }
        catch { throw "NSP.M365.Graph: failed to load $($file.Name): $_" }
    }
}

# Session transport. MgGraphRequest rides the single Connect-MgGraph context; Fake answers from a
# caller-supplied responder so consuming modules can test planners and executors offline.
$script:NSPGraphTransport = @{
    Name = 'MgGraphRequest'
    Responder = $null
    CallLog = [System.Collections.Generic.List[object]]::new()
}
