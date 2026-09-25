# CLAUDE.md

## What this is

`NSP.M365.Graph`: shared Graph connection, paging, and request transport for the NSP M365 tools.
Toolkit B4 in `C:\GitRepo\PoSHRepo\ClaudeStuff\04_Toolkit_Candidates.md`. The first consumer is
`NSP-PoSHToolkits\NSP-ConditionalAccess` (see its PLAN.md, "Graph layer").

## Hard rules

- Windows PowerShell 5.1 is the floor. The module must import on a bare host; Graph SDK modules are not
  `RequiredModules`.
- Nothing domain-specific lives here. No Conditional Access, Intune, or group logic, and no scope lists.
- The canonical data shape is Graph REST JSON as hashtables (camelCase).
- Never abbreviate Conditional Access as "CA" (NSP-CAManager owns that noun space).
- One function per file. `FunctionsToExport` matches `Public\`. Exported functions need `.SYNOPSIS` and `.EXAMPLE`.

## Testing

Pester 5+. Run `.\tools\Test-Repo.ps1`.
