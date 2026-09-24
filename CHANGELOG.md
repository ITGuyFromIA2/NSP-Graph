# Changelog

## 0.1.0 (unreleased)

- `Connect-NSPGraph` extracted from NSP-IntuneApps with its tests.
- `Invoke-NSPGraphRequest` / `Invoke-NSPGraphCollection` with a swappable transport
  (`MgGraphRequest`, `Fake`).
- `Register-NSPGraphAppRegistration`: generalized, REST-based tenant bootstrap for a tool's own
  delegated app registration, with plan/repair/adopt behavior, tested against an in-memory fake tenant.
