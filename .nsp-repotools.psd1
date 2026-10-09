@{
    # Settings for Find-NSPClientReference (NSP.RepoTools). Tracked: name what is accepted, never a client.
    ClientSweep = @{
        ExcludePath = @()
        Allow = @(
            # Invented tenant in tests and fixtures.
            'fixture.onmicrosoft.com'
        )
        AllowPattern = @()
    }
}
