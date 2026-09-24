@{
    IncludeDefaultRules = $true
    Severity = @('Error', 'Warning')
    ExcludeRules = @(
        # The operator dashboard is interactive host output by design.
        'PSAvoidUsingWriteHost'

        # Scope lists, plans, and inventories are collections; plural nouns read correctly.
        'PSUseSingularNouns'
    )
}
