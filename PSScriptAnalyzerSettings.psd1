@{
    Severity      = @("Error", "Warning")
    ExcludeRules  = @(
        # The toolkit deliberately renders an interactive, colorized terminal UI.
        "PSAvoidUsingWriteHost"
        # Internal workflow helpers use established domain names such as Scopes and Rows.
        "PSUseSingularNouns"
        # Resolve, Ensure, Escape, and Sort make internal intent clearer than generic verbs.
        "PSUseApprovedVerbs"
        # Entry scripts centralize -DryRun/-WhatIf handling around interactive mutations;
        # many New-* helpers only construct in-memory display objects.
        "PSUseShouldProcessForStateChangingFunctions"
    )
}
