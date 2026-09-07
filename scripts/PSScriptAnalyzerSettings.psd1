# PSScriptAnalyzer settings for scripts/.
# Exclusions are deliberate and documented here; anything else reported
# at Warning or Error severity fails the gate (see verify-quick.ps1
# 'script analyzer' step and .github/workflows/ci.yml).
@{
    # Defence-in-depth: include the default ruleset so a future PSScriptAnalyzer
    # version that promotes a rule's default severity from 'Information' to
    # 'Warning' cannot silently introduce a new gate failure mode without an
    # explicit settings change. This file is the single source of truth
    # for what the gate accepts.
    IncludeDefaultRules = $true
    Severity            = @('Error', 'Warning')

    # Gates print human-readable step status to the console by design;
    # they emit no pipeline data, so Write-Host is the correct stream.
    # PSUseSingularNouns: helper functions return arrays but the noun
    # describes one item ("Read-VerifyStepName" returns a list of
    # step names; the noun is the most specific descriptor of a single
    # item, not the count). The two helper names in scripts/verify-helpers.ps1
    # are deliberate; documented in W9.
    ExcludeRules = @(
        'PSAvoidUsingWriteHost'
        'PSReviewUnusedParameter'
        'PSUseSingularNouns'
}
