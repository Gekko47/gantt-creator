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
    ExcludeRules = @(
        'PSAvoidUsingWriteHost'

        # Known false positive: parameters consumed inside script-block
        # arguments (Invoke-Step 'name' { ... $Solution ... }) are
        # reported as unused. $Solution / $Configuration are used in
        # those blocks in verify-quick.ps1, verify.ps1, and
        # test-locked-restore.ps1.
        'PSReviewUnusedParameter'

        # PSUseDeclaredVarsMoreThanAssignments removed from global ExcludeRules.
        # Affected test files add [SuppressMessage("PSScriptAnalyzer",
        #   "PSUseDeclaredVarsMoreThanAssignments")] before the param block.
    )
}
