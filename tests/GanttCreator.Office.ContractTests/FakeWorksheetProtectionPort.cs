namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Models the protection port for the refresh's arming step: it holds the two flags the host
/// reports, records every call in order, and can be made to refuse either write.
/// </summary>
/// <remarks>
/// The state is modelled rather than asserted through the calls alone, because the point of the
/// arming rule is the ORDER — lock, then read, then release, then protect — and a port that only
/// counted calls could not tell an unprotect-before-protect from a protect-after-protect.
/// </remarks>
internal sealed class FakeWorksheetProtectionPort : IWorksheetProtectionPort
{
    /// <summary>The state <see cref="Read"/> reports. A fresh sheet reports neither flag.</summary>
    public WorksheetProtectionState? State { get; set; } = new(false, false);

    /// <summary>The refusal <see cref="Protect"/> returns, or <see langword="null"/> to succeed.</summary>
    public ProtectionWriteRefusal? ProtectRefusal { get; set; }

    /// <summary>The refusal <see cref="Unprotect"/> returns, or <see langword="null"/> to succeed.</summary>
    public ProtectionWriteRefusal? UnprotectRefusal { get; set; }

    /// <summary>Raised inside <see cref="Protect"/>, so a test can observe what had already run.</summary>
    public Action? OnProtect { get; set; }

    /// <summary>The calls this refresh made, in order.</summary>
    public List<string> Calls { get; } = [];

    /// <summary>How many times the sheet was protected.</summary>
    public int ProtectCount => Calls.Count(call => call == "Protect");

    /// <summary>How many times the sheet was unprotected.</summary>
    public int UnprotectCount => Calls.Count(call => call == "Unprotect");

    /// <inheritdoc />
    public WorksheetProtectionState? Read()
    {
        Calls.Add("Read");
        return State;
    }

    /// <inheritdoc />
    public ProtectionWriteOutcome Protect()
    {
        Calls.Add("Protect");
        OnProtect?.Invoke();
        if (ProtectRefusal is { } refusal)
        {
            return ProtectionWriteOutcome.Refused(refusal);
        }

        State = new WorksheetProtectionState(true, true);
        return ProtectionWriteOutcome.Ok();
    }

    /// <inheritdoc />
    public ProtectionWriteOutcome Unprotect()
    {
        Calls.Add("Unprotect");
        if (UnprotectRefusal is { } refusal)
        {
            return ProtectionWriteOutcome.Refused(refusal);
        }

        State = new WorksheetProtectionState(false, false);
        return ProtectionWriteOutcome.Ok();
    }
}
