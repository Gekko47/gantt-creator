namespace GanttCreator.Office.ContractTests;

/// <summary>R6.6: a copy port that records the requested name and can refuse.</summary>
internal sealed class FakeShapeCopyPort : IShapeCopyPort
{
    /// <summary>Gets the group names the stage asked to copy, in call order.</summary>
    public List<string> RequestedNames { get; } = [];

    /// <summary>Gets the refusal the fake reports, or <see langword="null"/> for success.</summary>
    public ShapeWriteRefusal? Refusal { get; init; }

    /// <summary>Gets a hook the test uses to record the operation sequence.</summary>
    public Action? AfterCopy { get; init; }

    /// <inheritdoc />
    public ShapeWriteOutcome CopyStaged(string groupName)
    {
        RequestedNames.Add(groupName);
        AfterCopy?.Invoke();
        return Refusal is { } refusal
            ? ShapeWriteOutcome.Refused(refusal)
            : ShapeWriteOutcome.Ok();
    }
}

/// <summary>R6.6: a clipboard reader that returns a scripted observation.</summary>
internal sealed class FakeClipboardFormatReader : IClipboardFormatReader
{
    /// <summary>Gets the number of reads the stage performed.</summary>
    public int ReadCalls { get; private set; }

    /// <summary>Gets the observation the fake returns.</summary>
    public ClipboardFormatRead Result { get; init; } =
        ClipboardFormatRead.ForSuccess(new ClipboardFormatSnapshot(1, []));

    /// <summary>Gets a hook the test uses to record the operation sequence.</summary>
    public Action? AfterRead { get; init; }

    /// <inheritdoc />
    public ClipboardFormatRead Read()
    {
        ReadCalls++;
        AfterRead?.Invoke();
        return Result;
    }

    /// <summary>Builds a successful snapshot with the given format names.</summary>
    /// <param name="names">The observed clipboard format names.</param>
    /// <returns>The successful read.</returns>
    public static ClipboardFormatRead Snapshot(params string[] names) =>
        ClipboardFormatRead.ForSuccess(new ClipboardFormatSnapshot(548, names));
}
