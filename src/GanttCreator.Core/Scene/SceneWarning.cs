namespace GanttCreator.Core.Scene;

/// <summary>A non-exported scene warning associated with one source row.</summary>
public sealed record SceneWarning
{
    /// <summary>Initialises a validated scene warning.</summary>
    /// <param name="ownerId">The stable source row identifier.</param>
    /// <param name="code">The stable warning code.</param>
    /// <param name="message">The deterministic warning message.</param>
    public SceneWarning(GanttRowId ownerId, string code, string message)
    {
        ArgumentNullException.ThrowIfNull(ownerId);
        ArgumentException.ThrowIfNullOrWhiteSpace(code);
        ArgumentException.ThrowIfNullOrWhiteSpace(message);
        OwnerId = ownerId;
        Code = code;
        Message = message;
    }

    /// <summary>Gets the stable source row identifier.</summary>
    public GanttRowId OwnerId { get; }

    /// <summary>Gets the stable warning code.</summary>
    public string Code { get; }

    /// <summary>Gets the deterministic warning message.</summary>
    public string Message { get; }
}
