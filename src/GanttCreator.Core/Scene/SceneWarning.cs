namespace GanttCreator.Core.Scene;

/// <summary>A non-exported scene warning associated with a row or the chart owner.</summary>
public sealed record SceneWarning
{
    /// <summary>Initialises a validated scene warning.</summary>
    /// <param name="ownerId">The stable row or chart scene owner.</param>
    /// <param name="code">The stable warning code.</param>
    /// <param name="message">The deterministic warning message.</param>
    public SceneWarning(SceneOwnerId ownerId, string code, string message)
    {
        ArgumentNullException.ThrowIfNull(ownerId);
        ArgumentException.ThrowIfNullOrWhiteSpace(code);
        ArgumentException.ThrowIfNullOrWhiteSpace(message);
        OwnerId = ownerId;
        Code = code;
        Message = message;
    }

    /// <summary>Gets the stable scene owner identifier.</summary>
    public SceneOwnerId OwnerId { get; }

    /// <summary>Gets the stable warning code.</summary>
    public string Code { get; }

    /// <summary>Gets the deterministic warning message.</summary>
    public string Message { get; }
}
