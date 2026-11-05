namespace GanttCreator.Office;

/// <summary>
/// Records which owned shapes the user has deliberately unlocked, so the automatic refresh can be
/// suspended while a manual edit is in progress (ADR-0055 clause 6).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why the scheduler is suspended rather than the shapes being exempted.</b> Exempting an unlocked
/// shape from reconciliation would make "every owned shape equals the scene" stop being an invariant:
/// the reconciler would grow a special case, and the next explicit Refresh would have to either honour
/// the exemption (leaving the chart permanently inconsistent with the table) or discard it (making the
/// exemption a lie). Suspending the scheduler keeps the invariant total, costs one set, and gives the
/// user a predictable rule: an explicit Lock or Refresh puts everything back.
/// </para>
/// <para>
/// Mutable and explicitly shared, like <see cref="ProtectionArmingState"/>: it lives as long as the
/// Excel session and is touched only on the main thread.
/// </para>
/// </remarks>
public sealed class ShapeUnlockState
{
    private readonly HashSet<string> _unlocked = new(StringComparer.Ordinal);

    /// <summary>The Excel session's single unlock set.</summary>
    public static ShapeUnlockState Shared { get; } = new();

    /// <summary>Gets how many owned shapes are currently unlocked.</summary>
    public int Count => _unlocked.Count;

    /// <summary>Gets whether any owned shape is unlocked.</summary>
    /// <remarks>
    /// The scheduler reads this before queueing an automatic refresh: a value of <see langword="true"/>
    /// means a manual edit may be in flight, so an unrelated cell change must not rebuild the chart
    /// underneath it.
    /// </remarks>
    public bool IsAnyUnlocked => _unlocked.Count > 0;

    /// <summary>Records that an owned shape was unlocked.</summary>
    /// <param name="primitiveId">The scene primitive id of the shape.</param>
    /// <returns><see langword="true"/> when the id was not already recorded.</returns>
    public bool Add(string primitiveId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);
        return _unlocked.Add(primitiveId);
    }

    /// <summary>Gets whether the identifier is currently unlocked.</summary>
    /// <param name="primitiveId">The scene primitive id of the shape.</param>
    /// <returns><see langword="true"/> when it was recorded by an unlock.</returns>
    public bool Contains(string primitiveId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);
        return _unlocked.Contains(primitiveId);
    }

    /// <summary>
    /// Forgets every unlock, which is what a Lock command or an explicit Refresh does before it
    /// re-locks and re-protects.
    /// </summary>
    /// <returns><see langword="true"/> when anything was cleared.</returns>
    public bool Clear()
    {
        if (_unlocked.Count == 0)
        {
            return false;
        }

        _unlocked.Clear();
        return true;
    }
}
