namespace GanttCreator.AddIn;

/// <summary>
/// The one failure boundary for the dynamic Ribbon getters. A getter is a pure read
/// of the cached snapshot, so a probe failure must return that getter's own safe
/// value rather than propagate into Excel's getter dispatch.
/// </summary>
/// <remarks>
/// Each getter chooses its own safe value (a disabled control fails closed for a
/// group, an enabled control fails open, text fails to empty, an index fails to
/// zero), so the boundary takes the value rather than deciding it.
/// </remarks>
internal static class RibbonGetter
{
    /// <summary>
    /// Runs one getter read, returning <paramref name="onFailure"/> when it throws.
    /// </summary>
    /// <typeparam name="T">The getter's return type.</typeparam>
    /// <param name="stateService">The state service the getter reads.</param>
    /// <param name="read">The read itself.</param>
    /// <param name="onFailure">The value this getter returns when the read fails.</param>
    /// <returns>The read value, or the getter's own safe value.</returns>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="stateService"/> or <paramref name="read"/> is <see langword="null"/>.
    /// </exception>
    internal static T Read<T>(RibbonStateService stateService, Func<RibbonStateService, T> read, T onFailure)
    {
        ArgumentNullException.ThrowIfNull(stateService);
        ArgumentNullException.ThrowIfNull(read);

        // CA1031: the getter dispatch is Excel's; a probe failure must degrade to
        // this getter's safe value, never become a host error dialog.
#pragma warning disable CA1031
        try
        {
            return read(stateService);
        }
        catch
        {
            return onFailure;
        }
#pragma warning restore CA1031
    }
}
