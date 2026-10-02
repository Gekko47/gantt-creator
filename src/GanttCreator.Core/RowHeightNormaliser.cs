namespace GanttCreator.Core;

/// <summary>
/// One managed row's row-height normalisation decision.
/// </summary>
/// <param name="RowNumber">The worksheet row number.</param>
/// <param name="IsStructuralRow">
/// Whether the row is a <c>Splitter</c> or <c>Spacer</c>. A structural row
/// follows its own token, never the managed height.
/// </param>
/// <param name="TargetHeightPt">
/// The height the row should carry, or <see langword="null"/> when it is already
/// correct and needs no write.
/// </param>
public sealed record RowHeightNormalisation(int RowNumber, bool IsStructuralRow, double? TargetHeightPt)
{
    /// <summary>Gets whether this row needs its height written.</summary>
    public bool NeedsWrite => TargetHeightPt is not null;
}

/// <summary>
/// The result of planning managed row-height normalisation (R4.7D, ADR-0026 D4).
/// </summary>
/// <param name="Rows">The per-row decisions, in worksheet row order.</param>
/// <param name="ManagedHeightPt">The height every ordinary managed row is set to.</param>
public sealed record RowHeightNormalisationPlan(
    IReadOnlyList<RowHeightNormalisation> Rows,
    double ManagedHeightPt)
{
    /// <summary>How many rows would be written, excluding no-ops.</summary>
    public int WriteCount => Rows.Count(static row => row.NeedsWrite);
}

/// <summary>
/// Plans the normalisation of managed row heights to a single value, so a
/// worksheet the user has dragged is restored rather than left to disagree with
/// the chart.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this is a plan and not an action.</b> Normalisation writes a height on
/// every managed row. A failure part-way through would leave the worksheet
/// half-normalised, which is a state neither the user nor the chart expects. The
/// plan is computed whole and the Office layer applies it.
/// </para>
/// <para>
/// <b>Why restoration is not optional.</b> A lane's height is the measured row
/// height (ADR-0026 D2), and the resolver takes the most common measured value.
/// One dragged row therefore survives measurement -- it simply loses the vote --
/// but a chart rendered against a sheet of arbitrary heights is no longer aligned
/// with the rows a user is reading. Normalising first is what makes the lane
/// measurement meaningful.
/// </para>
/// <para>
/// <b>Structural rows are exempt.</b> A <c>Splitter</c> or <c>Spacer</c> is not a
/// normal managed row; it follows its own token.
/// </para>
/// <para>
/// <b>No-op rows are excluded.</b> A row already at the target is left out, so a
/// correctly normalised sheet produces an empty write set and the adapter does
/// no work at all.
/// </para>
/// </remarks>
public static class RowHeightNormaliser
{
    /// <summary>The tolerance below which two heights count as equal, in points.</summary>
    public const double EqualityTolerancePt = 0.05;

    /// <summary>
    /// Plans normalisation for the supplied rows.
    /// </summary>
    /// <param name="rows">
    /// The rows to plan for, in any order. Each carries its row number, whether it
    /// is structural, and its currently measured height.
    /// </param>
    /// <param name="managedHeightPt">The target height for ordinary managed rows.</param>
    /// <param name="splitterHeightPt">The target height for a <c>Splitter</c> row.</param>
    /// <param name="spacerHeightPt">The target height for a <c>Spacer</c> row.</param>
    /// <returns>
    /// The plan, or a refusal. A non-positive target is refused rather than
    /// defaulted, because writing a nonsense row height would break the very
    /// alignment this restores.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="rows"/> is null.</exception>
    public static RowHeightNormalisationPlan? Plan(
        IReadOnlyList<MeasuredRowHeight> rows,
        double managedHeightPt,
        double splitterHeightPt,
        double spacerHeightPt)
    {
        ArgumentNullException.ThrowIfNull(rows);

        if (!IsPositive(managedHeightPt) || !IsPositive(splitterHeightPt) || !IsPositive(spacerHeightPt))
        {
            return null;
        }

        List<RowHeightNormalisation> planned = [];
        foreach (MeasuredRowHeight row in rows.OrderBy(static row => row.RowNumber))
        {
            // An unknown kind is refused rather than defaulted. `Kind` can arrive
            // from an out-of-range cast, and guessing "managed" for it would write a
            // height the caller never chose -- a wrong row height is exactly the
            // misalignment normalisation exists to prevent.
            var target = row.Kind switch
            {
                MeasuredRowKind.Splitter => splitterHeightPt,
                MeasuredRowKind.Spacer => spacerHeightPt,
                MeasuredRowKind.Managed => managedHeightPt,
                _ => double.NaN,
            };

            if (!IsPositive(target))
            {
                return null;
            }

            planned.Add(
                new RowHeightNormalisation(
                    row.RowNumber,
                    row.Kind is MeasuredRowKind.Splitter or MeasuredRowKind.Spacer,
                    Math.Abs(row.CurrentHeightPt - target) <= EqualityTolerancePt ? null : target));
        }

        return new RowHeightNormalisationPlan(planned, managedHeightPt);
    }

    private static bool IsPositive(double value) => double.IsFinite(value) && value > 0;
}
