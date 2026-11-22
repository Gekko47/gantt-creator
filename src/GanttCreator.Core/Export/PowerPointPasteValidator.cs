namespace GanttCreator.Core.Export;

/// <summary>
/// The closed set of typed reasons a pasted <c>ShapeRange</c> is rejected (R7.3 D1/D5).
/// </summary>
/// <remarks>
/// <para>
/// Each member is one user-visible refusal the command translates to one actionable message;
/// the message text belongs to the command, never to this enum. Every member has a positive
/// test that constructs the bad facts (or the bad expectation) it names.
/// </para>
/// <para>
/// There is deliberately no separate aspect-mismatch member: if both axes match within the
/// tolerance the aspect cannot mismatch, so a separate arm would be unreachable dead code.
/// The measured aspect delta is recorded on <see cref="PowerPointPasteValidation"/> instead.
/// </para>
/// </remarks>
public enum PowerPointPasteRefusal
{
    /// <summary>The host returned no shape range at all.</summary>
    NullRange = 0,

    /// <summary>The returned range contains no shapes.</summary>
    EmptyRange = 1,

    /// <summary>The returned count is not <c>1 + expectedChildCount</c>.</summary>
    CountMismatch = 2,

    /// <summary>The returned range contains no <c>msoGroup</c> shape.</summary>
    NoGroup = 3,

    /// <summary>The returned range contains more than one <c>msoGroup</c> shape.</summary>
    MultipleGroups = 4,

    /// <summary>The located group is not an <c>msoGroup</c> shape.</summary>
    WrongTypeGroup = 5,

    /// <summary>The returned bounds are not finite numbers.</summary>
    NonFiniteBounds = 6,

    /// <summary>The returned bounds are not strictly positive.</summary>
    NonPositiveBounds = 7,

    /// <summary>The returned bounds differ from the expectation by more than the tolerance.</summary>
    BoundsMismatch = 8,

    /// <summary>The expectation's child count is not strictly positive — a caller error.</summary>
    InvalidExpectedChildCount = 9,

    /// <summary>The expectation's bounds are not finite and strictly positive — a caller error.</summary>
    InvalidExpectedBounds = 10,
}

/// <summary>
/// The typed outcome of validating one pasted <c>ShapeRange</c> (R7.3).
/// </summary>
/// <param name="Succeeded">Whether the range satisfied every D1 rule.</param>
/// <param name="Refusal">The typed refusal when <paramref name="Succeeded"/> is <see langword="false"/>; otherwise <see langword="null"/>.</param>
/// <param name="ValidatedFacts">
/// The validated facts forwarded on success (D2), or <see langword="null"/> on a refusal.
/// </param>
/// <param name="MeasuredBoundsDeltaPt">
/// The larger of the two per-axis bounds differences in points. Recorded on success and on
/// <see cref="PowerPointPasteRefusal.BoundsMismatch"/>; zero for every other refusal.
/// </param>
/// <param name="MeasuredAspectDelta">
/// The difference between the measured and expected width/height ratios. Recorded on success
/// and on <see cref="PowerPointPasteRefusal.BoundsMismatch"/>; zero for every other refusal.
/// </param>
public sealed record PowerPointPasteValidation(
    bool Succeeded,
    PowerPointPasteRefusal? Refusal,
    PowerPointPasteFacts? ValidatedFacts,
    double MeasuredBoundsDeltaPt,
    double MeasuredAspectDelta);

/// <summary>
/// R7.3 D1: validates one pasted <c>ShapeRange</c> against the staged expectation.
/// </summary>
/// <remarks>
/// <para>
/// <b>The group is located by type, never by index (ADR-0066 D3).</b> The Office port walks
/// the returned range and reports how many <c>msoGroup</c> shapes it found and the located
/// group's type; the validator requires exactly one and requires it to be the documented
/// <see cref="MsoGroupShapeType"/>. A single-item expectation is wrong: the measured result
/// is the group plus its flattened children, so the expected count is <c>1 + N</c>.
/// </para>
/// <para>
/// <b>Precedence is total and deterministic.</b> The expectation is validated first — a bad
/// expectation is a caller error no host result can satisfy — then the facts in this order:
/// null range, empty range, count mismatch, no group, multiple groups, wrong-type group,
/// non-finite bounds, non-positive bounds, bounds mismatch. Impossible fact combinations
/// resolve by that order.
/// </para>
/// <para>
/// <b>Non-throwing is not success (D3).</b> A returned range that fails any rule is a typed
/// failure even though no COM call raised.
/// </para>
/// <para>
/// <b>Tolerance is the repository's (D5).</b> Bounds are compared per axis against
/// <see cref="PowerPointPasteExpectation.TolerancePt"/> (default
/// <c>HostOwnedGeometry.GeometryMatchTolerancePt</c>, 0.01 pt), and the measured bounds and
/// aspect deltas are recorded on the result rather than hidden.
/// </para>
/// </remarks>
public static class PowerPointPasteValidator
{
    /// <summary>
    /// The documented <c>PpShapeType.msoGroup</c> value (6). Core cannot name the interop enum,
    /// so the value is a documented constant pinned by a Core test against the documented
    /// enumeration.
    /// </summary>
    public const int MsoGroupShapeType = 6;

    /// <summary>
    /// Validates one pasted range's facts against the staged expectation.
    /// </summary>
    /// <param name="facts">The facts the Office port read, or <see langword="null"/> when the host returned no usable range.</param>
    /// <param name="expectation">The staged expectation the result must satisfy.</param>
    /// <returns>The typed validation outcome; a refusal is a result, never an exception.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="expectation"/> is <see langword="null"/>.</exception>
    public static PowerPointPasteValidation Validate(
        PowerPointPasteFacts? facts,
        PowerPointPasteExpectation expectation)
    {
        ArgumentNullException.ThrowIfNull(expectation);

        // The expectation is the caller's contract: a bad one fails before any host fact is
        // consulted, so the same bad expectation cannot produce two different refusals
        // depending on what the host happened to return.
        if (expectation.ExpectedChildCount <= 0)
        {
            return Refuse(PowerPointPasteRefusal.InvalidExpectedChildCount);
        }

        if (!IsFiniteAndPositive(expectation.ExpectedWidthPt)
            || !IsFiniteAndPositive(expectation.ExpectedHeightPt))
        {
            return Refuse(PowerPointPasteRefusal.InvalidExpectedBounds);
        }

        if (facts is null)
        {
            return Refuse(PowerPointPasteRefusal.NullRange);
        }

        if (facts.Count <= 0)
        {
            return Refuse(PowerPointPasteRefusal.EmptyRange);
        }

        if (facts.Count != expectation.ExpectedChildCount + 1)
        {
            return Refuse(PowerPointPasteRefusal.CountMismatch);
        }

        if (facts.GroupCount == 0)
        {
            return Refuse(PowerPointPasteRefusal.NoGroup);
        }

        if (facts.GroupCount > 1)
        {
            return Refuse(PowerPointPasteRefusal.MultipleGroups);
        }

        if (facts.GroupShapeType != MsoGroupShapeType)
        {
            return Refuse(PowerPointPasteRefusal.WrongTypeGroup);
        }

        if (!double.IsFinite(facts.WidthPt) || !double.IsFinite(facts.HeightPt))
        {
            return Refuse(PowerPointPasteRefusal.NonFiniteBounds);
        }

        if (facts.WidthPt <= 0 || facts.HeightPt <= 0)
        {
            return Refuse(PowerPointPasteRefusal.NonPositiveBounds);
        }

        double boundsDelta = Math.Max(
            Math.Abs(facts.WidthPt - expectation.ExpectedWidthPt),
            Math.Abs(facts.HeightPt - expectation.ExpectedHeightPt));
        double expectedAspect = expectation.ExpectedWidthPt / expectation.ExpectedHeightPt;
        double measuredAspect = facts.WidthPt / facts.HeightPt;
        double aspectDelta = Math.Abs(measuredAspect - expectedAspect);

        return boundsDelta > expectation.TolerancePt
            ? new PowerPointPasteValidation(
                Succeeded: false,
                Refusal: PowerPointPasteRefusal.BoundsMismatch,
                ValidatedFacts: null,
                MeasuredBoundsDeltaPt: boundsDelta,
                MeasuredAspectDelta: aspectDelta)
            : new PowerPointPasteValidation(
                Succeeded: true,
                Refusal: null,
                ValidatedFacts: facts,
                MeasuredBoundsDeltaPt: boundsDelta,
                MeasuredAspectDelta: aspectDelta);
    }

    private static PowerPointPasteValidation Refuse(PowerPointPasteRefusal refusal) =>
        new(
            Succeeded: false,
            Refusal: refusal,
            ValidatedFacts: null,
            MeasuredBoundsDeltaPt: 0,
            MeasuredAspectDelta: 0);

    private static bool IsFiniteAndPositive(double value) => double.IsFinite(value) && value > 0;
}
