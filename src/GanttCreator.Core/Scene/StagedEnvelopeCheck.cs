namespace GanttCreator.Core.Scene;

/// <summary>The envelope a staged shape set occupies, plus every shape outside it.</summary>
/// <param name="Union">
/// The bounding box of every supplied envelope, or <see langword="null"/> when no
/// envelope was supplied.
/// </param>
/// <param name="Strays">
/// The envelopes that lie outside the expected bounds, in input order; empty when every
/// envelope is inside. A stray is reported by value so a failing live gate can name the
/// shape that escaped rather than only reporting that the union disagreed.
/// </param>
/// <remarks>
/// <b>Diagnostics, not a pass/fail shortcut.</b> A shape outside the expected bounds
/// always moves the union outside them as well, so <see cref="Matches"/> does not need
/// the stray list to be correct — but a bare boolean cannot say <em>which</em> primitive
/// escaped, and the R6.3 stop condition requires finding it rather than padding the
/// bounds.
/// </remarks>
public sealed record StagedEnvelopeResult(RectD? Union, IReadOnlyList<RectD> Strays)
{
    /// <summary>
    /// Gets whether the union equals the expected bounds within
    /// <see cref="GeometryMath.Epsilon"/> on all four edges, with no stray shape.
    /// </summary>
    /// <param name="expected">The bounds the staged set must occupy.</param>
    /// <returns><see langword="true"/> when the staged set reproduces the expected bounds.</returns>
    public bool Matches(RectD expected) =>
        Union is { } union
        && Strays.Count == 0
        && GeometryMath.ApproximatelyEqual(union.Left, expected.Left)
        && GeometryMath.ApproximatelyEqual(union.Top, expected.Top)
        && GeometryMath.ApproximatelyEqual(union.Right, expected.Right)
        && GeometryMath.ApproximatelyEqual(union.Bottom, expected.Bottom);
}

/// <summary>
/// R6.3 D2: the staged composition's bounds-equality contract.
/// </summary>
/// <remarks>
/// <para>
/// The staged set's overall bounds must equal the R6.1 export bounds within
/// <see cref="GeometryMath.Epsilon"/>, and entity guide section 1 guarantees no
/// whitespace outside <see cref="GanttScene.ChartBounds"/> is exported — so the union of
/// the staged envelopes is expected to reproduce the translated export bounds exactly.
/// A mismatch in either direction (a shape outside, or a union short of an edge) fails
/// the contract; neither is repaired here.
/// </para>
/// <para>
/// Pure over <see cref="RectD"/> values so it runs in the Core test harness without an
/// Office fake, and so R6.4's group-bounds pin can consume the same comparison rather
/// than stating a second notion of equality.
/// </para>
/// </remarks>
public static class StagedEnvelopeCheck
{
    /// <summary>Evaluates a staged set's envelopes against the bounds they must occupy.</summary>
    /// <param name="envelopes">One bounding box per staged shape, in any order.</param>
    /// <param name="expected">The bounds the staged set must occupy.</param>
    /// <returns>The union, the stray envelopes, and the equality answer.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="envelopes"/> is <see langword="null"/>.</exception>
    public static StagedEnvelopeResult Evaluate(IReadOnlyList<RectD> envelopes, RectD expected)
    {
        ArgumentNullException.ThrowIfNull(envelopes);

        if (envelopes.Count == 0)
        {
            return new StagedEnvelopeResult(null, []);
        }

        List<RectD> strays = [];
        double left = double.PositiveInfinity;
        double top = double.PositiveInfinity;
        double right = double.NegativeInfinity;
        double bottom = double.NegativeInfinity;

        foreach (RectD envelope in envelopes)
        {
            if (envelope.Left < expected.Left - GeometryMath.Epsilon
                || envelope.Top < expected.Top - GeometryMath.Epsilon
                || envelope.Right > expected.Right + GeometryMath.Epsilon
                || envelope.Bottom > expected.Bottom + GeometryMath.Epsilon)
            {
                strays.Add(envelope);
            }

            left = Math.Min(left, envelope.Left);
            top = Math.Min(top, envelope.Top);
            right = Math.Max(right, envelope.Right);
            bottom = Math.Max(bottom, envelope.Bottom);
        }

        // A union is derived from the same envelopes the scene guaranteed are inside
        // the export bounds, so the width/height are non-negative by construction; the
        // RectD constructor still validates finiteness so a non-finite input cannot
        // masquerade as an equal envelope.
        return new StagedEnvelopeResult(
            new RectD(left, top, right - left, bottom - top),
            strays);
    }
}
