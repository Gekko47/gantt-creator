using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// R6.3 D2: the staged set's bounds-equality contract. Every refusal-shaped outcome has
/// a positive test that constructs the bad input (AGENTS.md validator rule), and the
/// tolerance boundary is pinned in both directions.
/// </summary>
public sealed class StagedEnvelopeCheckTests
{
    private static readonly RectD _expected = new(100, 50, 200, 120);

    [Fact]
    public void A_staged_set_that_fills_the_expected_bounds_matches()
    {
        // The four edges are touched by four shapes, and an interior shape adds nothing
        // to the union - the ordinary composition the contract exists for.
        StagedEnvelopeResult result = StagedEnvelopeCheck.Evaluate(
            [
                new RectD(100, 50, 200, 10),
                new RectD(100, 160, 200, 10),
                new RectD(100, 50, 5, 120),
                new RectD(295, 50, 5, 120),
                new RectD(150, 90, 20, 20),
            ],
            _expected);

        Assert.True(result.Matches(_expected));
        Assert.Empty(result.Strays);
        Assert.Equal(_expected, result.Union);
    }

    [Fact]
    public void A_shape_outside_the_expected_bounds_is_reported_as_a_stray_and_fails()
    {
        // The positive stray test the work item requires: one primitive escaped the
        // section 1 union, and the check must NAME it rather than pad the bounds.
        RectD stray = new(320, 60, 10, 10);
        StagedEnvelopeResult result = StagedEnvelopeCheck.Evaluate(
            [new RectD(100, 50, 200, 120), stray],
            _expected);

        Assert.False(result.Matches(_expected));
        RectD reported = Assert.Single(result.Strays);
        Assert.Equal(stray, reported);
    }

    [Fact]
    public void A_union_short_of_an_edge_fails_even_with_no_stray()
    {
        // Nothing lies outside the bounds, but the set does not reproduce them: the
        // composition lost a primitive (or the frame never closed), and the contract
        // is equality, not containment.
        StagedEnvelopeResult result = StagedEnvelopeCheck.Evaluate(
            [new RectD(100, 50, 200, 100)],
            _expected);

        Assert.Empty(result.Strays);
        Assert.False(result.Matches(_expected));
    }

    [Fact]
    public void An_empty_staged_set_has_no_union_and_fails()
    {
        StagedEnvelopeResult result = StagedEnvelopeCheck.Evaluate([], _expected);

        Assert.Null(result.Union);
        Assert.Empty(result.Strays);
        Assert.False(result.Matches(_expected));
    }

    [Fact]
    public void A_shape_exactly_at_the_tolerance_boundary_is_inside()
    {
        // The tolerance is inclusivity: an edge within GeometryMath.Epsilon of the
        // expected edge is in bounds, and one just beyond it is a stray. Both
        // directions, so the comparison cannot be loosened or tightened silently.
        RectD inside = new(
            _expected.Left - GeometryMath.Epsilon,
            _expected.Top,
            _expected.Width + GeometryMath.Epsilon,
            _expected.Height);
        RectD outside = new(
            _expected.Left - (GeometryMath.Epsilon * 2),
            _expected.Top,
            _expected.Width + (GeometryMath.Epsilon * 2),
            _expected.Height);

        Assert.Empty(StagedEnvelopeCheck.Evaluate([inside], _expected).Strays);
        _ = Assert.Single(StagedEnvelopeCheck.Evaluate([outside], _expected).Strays);
    }

    [Fact]
    public void A_host_rounded_envelope_matches_within_the_host_tolerance()
    {
        // R6.4: the group envelope is a host read-back stored in single precision. The
        // comparison passes the host tolerance explicitly rather than widening the scene
        // tolerance (measured 2026-11-18: 514.3622047pt requested, 514.3621826pt read back).
        RectD rounded = new(
            _expected.Left + 0.002,
            _expected.Top - 0.002,
            _expected.Width - 0.004,
            _expected.Height + 0.004);

        StagedEnvelopeResult result = StagedEnvelopeCheck.Evaluate(
            [rounded],
            _expected,
            HostOwnedGeometry.GeometryMatchTolerancePt);

        Assert.Empty(result.Strays);
        Assert.True(result.Matches(_expected));
        Assert.Equal(HostOwnedGeometry.GeometryMatchTolerancePt, result.TolerancePt);
    }

    [Fact]
    public void An_envelope_beyond_the_host_tolerance_is_still_a_stray()
    {
        // The host tolerance is not a blanket: twice it is a real disagreement, and the
        // stray list names the envelope rather than only failing the union comparison.
        RectD beyond = new(
            _expected.Left - (HostOwnedGeometry.GeometryMatchTolerancePt * 2),
            _expected.Top,
            _expected.Width + (HostOwnedGeometry.GeometryMatchTolerancePt * 2),
            _expected.Height);

        StagedEnvelopeResult result = StagedEnvelopeCheck.Evaluate(
            [beyond],
            _expected,
            HostOwnedGeometry.GeometryMatchTolerancePt);

        _ = Assert.Single(result.Strays);
        Assert.False(result.Matches(_expected));
    }

    [Fact]
    public void The_default_tolerance_remains_the_scene_epsilon()
    {
        // Every R6.3 call site keeps the exact comparison: an envelope a thousandth of a
        // point off does not match under the default.
        RectD off = new(
            _expected.Left + 0.001,
            _expected.Top,
            _expected.Width - 0.001,
            _expected.Height);

        StagedEnvelopeResult result = StagedEnvelopeCheck.Evaluate([off], _expected);

        Assert.False(result.Matches(_expected));
        Assert.Equal(GeometryMath.Epsilon, result.TolerancePt);
    }

    [Fact]
    public void A_null_envelope_list_is_rejected()
    {
        _ = Assert.Throws<ArgumentNullException>(
            () => StagedEnvelopeCheck.Evaluate(null!, _expected));
    }
}
