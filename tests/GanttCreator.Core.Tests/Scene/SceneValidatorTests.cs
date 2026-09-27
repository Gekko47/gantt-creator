using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for the R3.12 scene-invariant inspector. Every finding the validator can
/// emit has a positive test here that constructs the broken candidate and asserts
/// the finding fires, so a check cannot silently stop working.
/// </summary>
/// <remarks>
/// The invariants the validator does <em>not</em> own are unreachable by
/// construction — <c>RectD</c> throws on a non-finite or negative extent,
/// <c>ScenePolygon</c> throws on fewer than three points, and
/// <c>GanttScene.TryCreate</c> refuses duplicate IDs and unresolved children — so
/// they are proven by <c>RectDTests</c> and <c>GanttSceneTests</c> instead.
/// </remarks>
public sealed class SceneValidatorTests
{
    private static readonly RectD _chartBounds = new(0, 0, 500, 300);
    private static readonly SceneOwnerId _owner = SceneOwnerId.ForRow(GanttRowId.New());
    private static readonly SceneStyle _style = new("Test");

    private static SceneLine Line(string id, PointD from, PointD to) =>
        new(id, _owner, ZLayer.Grid, from, to, _style);

    [Fact]
    public void A_clean_candidate_reports_no_findings()
    {
        SceneValidationReport report = SceneValidator.Validate(
            [
                new SceneRect(ScenePrimitive.CreateId(_owner, "bar"), _owner, ZLayer.ActivityBody, new RectD(10, 10, 40, 8), _style),
                Line(ScenePrimitive.CreateId(_owner, "grid"), new PointD(0, 0), new PointD(500, 0)),
                new SceneText(ScenePrimitive.CreateId(_owner, "label"), _owner, ZLayer.Label, "Inside", new RectD(12, 11, 20, 8), _style, GanttTextAlignment.Left),
            ],
            _chartBounds);

        Assert.True(report.IsClean, "Findings: " + string.Join("; ", report.Findings));
    }

    [Fact]
    public void A_null_primitive_list_is_reported_as_clean_rather_than_throwing()
    {
        SceneValidationReport report = SceneValidator.Validate(null, _chartBounds);

        Assert.True(report.IsClean);
    }

    [Fact]
    public void A_non_finite_line_endpoint_is_detected()
    {
        // PointD performs no validation, so this is the one geometry that can carry
        // NaN into a constructed primitive. Both a NaN and an Infinity are covered
        // because they reach the check by different arithmetic.
        SceneValidationReport nan = SceneValidator.Validate(
            [Line(ScenePrimitive.CreateId(_owner, "grid"), new PointD(double.NaN, 0), new PointD(10, 0))],
            _chartBounds);
        SceneValidationReport infinity = SceneValidator.Validate(
            [Line(ScenePrimitive.CreateId(_owner, "grid"), new PointD(0, 0), new PointD(double.PositiveInfinity, 0))],
            _chartBounds);

        Assert.Equal(SceneValidationReport.NonFiniteLineEndpointCode, Assert.Single(nan.Findings).Code);
        Assert.Equal(SceneValidationReport.NonFiniteLineEndpointCode, Assert.Single(infinity.Findings).Code);
    }

    [Fact]
    public void A_label_fully_outside_the_chart_bounds_is_detected()
    {
        // The label sits entirely to the right of the chart's right edge.
        SceneValidationReport report = SceneValidator.Validate(
            [new SceneText(ScenePrimitive.CreateId(_owner, "label"), _owner, ZLayer.Label, "Outside", new RectD(600, 10, 40, 8), _style, GanttTextAlignment.Left)],
            _chartBounds);

        Assert.Equal(SceneValidationReport.LabelOutsideChartCode, Assert.Single(report.Findings).Code);
    }

    [Fact]
    public void A_label_straddling_the_chart_edge_is_detected()
    {
        // A label that only partly fits is not contained, so it is a finding; the
        // containment rule is deliberately not tolerant of a partial overlap.
        SceneValidationReport report = SceneValidator.Validate(
            [new SceneText(ScenePrimitive.CreateId(_owner, "label"), _owner, ZLayer.Label, "Straddle", new RectD(480, 10, 40, 8), _style, GanttTextAlignment.Left)],
            _chartBounds);

        Assert.Equal(SceneValidationReport.LabelOutsideChartCode, Assert.Single(report.Findings).Code);
    }

    [Fact]
    public void A_repeated_warning_for_one_owner_and_code_is_detected()
    {
        SceneWarning warning = new(_owner, "AmbiguousStackOverlap", "Two events share a slot.");

        SceneValidationReport report = SceneValidator.Validate(
            [],
            _chartBounds,
            [warning, new SceneWarning(_owner, "AmbiguousStackOverlap", "A different wording of the same condition.")]);

        SceneValidationFinding finding = Assert.Single(report.Findings);
        Assert.Equal(SceneValidationReport.DuplicateWarningCode, finding.Code);
    }

    [Fact]
    public void The_same_warning_code_for_two_different_owners_is_not_a_duplicate()
    {
        // Warnings are keyed by owner and code together: the same condition on two
        // rows is two real findings, not one repeated one.
        SceneOwnerId other = SceneOwnerId.ForRow(GanttRowId.New());
        SceneWarning first = new(_owner, "AmbiguousStackOverlap", "Row A.");
        SceneWarning second = new(other, "AmbiguousStackOverlap", "Row B.");

        SceneValidationReport report = SceneValidator.Validate([], _chartBounds, [first, second]);

        Assert.True(report.IsClean, "Findings: " + string.Join("; ", report.Findings));
    }

    [Fact]
    public void A_malformed_primitive_id_is_detected()
    {
        // ScenePrimitive only rejects a blank ID, so a role-less ID reaches the
        // validator and must be reported rather than assumed well formed.
        SceneValidationReport report = SceneValidator.Validate(
            [Line("row-without-a-role", new PointD(0, 0), new PointD(10, 0))],
            _chartBounds);

        Assert.Equal(SceneValidationReport.MalformedPrimitiveIdCode, Assert.Single(report.Findings).Code);
    }

    [Fact]
    public void A_role_derived_id_is_not_reported_as_malformed()
    {
        SceneValidationReport report = SceneValidator.Validate(
            [Line(ScenePrimitive.CreateId(_owner, "grid"), new PointD(0, 0), new PointD(10, 0))],
            _chartBounds);

        Assert.True(report.IsClean);
    }

    [Fact]
    public void Findings_are_reported_in_deterministic_discovery_order()
    {
        // Two different broken primitives, validated in both input orders, must
        // report the same findings in the same sequence.
        ScenePrimitive badId = Line("no-role-here", new PointD(0, 0), new PointD(10, 0));
        ScenePrimitive badGeometry = Line(
            ScenePrimitive.CreateId(_owner, "grid"),
            new PointD(double.NaN, 0),
            new PointD(10, 0));

        string[] first = Codes(SceneValidator.Validate([badId, badGeometry], _chartBounds));
        string[] second = Codes(SceneValidator.Validate([badGeometry, badId], _chartBounds));

        Assert.Equal(first, second);
        Assert.Equal(SceneValidationReport.MalformedPrimitiveIdCode, first[0]);
        Assert.Equal(SceneValidationReport.NonFiniteLineEndpointCode, first[1]);
    }

    [Fact]
    public void A_built_scene_from_the_orchestrator_reports_no_findings()
    {
        // The end-to-end proof: the scene R3.12 exists to certify is invariant-clean
        // under the validator that inspects it.
        SceneValidationReport report = SceneValidator.Validate(SceneBuilderTests.BuildScene());

        Assert.True(report.IsClean, "Findings: " + string.Join("; ", report.Findings));
    }

    private static string[] Codes(SceneValidationReport report) =>
        [.. report.Findings.Select(finding => finding.Code)];
}
