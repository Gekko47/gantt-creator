using GanttCreator.Core.Export;

namespace GanttCreator.Core.Tests.Export;

/// <summary>
/// R7.3: the rejection matrix for one pasted <c>ShapeRange</c>. Every refusal the validator
/// can produce has a positive test that constructs the bad facts (or the bad expectation),
/// the precedence between overlapping conditions is pinned, and the success pin forwards the
/// validated facts.
/// </summary>
/// <remarks>
/// The healthy case starts from the ADR-0066 measured form: a range of <c>1 + 4</c> — the
/// group plus its four flattened children — at 340 x 190 pt. <see cref="PowerPointPasteFacts"/>
/// is a record, so each axis test overrides only the member its axis is about.
/// </remarks>
public sealed class PowerPointPasteValidatorTests
{
    /// <summary>The measured staged child count N (ADR-0066).</summary>
    private const int _stagedChildCount = 4;

    /// <summary>The measured staged group envelope width in points (ADR-0066).</summary>
    private const double _stagedWidthPt = 340;

    /// <summary>The measured staged group envelope height in points (ADR-0066).</summary>
    private const double _stagedHeightPt = 190;

    [Fact]
    public void A_null_range_is_NullRange()
    {
        PowerPointPasteValidation validation =
            PowerPointPasteValidator.Validate(null, HealthyExpectation());

        AssertRefused(validation, PowerPointPasteRefusal.NullRange);
    }

    [Fact]
    public void An_empty_range_is_EmptyRange()
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts() with { Count = 0, GroupCount = 0, GroupShapeType = null },
            HealthyExpectation());

        AssertRefused(validation, PowerPointPasteRefusal.EmptyRange);
    }

    [Fact]
    public void A_count_that_is_not_one_plus_the_staged_children_is_CountMismatch()
    {
        // N = 4 expects 5; the host returned 6, so the range is not the measured form.
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts() with { Count = _stagedChildCount + 2 },
            HealthyExpectation());

        AssertRefused(validation, PowerPointPasteRefusal.CountMismatch);
    }

    [Fact]
    public void A_range_with_no_group_is_NoGroup()
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts() with { GroupCount = 0, GroupShapeType = null },
            HealthyExpectation());

        AssertRefused(validation, PowerPointPasteRefusal.NoGroup);
    }

    [Fact]
    public void A_range_with_two_groups_is_MultipleGroups()
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts() with { GroupCount = 2 },
            HealthyExpectation());

        AssertRefused(validation, PowerPointPasteRefusal.MultipleGroups);
    }

    [Fact]
    public void A_group_of_the_wrong_type_is_WrongTypeGroup()
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts() with { GroupShapeType = 1 },
            HealthyExpectation());

        AssertRefused(validation, PowerPointPasteRefusal.WrongTypeGroup);
    }

    [Theory]
    [InlineData(double.NaN, _stagedHeightPt)]
    [InlineData(_stagedWidthPt, double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity, _stagedHeightPt)]
    public void Non_finite_bounds_are_NonFiniteBounds(double widthPt, double heightPt)
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts() with { WidthPt = widthPt, HeightPt = heightPt },
            HealthyExpectation());

        AssertRefused(validation, PowerPointPasteRefusal.NonFiniteBounds);
    }

    [Theory]
    [InlineData(0, _stagedHeightPt)]
    [InlineData(_stagedWidthPt, 0)]
    [InlineData(-1, _stagedHeightPt)]
    public void Non_positive_bounds_are_NonPositiveBounds(double widthPt, double heightPt)
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts() with { WidthPt = widthPt, HeightPt = heightPt },
            HealthyExpectation());

        AssertRefused(validation, PowerPointPasteRefusal.NonPositiveBounds);
    }

    [Fact]
    public void Bounds_beyond_the_tolerance_are_BoundsMismatch_with_the_delta_recorded()
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts() with { WidthPt = _stagedWidthPt + 0.5 },
            HealthyExpectation());

        AssertRefused(validation, PowerPointPasteRefusal.BoundsMismatch);
        Assert.Equal(0.5, validation.MeasuredBoundsDeltaPt, precision: 6);
    }

    [Fact]
    public void Bounds_inside_the_tolerance_are_accepted_and_the_deltas_recorded()
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts() with { WidthPt = _stagedWidthPt + 0.005 },
            HealthyExpectation());

        Assert.True(validation.Succeeded);
        Assert.Null(validation.Refusal);
        Assert.Equal(0.005, validation.MeasuredBoundsDeltaPt, precision: 6);
        Assert.True(validation.MeasuredAspectDelta > 0);
    }

    [Fact]
    public void A_valid_result_forwards_the_validated_facts()
    {
        PowerPointPasteFacts facts = HealthyFacts();

        PowerPointPasteValidation validation =
            PowerPointPasteValidator.Validate(facts, HealthyExpectation());

        Assert.True(validation.Succeeded);
        Assert.Null(validation.Refusal);
        Assert.Same(facts, validation.ValidatedFacts);
        Assert.Equal(_stagedChildCount + 1, validation.ValidatedFacts!.Count);
        Assert.Equal(1, validation.ValidatedFacts.GroupCount);
        Assert.Equal(PowerPointPasteValidator.MsoGroupShapeType, validation.ValidatedFacts.GroupShapeType);
    }

    [Fact]
    public void A_zero_child_expectation_is_InvalidExpectedChildCount()
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts(),
            HealthyExpectation() with { ExpectedChildCount = 0 });

        AssertRefused(validation, PowerPointPasteRefusal.InvalidExpectedChildCount);
    }

    [Theory]
    [InlineData(0, _stagedHeightPt)]
    [InlineData(_stagedWidthPt, -1)]
    [InlineData(double.NaN, _stagedHeightPt)]
    public void Invalid_expected_bounds_are_InvalidExpectedBounds(double widthPt, double heightPt)
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts(),
            HealthyExpectation() with { ExpectedWidthPt = widthPt, ExpectedHeightPt = heightPt });

        AssertRefused(validation, PowerPointPasteRefusal.InvalidExpectedBounds);
    }

    [Fact]
    public void A_null_expectation_is_refused()
    {
        ArgumentNullException exception = Assert.Throws<ArgumentNullException>(
            () => PowerPointPasteValidator.Validate(HealthyFacts(), null!));

        Assert.Equal("expectation", exception.ParamName);
    }

    [Fact]
    public void An_invalid_expectation_is_refused_before_a_null_range()
    {
        // Precedence is total: a bad expectation is a caller error, so it wins over a host
        // result that could never have satisfied it.
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            null,
            HealthyExpectation() with { ExpectedChildCount = 0 });

        AssertRefused(validation, PowerPointPasteRefusal.InvalidExpectedChildCount);
    }

    [Fact]
    public void A_count_mismatch_wins_over_a_bounds_defect()
    {
        PowerPointPasteValidation validation = PowerPointPasteValidator.Validate(
            HealthyFacts() with { Count = _stagedChildCount + 2, WidthPt = double.NaN },
            HealthyExpectation());

        AssertRefused(validation, PowerPointPasteRefusal.CountMismatch);
    }

    // Core cannot name the PowerPoint interop enum; the documented PpShapeType.msoGroup
    // value is a constant pinned here against the documented enumeration.
    [Fact]
    public void The_documented_mso_group_shape_type_is_six() =>
        Assert.Equal(6, PowerPointPasteValidator.MsoGroupShapeType);

    private static PowerPointPasteExpectation HealthyExpectation() =>
        new(_stagedChildCount, _stagedWidthPt, _stagedHeightPt);

    private static PowerPointPasteFacts HealthyFacts() =>
        new(
            Count: _stagedChildCount + 1,
            GroupCount: 1,
            GroupShapeType: PowerPointPasteValidator.MsoGroupShapeType,
            WidthPt: _stagedWidthPt,
            HeightPt: _stagedHeightPt);

    private static void AssertRefused(
        PowerPointPasteValidation validation,
        PowerPointPasteRefusal refusal)
    {
        Assert.False(validation.Succeeded);
        Assert.Equal(refusal, validation.Refusal);
        Assert.Null(validation.ValidatedFacts);
    }
}
