using GanttCreator.Core.Export;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R7.3: the paste stage's host-read mapping over a fake port. A fake host's typed facts
/// reach the Core validator; a refused paste observation becomes the D4 typed stage failure;
/// a non-throwing paste that fails validation is not success; and the classifier's table is
/// pinned without a host.
/// </summary>
/// <remarks>
/// The adapter's COM traversal is what the live gate exercises — this file deliberately owns
/// no host, following <see cref="PowerPointTransferAdapterTests"/>. The exception path is
/// covered host-free through the adapter's refused-observation seam: the adapter's catch
/// delegates to <see cref="PowerPointPasteFailureClassifier.RefusedObservation"/>, and these
/// tests drive that observation through the stage.
/// </remarks>
public sealed class PowerPointPasteStageTests
{
    /// <summary>The measured staged child count N (ADR-0066).</summary>
    private const int _stagedChildCount = 4;

    /// <summary>The measured staged group envelope width in points (ADR-0066).</summary>
    private const double _stagedWidthPt = 340;

    /// <summary>The measured staged group envelope height in points (ADR-0066).</summary>
    private const double _stagedHeightPt = 190;

    /// <summary>The documented candidate clipboard-empty HRESULT (<c>0x80048240</c>).</summary>
    private const int _clipboardEmptyHResult = unchecked((int)0x80048240);

    /// <summary>A generic COM failure code (E_FAIL).</summary>
    private const int _genericFailureHResult = unchecked((int)0x80004005);

    [Fact]
    public void A_valid_paste_forwards_the_validated_facts_and_the_group_identity()
    {
        PowerPointPasteStageResult result = RunStage(ValidObservation(), HealthyExpectation());

        Assert.True(result.Succeeded);
        Assert.Null(result.Refusal);
        Assert.Null(result.HResult);
        Assert.NotNull(result.Validation);
        Assert.True(result.Validation.Succeeded);

        PowerPointPasteFacts facts = Assert.IsType<PowerPointPasteFacts>(result.Validation.ValidatedFacts);
        Assert.Equal(_stagedChildCount + 1, facts.Count);
        Assert.Equal(1, facts.GroupCount);
        Assert.Equal("GanttCreatorStaging:export-group", result.GroupName);
        Assert.Equal(42, result.GroupId);
    }

    [Fact]
    public void A_non_throwing_paste_that_fails_validation_is_not_success()
    {
        // D3: the paste returned a range and no exception was thrown, but the range is not
        // the measured 1 + N form, so the stage must not report success.
        PowerPointPasteStageResult result = RunStage(
            ValidObservation() with { Count = _stagedChildCount + 2 },
            HealthyExpectation());

        Assert.False(result.Succeeded);
        Assert.Null(result.Refusal);
        Assert.Null(result.HResult);
        Assert.NotNull(result.Validation);
        Assert.Equal(PowerPointPasteRefusal.CountMismatch, result.Validation.Refusal);
    }

    [Fact]
    public void A_refused_paste_with_the_clipboard_empty_hresult_is_PastePayloadMissing()
    {
        PowerPointPasteStageResult result = RunStage(
            PowerPointPasteFailureClassifier.RefusedObservation(_clipboardEmptyHResult),
            HealthyExpectation());

        Assert.False(result.Succeeded);
        Assert.Equal(PowerPointPasteStageRefusal.PastePayloadMissing, result.Refusal);
        Assert.Equal(_clipboardEmptyHResult, result.HResult);
        Assert.Null(result.Validation);
    }

    [Fact]
    public void A_refused_paste_with_an_unrecognised_hresult_is_HostRejected()
    {
        PowerPointPasteStageResult result = RunStage(
            PowerPointPasteFailureClassifier.RefusedObservation(_genericFailureHResult),
            HealthyExpectation());

        Assert.False(result.Succeeded);
        Assert.Equal(PowerPointPasteStageRefusal.HostRejected, result.Refusal);
        Assert.Equal(_genericFailureHResult, result.HResult);
    }

    [Fact]
    public void A_refused_paste_without_an_hresult_is_NoSlide()
    {
        PowerPointPasteStageResult result = RunStage(
            PowerPointPasteFailureClassifier.RefusedObservation(null),
            HealthyExpectation());

        Assert.False(result.Succeeded);
        Assert.Equal(PowerPointPasteStageRefusal.NoSlide, result.Refusal);
        Assert.Null(result.HResult);
    }

    [Fact]
    public void A_paste_special_exception_becomes_the_D4_typed_failure_host_free()
    {
        // The acceptance line, proven without a host: a PasteSpecial COM failure becomes the
        // adapter's refused observation, and the stage maps it to the D4 typed failure.
        PowerPointPasteObservation refused =
            PowerPointPasteFailureClassifier.RefusedObservation(_clipboardEmptyHResult);
        Assert.False(refused.PasteSucceeded);

        PowerPointPasteStageResult result = RunStage(refused, HealthyExpectation());

        Assert.Equal(PowerPointPasteStageRefusal.PastePayloadMissing, result.Refusal);
    }

    [Fact]
    public void The_classifier_maps_the_documented_candidate_clipboard_hresult() =>
        Assert.Equal(
            PowerPointPasteStageRefusal.PastePayloadMissing,
            PowerPointPasteFailureClassifier.Classify(_clipboardEmptyHResult));

    [Fact]
    public void The_classifier_maps_an_unrecognised_hresult_to_HostRejected() =>
        Assert.Equal(
            PowerPointPasteStageRefusal.HostRejected,
            PowerPointPasteFailureClassifier.Classify(_genericFailureHResult));

    [Fact]
    public void The_classifier_maps_a_missing_hresult_to_NoSlide() =>
        Assert.Equal(
            PowerPointPasteStageRefusal.NoSlide,
            PowerPointPasteFailureClassifier.Classify(null));

    [Fact]
    public void A_null_port_is_refused() =>
        _ = Assert.Throws<ArgumentNullException>(() => new PowerPointPasteStage(null!));

    [Fact]
    public void A_null_expectation_is_refused()
    {
        PowerPointPasteStage stage = new(new FakePastePort(ValidObservation()));

        ArgumentNullException exception =
            Assert.Throws<ArgumentNullException>(() => stage.Paste(null!));

        Assert.Equal("expectation", exception.ParamName);
    }

    private static PowerPointPasteExpectation HealthyExpectation() =>
        new(_stagedChildCount, _stagedWidthPt, _stagedHeightPt);

    private static PowerPointPasteObservation ValidObservation() =>
        new(
            PasteSucceeded: true,
            HResult: null,
            Count: _stagedChildCount + 1,
            GroupCount: 1,
            GroupShapeType: PowerPointPasteValidator.MsoGroupShapeType,
            GroupChildCount: _stagedChildCount,
            GroupName: "GanttCreatorStaging:export-group",
            GroupId: 42,
            WidthPt: _stagedWidthPt,
            HeightPt: _stagedHeightPt);

    private static PowerPointPasteStageResult RunStage(
        PowerPointPasteObservation observation,
        PowerPointPasteExpectation expectation) =>
        new PowerPointPasteStage(new FakePastePort(observation)).Paste(expectation);

    private sealed class FakePastePort(PowerPointPasteObservation observation)
        : IPowerPointPastePort
    {
        public PowerPointPasteObservation PasteSpecialShape() => observation;

        public PowerPointGroupDeleteResult DeletePastedGroup(int groupId) =>
            throw new InvalidOperationException("The paste stage never deletes residue.");
    }
}
