using GanttCreator.Core.Export;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R7.5 D1: the mapping matrix — every decision-table refusal arm, the two non-refusal arms,
/// the creation arm, the R6.6 clipboard member, the paste and placement stage refusals, and the
/// construction members. Every taxonomy member is reachable and every arm has a positive test.
/// </summary>
public sealed class PowerPointTransferFailureMapperTests
{
    private const int _memberNotFound = unchecked((int)0x80020006);
    private const int _transient = unchecked((int)0x80010001);
    private const int _unrecognised = unchecked((int)0x80004005);

    // ----- R7.2 target decision: every refusal arm -----

    [Theory]
    [InlineData(PowerPointTargetRefusal.SlideShowActive, PowerPointTransferFailureKind.SlideShowActive)]
    [InlineData(PowerPointTargetRefusal.ProtectedView, PowerPointTransferFailureKind.PresentationProtectedView)]
    [InlineData(PowerPointTargetRefusal.ReadOnlyPresentation, PowerPointTransferFailureKind.PresentationReadOnly)]
    [InlineData(PowerPointTargetRefusal.NoWindow, PowerPointTransferFailureKind.NoWindow)]
    [InlineData(PowerPointTargetRefusal.UnsupportedView, PowerPointTransferFailureKind.UnsupportedView)]
    [InlineData(PowerPointTargetRefusal.NoSlide, PowerPointTransferFailureKind.NoSlide)]
    [InlineData(PowerPointTargetRefusal.NoActiveSlide, PowerPointTransferFailureKind.NoSlide)]
    public void Each_simple_target_refusal_maps_to_its_taxonomy_member(
        PowerPointTargetRefusal refusal,
        PowerPointTransferFailureKind expected)
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapTargetDecision(Refused(refusal)));

        Assert.Equal(expected, failure.Kind);
        Assert.False(string.IsNullOrWhiteSpace(failure.Message));
        Assert.Contains("stage=target", failure.TechnicalRecord, StringComparison.Ordinal);
    }

    [Fact]
    public void Attach_refused_with_the_missing_member_hresult_is_UnsupportedBuild()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapTargetDecision(
                Refused(PowerPointTargetRefusal.AttachRefused, _memberNotFound)));

        Assert.Equal(PowerPointTransferFailureKind.UnsupportedBuild, failure.Kind);
        Assert.Equal(_memberNotFound, failure.HResult);
    }

    [Fact]
    public void Attach_refused_with_an_unrecognised_hresult_is_HostRejected()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapTargetDecision(
                Refused(PowerPointTargetRefusal.AttachRefused, _unrecognised)));

        Assert.Equal(PowerPointTransferFailureKind.HostRejected, failure.Kind);
        Assert.Equal(_unrecognised, failure.HResult);
    }

    [Fact]
    public void Facts_unreadable_with_a_transient_hresult_is_HostBusy()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapTargetDecision(
                Refused(PowerPointTargetRefusal.FactsUnreadable, _transient)));

        Assert.Equal(PowerPointTransferFailureKind.HostBusy, failure.Kind);
        Assert.Equal(_transient, failure.HResult);
    }

    [Fact]
    public void A_proceed_decision_is_not_a_failure() =>
        Assert.Null(
            PowerPointTransferFailureMapper.MapTargetDecision(
                new PowerPointTargetDecision(
                    PowerPointTargetAction.ProceedWithActiveSlide,
                    Refusal: null,
                    StartedByUs: false,
                    ActiveSlideIndex: 1,
                    HResult: null)));

    [Fact]
    public void A_create_decision_is_not_a_failure() =>
        Assert.Null(
            PowerPointTransferFailureMapper.MapTargetDecision(
                new PowerPointTargetDecision(
                    PowerPointTargetAction.CreatePresentation,
                    Refusal: null,
                    StartedByUs: true,
                    ActiveSlideIndex: null,
                    HResult: null)));

    // ----- creation (PD-2) -----

    [Fact]
    public void A_refused_creation_is_NoPresentation()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapCreation(
                new PowerPointPresentationCreation(Created: false, SlideIndex: null, HResult: _unrecognised)));

        Assert.Equal(PowerPointTransferFailureKind.NoPresentation, failure.Kind);
        Assert.Equal(_unrecognised, failure.HResult);
        Assert.Contains("stage=create", failure.TechnicalRecord, StringComparison.Ordinal);
    }

    [Fact]
    public void A_successful_creation_is_not_a_failure() =>
        Assert.Null(
            PowerPointTransferFailureMapper.MapCreation(
                new PowerPointPresentationCreation(Created: true, SlideIndex: 1, HResult: null)));

    // ----- R6.6 copy -----

    [Fact]
    public void A_clipboard_unavailable_copy_is_ClipboardUnavailable()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapCopyResult(
                new StagingCopyResult(
                    Succeeded: false,
                    Detail: "The clipboard was busy (clipboard-unavailable:0x00000005).",
                    Refusal: StagedCopyRefusal.ClipboardUnavailable)));

        Assert.Equal(PowerPointTransferFailureKind.ClipboardUnavailable, failure.Kind);
        Assert.Contains("clipboard-unavailable", failure.TechnicalRecord, StringComparison.Ordinal);
    }

    [Fact]
    public void A_non_clipboard_copy_refusal_is_not_a_transfer_failure() =>
        Assert.Null(
            PowerPointTransferFailureMapper.MapCopyResult(
                new StagingCopyResult(
                    Succeeded: false,
                    Detail: "copy refused",
                    Refusal: StagedCopyRefusal.CopyRefused)));

    [Fact]
    public void A_successful_copy_is_not_a_failure() =>
        Assert.Null(PowerPointTransferFailureMapper.MapCopyResult(StagingCopyResult.Success));

    [Fact]
    public void A_null_copy_result_is_not_a_failure() =>
        Assert.Null(PowerPointTransferFailureMapper.MapCopyResult(null));

    // ----- R7.3 paste stage -----

    [Fact]
    public void Paste_payload_missing_maps_to_PastePayloadMissing()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPasteStage(
                PasteRefused(PowerPointPasteStageRefusal.PastePayloadMissing, _unrecognised)));

        Assert.Equal(PowerPointTransferFailureKind.PastePayloadMissing, failure.Kind);
    }

    [Fact]
    public void Partial_document_maps_to_PartialDocument()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPasteStage(
                PasteRefused(PowerPointPasteStageRefusal.PartialDocument, _unrecognised)));

        Assert.Equal(PowerPointTransferFailureKind.PartialDocument, failure.Kind);
    }

    [Fact]
    public void Paste_no_slide_maps_to_NoSlide()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPasteStage(
                PasteRefused(PowerPointPasteStageRefusal.NoSlide)));

        Assert.Equal(PowerPointTransferFailureKind.NoSlide, failure.Kind);
    }

    [Theory]
    [InlineData(_memberNotFound, PowerPointTransferFailureKind.UnsupportedBuild)]
    [InlineData(_transient, PowerPointTransferFailureKind.HostBusy)]
    [InlineData(_unrecognised, PowerPointTransferFailureKind.HostRejected)]
    public void Paste_host_rejected_is_refined_by_the_hresult(int hresult, PowerPointTransferFailureKind expected)
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPasteStage(
                PasteRefused(PowerPointPasteStageRefusal.HostRejected, hresult)));

        Assert.Equal(expected, failure.Kind);
        Assert.Equal(hresult, failure.HResult);
    }

    [Fact]
    public void A_paste_that_fails_validation_is_PasteValidationFailed()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPasteStage(
                new PowerPointPasteStageResult(
                    Succeeded: false,
                    Refusal: null,
                    HResult: null,
                    Validation: FailedValidation(),
                    GroupName: null,
                    GroupId: null)));

        Assert.Equal(PowerPointTransferFailureKind.PasteValidationFailed, failure.Kind);
        Assert.Contains("CountMismatch", failure.TechnicalRecord, StringComparison.Ordinal);
    }

    [Fact]
    public void The_mapper_never_overrides_a_specific_stage_classification()
    {
        // The stage already classified this as PastePayloadMissing; a missing-member HRESULT is
        // carried but must not be re-derived into UnsupportedBuild (the never-override rule).
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPasteStage(
                PasteRefused(PowerPointPasteStageRefusal.PastePayloadMissing, _memberNotFound)));

        Assert.Equal(PowerPointTransferFailureKind.PastePayloadMissing, failure.Kind);
        Assert.Equal(_memberNotFound, failure.HResult);
    }

    // ----- R7.4 placement stage -----

    [Theory]
    [InlineData(PowerPointPlacementRefusal.AlreadyApplied, PowerPointTransferFailureKind.TransferSequenceViolated)]
    [InlineData(PowerPointPlacementRefusal.GroupMissing, PowerPointTransferFailureKind.HostRejected)]
    [InlineData(PowerPointPlacementRefusal.ReadBackMismatch, PowerPointTransferFailureKind.HostRejected)]
    [InlineData(PowerPointPlacementRefusal.NoSlide, PowerPointTransferFailureKind.NoSlide)]
    public void Each_simple_placement_refusal_maps_to_its_taxonomy_member(
        PowerPointPlacementRefusal refusal,
        PowerPointTransferFailureKind expected)
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPlacement(PlacementRefused(refusal)));

        Assert.Equal(expected, failure.Kind);
    }

    [Theory]
    [InlineData(_memberNotFound, PowerPointTransferFailureKind.UnsupportedBuild)]
    [InlineData(_transient, PowerPointTransferFailureKind.HostBusy)]
    [InlineData(_unrecognised, PowerPointTransferFailureKind.HostRejected)]
    public void Placement_host_rejected_is_refined_by_the_hresult(int hresult, PowerPointTransferFailureKind expected)
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPlacement(
                PlacementRefused(PowerPointPlacementRefusal.HostRejected, hresult)));

        Assert.Equal(expected, failure.Kind);
    }

    [Fact]
    public void Slide_size_unavailable_without_an_hresult_is_a_sequence_violation()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPlacement(
                PlacementRefused(PowerPointPlacementRefusal.SlideSizeUnavailable)));

        Assert.Equal(PowerPointTransferFailureKind.TransferSequenceViolated, failure.Kind);
        Assert.Contains("stage=placement-size", failure.TechnicalRecord, StringComparison.Ordinal);
    }

    [Fact]
    public void Slide_size_unavailable_with_the_missing_member_hresult_is_UnsupportedBuild()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPlacement(
                PlacementRefused(PowerPointPlacementRefusal.SlideSizeUnavailable, _memberNotFound)));

        Assert.Equal(PowerPointTransferFailureKind.UnsupportedBuild, failure.Kind);
    }

    [Fact]
    public void A_fit_refusal_is_an_invalid_transfer_request()
    {
        PowerPointTransferFailure failure = Failure(
            PowerPointTransferFailureMapper.MapPlacement(
                new PowerPointPlacementResult(
                    Succeeded: false,
                    Refusal: null,
                    HResult: null,
                    FitRefusal: SlideFitRefusal.InvalidGroupBounds,
                    Placement: null,
                    Observation: null,
                    Deltas: null)));

        Assert.Equal(PowerPointTransferFailureKind.InvalidTransferRequest, failure.Kind);
        Assert.Contains("stage=placement-fit", failure.TechnicalRecord, StringComparison.Ordinal);
    }

    [Fact]
    public void A_verified_placement_is_not_a_failure() =>
        Assert.Null(
            PowerPointTransferFailureMapper.MapPlacement(
                new PowerPointPlacementResult(
                    Succeeded: true,
                    Refusal: null,
                    HResult: null,
                    FitRefusal: null,
                    Placement: null,
                    Observation: null,
                    Deltas: null)));

    // ----- construction and ownership members -----

    [Fact]
    public void An_invalid_request_is_the_construction_member()
    {
        PowerPointTransferFailure failure = PowerPointTransferFailureMapper.InvalidRequest("null-placement");

        Assert.Equal(PowerPointTransferFailureKind.InvalidTransferRequest, failure.Kind);
        Assert.Null(failure.HResult);
        Assert.Contains("detail=null-placement", failure.TechnicalRecord, StringComparison.Ordinal);
    }

    [Fact]
    public void A_sequence_violation_is_a_construction_member()
    {
        PowerPointTransferFailure failure = PowerPointTransferFailureMapper.SequenceViolated("disposed");

        Assert.Equal(PowerPointTransferFailureKind.TransferSequenceViolated, failure.Kind);
        Assert.Contains("stage=sequence", failure.TechnicalRecord, StringComparison.Ordinal);
    }

    [Fact]
    public void A_blank_guard_is_refused() =>
        _ = Assert.Throws<ArgumentException>(() => PowerPointTransferFailureMapper.InvalidRequest(" "));

    [Fact]
    public void A_blank_sequence_reason_is_refused() =>
        _ = Assert.Throws<ArgumentException>(() => PowerPointTransferFailureMapper.SequenceViolated(" "));

    [Fact]
    public void An_ownership_failure_carries_the_release_hresult()
    {
        PowerPointTransferFailure failure = PowerPointTransferFailureMapper.OwnershipFailed(_unrecognised);

        Assert.Equal(PowerPointTransferFailureKind.OwnershipFailed, failure.Kind);
        Assert.Equal(_unrecognised, failure.HResult);
        Assert.Contains("stage=ownership", failure.TechnicalRecord, StringComparison.Ordinal);
    }

    [Fact]
    public void The_host_classifier_returns_the_three_host_members()
    {
        Assert.Equal(
            PowerPointTransferFailureKind.UnsupportedBuild,
            PowerPointTransferFailureMapper.ClassifyHostFailure(_memberNotFound));
        Assert.Equal(
            PowerPointTransferFailureKind.HostBusy,
            PowerPointTransferFailureMapper.ClassifyHostFailure(_transient));
        Assert.Equal(
            PowerPointTransferFailureKind.HostRejected,
            PowerPointTransferFailureMapper.ClassifyHostFailure(_unrecognised));
        Assert.Equal(
            PowerPointTransferFailureKind.HostRejected,
            PowerPointTransferFailureMapper.ClassifyHostFailure(null));
    }

    private static PowerPointTargetDecision Refused(PowerPointTargetRefusal refusal, int? hresult = null) =>
        new(
            PowerPointTargetAction.Refuse,
            refusal,
            StartedByUs: false,
            ActiveSlideIndex: null,
            HResult: hresult);

    private static PowerPointPasteStageResult PasteRefused(
        PowerPointPasteStageRefusal refusal,
        int? hresult = null) =>
        new(
            Succeeded: false,
            Refusal: refusal,
            HResult: hresult,
            Validation: null,
            GroupName: null,
            GroupId: null);

    private static PowerPointPlacementResult PlacementRefused(
        PowerPointPlacementRefusal refusal,
        int? hresult = null) =>
        new(
            Succeeded: false,
            Refusal: refusal,
            HResult: hresult,
            FitRefusal: null,
            Placement: null,
            Observation: null,
            Deltas: null);

    private static PowerPointPasteValidation FailedValidation() =>
        PowerPointPasteValidator.Validate(
            new PowerPointPasteFacts(Count: 2, GroupCount: 0, GroupShapeType: null, WidthPt: 10, HeightPt: 10),
            new PowerPointPasteExpectation(4, 340, 190));

    private static PowerPointTransferFailure Failure(PowerPointTransferFailure? failure) =>
        Assert.IsType<PowerPointTransferFailure>(failure);
}
