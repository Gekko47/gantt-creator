using GanttCreator.Core.Export;

namespace GanttCreator.Core.Tests.Export;

/// <summary>
/// R7.2: the complete attach/start and active-slide decision matrix. Every case maps to
/// exactly one typed outcome, every refusal has a positive test that constructs the bad
/// facts, and the precedence between overlapping conditions is pinned.
/// </summary>
/// <remarks>
/// <see cref="PowerPointTargetFacts"/> is a record, so each case starts from
/// <see cref="HealthyFacts"/> and overrides only the facts its axis is about. That keeps a
/// test's intent visible: a test that reads-only must not accidentally also vary the view.
/// </remarks>
public sealed class PowerPointTargetDecisionTests
{
    /// <summary>The generic COM failure code (E_FAIL).</summary>
    private const int _genericFailureHResult = unchecked((int)0x80004005);

    /// <summary>The access-denied COM failure code (E_ACCESSDENIED).</summary>
    private const int _accessDeniedHResult = unchecked((int)0x80070005);

    /// <summary>The missing-member COM failure code (DISP_E_MEMBERNOTFOUND).</summary>
    private const int _missingMemberHResult = unchecked((int)0x80020006);

    [Fact]
    public void Decide_refuses_null_facts()
    {
        ArgumentNullException exception =
            Assert.Throws<ArgumentNullException>(() => PowerPointTargetDecision.Decide(null!));

        Assert.Equal("facts", exception.ParamName);
    }

    [Fact]
    public void A_refused_attach_is_AttachRefused_with_the_hresult()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                AttachAvailable = false,
                AttachHResult = _genericFailureHResult,
            });

        AssertRefused(decision, PowerPointTargetRefusal.AttachRefused, _genericFailureHResult);
    }

    [Fact]
    public void A_refused_read_is_FactsUnreadable_with_the_missing_member_hresult()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                ReadSucceeded = false,
                ReadHResult = _missingMemberHResult,
            });

        AssertRefused(decision, PowerPointTargetRefusal.FactsUnreadable, _missingMemberHResult);
    }

    [Fact]
    public void A_running_slide_show_is_refused()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with { SlideShowActive = true });

        AssertRefused(decision, PowerPointTargetRefusal.SlideShowActive);
    }

    [Fact]
    public void A_foreground_protected_view_window_is_refused_over_a_background_deck()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                PresentationCount = 1,
                ProtectedViewWindowCount = 1,
                ActiveProtectedViewWindowPresent = true,
            });

        AssertRefused(decision, PowerPointTargetRefusal.ProtectedView);
    }

    [Fact]
    public void No_presentation_with_a_protected_view_window_is_refused_rather_than_created()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                PresentationCount = 0,
                ProtectedViewWindowCount = 1,
                ActiveWindowPresent = false,
                ActiveViewType = null,
                SlideCount = 0,
                ActiveSlideIndex = null,
            });

        AssertRefused(decision, PowerPointTargetRefusal.ProtectedView);
    }

    [Fact]
    public void A_background_protected_view_window_does_not_block_a_healthy_deck()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                PresentationCount = 1,
                ProtectedViewWindowCount = 1,
                ActiveProtectedViewWindowPresent = false,
            });

        AssertProceeded(decision, activeSlideIndex: 1);
    }

    [Fact]
    public void No_presentation_and_no_protected_view_creates_a_presentation()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                StartedByUs = true,
                PresentationCount = 0,
                ProtectedViewWindowCount = 0,
                ActiveWindowPresent = false,
                ActiveViewType = null,
                SlideCount = 0,
                ActiveSlideIndex = null,
            });

        Assert.Equal(PowerPointTargetAction.CreatePresentation, decision.Action);
        Assert.Null(decision.Refusal);
        Assert.True(decision.StartedByUs);
        Assert.Null(decision.ActiveSlideIndex);
        Assert.Null(decision.HResult);
    }

    [Fact]
    public void An_attached_host_with_no_presentation_creates_a_presentation_without_claiming_a_start()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                StartedByUs = false,
                PresentationCount = 0,
                ProtectedViewWindowCount = 0,
                ActiveWindowPresent = false,
                ActiveViewType = null,
                SlideCount = 0,
                ActiveSlideIndex = null,
            });

        Assert.Equal(PowerPointTargetAction.CreatePresentation, decision.Action);
        Assert.False(decision.StartedByUs);
    }

    [Fact]
    public void A_read_only_presentation_is_refused()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with { ActivePresentationReadOnly = true });

        AssertRefused(decision, PowerPointTargetRefusal.ReadOnlyPresentation);
    }

    [Fact]
    public void A_windowless_presentation_is_refused()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                ActiveWindowPresent = false,
                ActiveViewType = null,
                SlideCount = 0,
                ActiveSlideIndex = null,
            });

        AssertRefused(decision, PowerPointTargetRefusal.NoWindow);
    }

    [Fact]
    public void A_null_view_type_with_a_window_is_refused_as_an_unsupported_view()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with { ActiveViewType = null });

        AssertRefused(decision, PowerPointTargetRefusal.UnsupportedView);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(9)]
    public void A_pasteable_view_proceeds_with_the_resolved_slide(int viewType)
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with { ActiveViewType = viewType });

        AssertProceeded(decision, activeSlideIndex: 1);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(4)]
    [InlineData(5)]
    [InlineData(6)]
    [InlineData(7)]
    [InlineData(8)]
    [InlineData(10)]
    [InlineData(11)]
    [InlineData(12)]
    [InlineData(99)]
    public void A_non_pasteable_view_is_refused_as_an_unsupported_view(int viewType)
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with { ActiveViewType = viewType });

        AssertRefused(decision, PowerPointTargetRefusal.UnsupportedView);
    }

    [Fact]
    public void A_presentation_with_no_slides_is_refused()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                SlideCount = 0,
                ActiveSlideIndex = null,
            });

        AssertRefused(decision, PowerPointTargetRefusal.NoSlide);
    }

    [Fact]
    public void A_view_with_no_active_slide_is_refused()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                SlideCount = 2,
                ActiveSlideIndex = null,
            });

        AssertRefused(decision, PowerPointTargetRefusal.NoActiveSlide);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void One_or_many_presentations_proceed_against_the_resolved_active_slide(int presentationCount)
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                PresentationCount = presentationCount,
                ActiveSlideIndex = 4,
            });

        AssertProceeded(decision, activeSlideIndex: 4);
    }

    [Fact]
    public void The_healthy_target_proceeds_and_echoes_the_start_announcement()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with { StartedByUs = true });

        Assert.Equal(PowerPointTargetAction.ProceedWithActiveSlide, decision.Action);
        Assert.Null(decision.Refusal);
        Assert.True(decision.StartedByUs);
        Assert.Equal(1, decision.ActiveSlideIndex);
        Assert.Null(decision.HResult);
    }

    [Fact]
    public void A_slide_show_beats_a_read_only_presentation()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                SlideShowActive = true,
                ActivePresentationReadOnly = true,
            });

        AssertRefused(decision, PowerPointTargetRefusal.SlideShowActive);
    }

    [Fact]
    public void A_foreground_protected_view_beats_creation()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                PresentationCount = 0,
                ProtectedViewWindowCount = 1,
                ActiveProtectedViewWindowPresent = true,
                ActiveWindowPresent = false,
                ActiveViewType = null,
                SlideCount = 0,
                ActiveSlideIndex = null,
            });

        AssertRefused(decision, PowerPointTargetRefusal.ProtectedView);
    }

    [Fact]
    public void A_read_only_presentation_beats_a_missing_window()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                ActivePresentationReadOnly = true,
                ActiveWindowPresent = false,
                ActiveViewType = null,
            });

        AssertRefused(decision, PowerPointTargetRefusal.ReadOnlyPresentation);
    }

    [Fact]
    public void An_unsupported_view_beats_a_missing_slide()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                ActiveViewType = 7,
                SlideCount = 0,
                ActiveSlideIndex = null,
            });

        AssertRefused(decision, PowerPointTargetRefusal.UnsupportedView);
    }

    [Fact]
    public void A_refused_attach_beats_every_later_case()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                AttachAvailable = false,
                AttachHResult = _accessDeniedHResult,
                SlideShowActive = true,
                ActiveProtectedViewWindowPresent = true,
            });

        AssertRefused(decision, PowerPointTargetRefusal.AttachRefused, _accessDeniedHResult);
    }

    [Fact]
    public void A_refused_read_beats_every_later_case()
    {
        PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(
            HealthyFacts() with
            {
                ReadSucceeded = false,
                ReadHResult = _missingMemberHResult,
                SlideShowActive = true,
                ActiveProtectedViewWindowPresent = true,
            });

        AssertRefused(decision, PowerPointTargetRefusal.FactsUnreadable, _missingMemberHResult);
    }

    [Fact]
    public void The_same_facts_decide_the_same_outcome()
    {
        PowerPointTargetFacts facts = HealthyFacts() with { ActiveSlideIndex = 3 };

        PowerPointTargetDecision first = PowerPointTargetDecision.Decide(facts);
        PowerPointTargetDecision second = PowerPointTargetDecision.Decide(facts);

        Assert.Equal(first, second);
    }

    [Fact]
    public void The_pasteable_view_set_is_slide_notes_page_and_normal() =>
        Assert.Equal([1, 3, 9], PowerPointTargetDecision.PasteableViewTypes.Order());

    [Fact]
    public void Only_the_unreadable_refusals_carry_an_hresult()
    {
        (PowerPointTargetFacts Facts, PowerPointTargetRefusal Refusal)[] cases =
        [
            (HealthyFacts() with { AttachAvailable = false, AttachHResult = _genericFailureHResult }, PowerPointTargetRefusal.AttachRefused),
            (HealthyFacts() with { ReadSucceeded = false, ReadHResult = _missingMemberHResult }, PowerPointTargetRefusal.FactsUnreadable),
            (HealthyFacts() with { SlideShowActive = true }, PowerPointTargetRefusal.SlideShowActive),
            (HealthyFacts() with { ActiveProtectedViewWindowPresent = true }, PowerPointTargetRefusal.ProtectedView),
            (HealthyFacts() with { ActivePresentationReadOnly = true }, PowerPointTargetRefusal.ReadOnlyPresentation),
            (HealthyFacts() with { ActiveWindowPresent = false, ActiveViewType = null }, PowerPointTargetRefusal.NoWindow),
            (HealthyFacts() with { ActiveViewType = 7 }, PowerPointTargetRefusal.UnsupportedView),
            (HealthyFacts() with { SlideCount = 0 }, PowerPointTargetRefusal.NoSlide),
            (HealthyFacts() with { ActiveSlideIndex = null }, PowerPointTargetRefusal.NoActiveSlide),
        ];

        Assert.All(
            cases,
            item =>
            {
                PowerPointTargetDecision decision = PowerPointTargetDecision.Decide(item.Facts);

                Assert.Equal(PowerPointTargetAction.Refuse, decision.Action);
                Assert.Equal(item.Refusal, decision.Refusal);
                Assert.Null(decision.ActiveSlideIndex);

                if (item.Refusal is PowerPointTargetRefusal.AttachRefused or PowerPointTargetRefusal.FactsUnreadable)
                {
                    _ = Assert.NotNull(decision.HResult);
                }
                else
                {
                    Assert.Null(decision.HResult);
                }
            });
    }

    private static void AssertRefused(
        PowerPointTargetDecision decision,
        PowerPointTargetRefusal expectedRefusal,
        int? expectedHResult = null)
    {
        Assert.Equal(PowerPointTargetAction.Refuse, decision.Action);
        Assert.Equal(expectedRefusal, decision.Refusal);
        Assert.Null(decision.ActiveSlideIndex);
        Assert.Equal(expectedHResult, decision.HResult);
    }

    private static void AssertProceeded(PowerPointTargetDecision decision, int activeSlideIndex)
    {
        Assert.Equal(PowerPointTargetAction.ProceedWithActiveSlide, decision.Action);
        Assert.Null(decision.Refusal);
        Assert.Equal(activeSlideIndex, decision.ActiveSlideIndex);
        Assert.Null(decision.HResult);
    }

    private static PowerPointTargetFacts HealthyFacts() => new(
        AttachAvailable: true,
        AttachHResult: null,
        StartedByUs: false,
        ReadSucceeded: true,
        ReadHResult: null,
        PresentationCount: 1,
        ProtectedViewWindowCount: 0,
        ActiveProtectedViewWindowPresent: false,
        ActivePresentationReadOnly: false,
        ActiveWindowPresent: true,
        ActiveViewType: 1,
        SlideShowActive: false,
        SlideCount: 1,
        ActiveSlideIndex: 1);
}
