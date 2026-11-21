using GanttCreator.Core.Export;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R7.2: the Office fact gatherer maps a fake host's facts to the Core decision table. The fake
/// port is a plain implementation — no COM is faked — so the tests prove the mapping, the
/// observation-count precedence and the typed-refusal paths without a host.
/// </summary>
public sealed class PowerPointTargetResolverTests
{
    /// <summary>The missing-member COM failure code (DISP_E_MEMBERNOTFOUND).</summary>
    private const int _missingMemberHResult = unchecked((int)0x80020006);

    [Fact]
    public void A_null_port_is_refused()
    {
        ArgumentNullException exception =
            Assert.Throws<ArgumentNullException>(() => new PowerPointTargetResolver(null!));

        Assert.Equal("port", exception.ParamName);
    }

    [Fact]
    public void A_healthy_fake_target_resolves_to_proceed()
    {
        FakePort port = new();

        PowerPointTargetDecision decision = new PowerPointTargetResolver(port).Resolve();

        Assert.Equal(PowerPointTargetAction.ProceedWithActiveSlide, decision.Action);
        Assert.Null(decision.Refusal);
        Assert.Equal(1, decision.ActiveSlideIndex);
        Assert.Null(decision.HResult);
    }

    [Fact]
    public void The_resolver_attaches_once_then_observes_once()
    {
        FakePort port = new();

        _ = new PowerPointTargetResolver(port).Resolve();

        Assert.Equal(1, port.AttachCalls);
        Assert.Equal(1, port.ObservationCalls);
    }

    [Fact]
    public void A_refused_attach_resolves_to_AttachRefused()
    {
        FakePort port = new()
        {
            ApplicationFacts = new PowerPointApplicationFacts(
                Available: false,
                StartedByUs: false,
                ProcessId: 0,
                PresentationCount: 0,
                Build: null,
                HResult: unchecked((int)0x80004005)),
        };

        PowerPointTargetDecision decision = new PowerPointTargetResolver(port).Resolve();

        Assert.Equal(PowerPointTargetAction.Refuse, decision.Action);
        Assert.Equal(PowerPointTargetRefusal.AttachRefused, decision.Refusal);
        Assert.Equal(unchecked((int)0x80004005), decision.HResult);
    }

    [Fact]
    public void A_refused_observation_resolves_to_FactsUnreadable_with_the_missing_member_hresult()
    {
        FakePort port = new()
        {
            Observation = EmptyObservation() with
            {
                ReadSucceeded = false,
                HResult = _missingMemberHResult,
            },
        };

        PowerPointTargetDecision decision = new PowerPointTargetResolver(port).Resolve();

        Assert.Equal(PowerPointTargetAction.Refuse, decision.Action);
        Assert.Equal(PowerPointTargetRefusal.FactsUnreadable, decision.Refusal);
        Assert.Equal(_missingMemberHResult, decision.HResult);
    }

    [Fact]
    public void The_observation_count_is_authoritative_over_the_attach_count()
    {
        FakePort port = new()
        {
            // The attach call saw no presentation; a deck was opened before the target read.
            ApplicationFacts = new PowerPointApplicationFacts(
                Available: true,
                StartedByUs: false,
                ProcessId: 42,
                PresentationCount: 0,
                Build: "16.0",
                HResult: null),
        };

        PowerPointTargetDecision decision = new PowerPointTargetResolver(port).Resolve();

        Assert.Equal(PowerPointTargetAction.ProceedWithActiveSlide, decision.Action);
        Assert.NotEqual(PowerPointTargetAction.CreatePresentation, decision.Action);
    }

    [Fact]
    public void No_presentation_resolves_to_CreatePresentation_and_echoes_the_start()
    {
        FakePort port = new()
        {
            ApplicationFacts = new PowerPointApplicationFacts(
                Available: true,
                StartedByUs: true,
                ProcessId: 42,
                PresentationCount: 0,
                Build: "16.0",
                HResult: null),
            Observation = EmptyObservation(),
        };

        PowerPointTargetDecision decision = new PowerPointTargetResolver(port).Resolve();

        Assert.Equal(PowerPointTargetAction.CreatePresentation, decision.Action);
        Assert.Null(decision.Refusal);
        Assert.True(decision.StartedByUs);
        Assert.Null(decision.ActiveSlideIndex);
        Assert.Null(decision.HResult);
    }

    [Fact]
    public void A_foreground_protected_view_resolves_to_ProtectedView()
    {
        FakePort port = new()
        {
            Observation = HealthyObservation() with
            {
                ProtectedViewWindowCount = 1,
                ActiveProtectedViewWindowPresent = true,
            },
        };

        PowerPointTargetDecision decision = new PowerPointTargetResolver(port).Resolve();

        Assert.Equal(PowerPointTargetAction.Refuse, decision.Action);
        Assert.Equal(PowerPointTargetRefusal.ProtectedView, decision.Refusal);
    }

    [Fact]
    public void A_readable_empty_state_is_a_typed_outcome_not_an_exception()
    {
        FakePort port = new() { Observation = EmptyObservation() };

        PowerPointTargetDecision decision = new PowerPointTargetResolver(port).Resolve();

        Assert.Equal(PowerPointTargetAction.CreatePresentation, decision.Action);
        Assert.Null(decision.Refusal);
    }

    private static PowerPointTargetObservation HealthyObservation() => new(
        ReadSucceeded: true,
        HResult: null,
        PresentationCount: 1,
        ProtectedViewWindowCount: 0,
        ActiveProtectedViewWindowPresent: false,
        ActivePresentationReadOnly: false,
        ActiveWindowPresent: true,
        ActiveViewType: 1,
        SlideShowActive: false,
        SlideCount: 1,
        ActiveSlideIndex: 1);

    private static PowerPointTargetObservation EmptyObservation() => new(
        ReadSucceeded: true,
        HResult: null,
        PresentationCount: 0,
        ProtectedViewWindowCount: 0,
        ActiveProtectedViewWindowPresent: false,
        ActivePresentationReadOnly: false,
        ActiveWindowPresent: false,
        ActiveViewType: null,
        SlideShowActive: false,
        SlideCount: 0,
        ActiveSlideIndex: null);

    private sealed class FakePort : IPowerPointApplicationPort
    {
        internal PowerPointApplicationFacts ApplicationFacts { get; init; } = new(
            Available: true,
            StartedByUs: false,
            ProcessId: 42,
            PresentationCount: 1,
            Build: "16.0",
            HResult: null);

        internal PowerPointTargetObservation Observation { get; init; } = HealthyObservation();

        internal int AttachCalls { get; private set; }

        internal int ObservationCalls { get; private set; }

        public PowerPointApplicationFacts AttachOrStart()
        {
            AttachCalls++;
            return ApplicationFacts;
        }

        public PowerPointPresentationFacts ReadPresentationFacts() =>
            throw new InvalidOperationException("The resolver must use the target observation, not the landed fact read.");

        public PowerPointTargetObservation ReadTargetObservation()
        {
            ObservationCalls++;
            return Observation;
        }
    }
}
