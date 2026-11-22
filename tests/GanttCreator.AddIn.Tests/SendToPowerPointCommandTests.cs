using GanttCreator.Core;
using GanttCreator.Core.Export;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// R7.6: the Send to PowerPoint command owns the pipeline order and nothing more. These facts
/// pin the stage order (pre-flight then lifecycle then transfer), the transfer chain over the
/// landed R7.1–R7.5 stages, the D-R7.2 announcement, the ownership matrix, the truthful residue
/// cleanup, every typed refusal line, and presenter degradation.
/// </summary>
/// <remarks>
/// The lifecycle seam is faked to invoke the transfer hook it is given, exactly as the landed
/// lifecycle does after a verified copy; the lifecycle's own ordering is pinned by
/// <c>StagingLifecycleTests</c>. The PowerPoint stages are the real production types over fake
/// ports, so the command's composition is what these facts exercise.
/// </remarks>
public class SendToPowerPointCommandTests
{
    private static readonly DateOnly _start = new(2026, 1, 5);
    private static readonly DateOnly _finish = new(2026, 1, 9);

    [Fact]
    public void Run_pins_the_pipeline_order()
    {
        Harness harness = new();

        _ = harness.Run();

        Assert.Equal(
            [
                "read",
                "config",
                "measure",
                "export",
                "resolve",
                "lifecycle",
                "target",
                "paste",
                "read-size",
                "place",
            ],
            harness.Order);
    }

    [Fact]
    public void Run_sends_the_verified_group_to_the_active_slide()
    {
        Harness harness = new();

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.True(outcome.Succeeded, outcome.Message);
        Assert.Equal(3, outcome.StagedShapeCount);
        Assert.False(outcome.PresentationCreated);
        Assert.Equal("Sent 3 shapes to the active slide.", outcome.Message);
        Assert.Equal([outcome.Message], harness.Successes);
        Assert.Empty(harness.Failures);

        // D2: exactly the expected port members were called, in chain order.
        Assert.Equal(["attach", "observe"], harness.Application.Calls);
        Assert.Equal(["paste"], harness.Paste.Calls);
        Assert.Equal(["read-size", "place"], harness.Placement.Calls);
        Assert.Null(outcome.Refusal);
    }

    [Fact]
    public void Run_supplies_the_transfer_hook_to_the_lifecycle()
    {
        Harness harness = new();
        _ = harness.Run();

        Assert.NotNull(harness.SeenLifecycleRequest);
        Assert.NotNull(harness.SeenLifecycleRequest.Transfer);
    }

    [Fact]
    public void Run_uses_the_singular_shape_word_for_one_shape()
    {
        Harness harness = new()
        {
            StagedIds = ["panel-cell:0"],
        };
        harness.Paste.Observation = SuccessfulPasteObservation(1);

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.Equal("Sent 1 shape to the active slide.", outcome.Message);
    }

    /// <summary>
    /// D-R7.2: when no presentation is open the transfer creates one and the success line
    /// announces it. The create member is called exactly once.
    /// </summary>
    [Fact]
    public void Run_creates_and_announces_a_presentation_when_none_is_open()
    {
        Harness harness = new();
        harness.Application.TargetObservation = TargetObservation(presentationCount: 0);

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.True(outcome.Succeeded, outcome.Message);
        Assert.True(outcome.PresentationCreated);
        Assert.Equal(
            "Sent 3 shapes to the active slide. PowerPoint was not open, so a new presentation was created.",
            outcome.Message);
        Assert.Equal(["attach", "observe", "create"], harness.Application.Calls);
    }

    [Fact]
    public void Run_presents_the_mapper_message_for_a_refused_target()
    {
        Harness harness = new();
        harness.Application.TargetObservation = TargetObservation(readOnly: true);

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(SendToPowerPointRefusal.TransferRefused, outcome.Refusal);
        Assert.Equal(
            "The active presentation is read-only or password-protected. Open an editable copy and try again.",
            outcome.Message);
        Assert.Empty(harness.Paste.Calls);
        Assert.Contains(
            harness.Records,
            record => record.Contains("kind=PresentationReadOnly", StringComparison.Ordinal));
    }

    [Fact]
    public void Run_creates_then_reports_the_failure_and_records_the_created_deck()
    {
        Harness harness = new();
        harness.Application.TargetObservation = TargetObservation(presentationCount: 0);
        harness.Paste.Observation = new PowerPointPasteObservation(
            PasteSucceeded: false,
            HResult: unchecked((int)0x80048240),
            Count: 0,
            GroupCount: 0,
            GroupShapeType: null,
            GroupChildCount: 0,
            GroupName: null,
            GroupId: null,
            WidthPt: 0,
            HeightPt: 0);

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(SendToPowerPointRefusal.TransferRefused, outcome.Refusal);
        Assert.True(outcome.PresentationCreated);
        Assert.Equal(
            "The clipboard holds no shape data to paste. Refresh the chart, copy it again, then send it.",
            outcome.Message);
        Assert.Contains(
            harness.Records,
            record => record.Contains("presentation-created=true", StringComparison.Ordinal));
    }

    /// <summary>
    /// A placement failure removes the pasted group; the removal is confirmed, so the failure's
    /// own message is presented and no residue is left.
    /// </summary>
    [Fact]
    public void Run_cleans_up_the_pasted_group_when_placement_fails()
    {
        Harness harness = new();
        harness.Placement.NextObservation = new PowerPointPlacementObservation(
            PowerPointPlacementApplyOutcome.GroupMissing,
            null,
            0,
            0,
            0,
            0);

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(SendToPowerPointRefusal.TransferRefused, outcome.Refusal);
        Assert.Equal("PowerPoint refused the operation. Try again.", outcome.Message);
        Assert.Equal(["paste", "delete"], harness.Paste.Calls);
        Assert.Equal(42, harness.Paste.LastDeletedGroupId);
    }

    /// <summary>
    /// R7.6 review F2: the PasteValidationFailed message claims the pasted shapes were removed.
    /// When the delete is refused, that claim is false, so the truthful host-failure message is
    /// presented and the cleanup record names the leftover.
    /// </summary>
    [Fact]
    public void Run_never_claims_removal_when_the_residue_delete_is_refused()
    {
        Harness harness = new();
        harness.Paste.Observation = SuccessfulPasteObservation(3) with { Count = 99 };
        harness.Paste.DeleteResult = new PowerPointGroupDeleteResult(
            PowerPointGroupDeleteOutcome.HostRejected,
            unchecked((int)0x80004005));

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal("PowerPoint refused the operation. Try again.", outcome.Message);
        Assert.DoesNotContain("were removed", outcome.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(["paste", "delete"], harness.Paste.Calls);
        Assert.Contains(
            harness.Records,
            record => record.Contains("residue-cleanup", StringComparison.Ordinal));
    }

    /// <summary>The positive control: a confirmed delete keeps the mapper's removal claim.</summary>
    [Fact]
    public void Run_claims_removal_when_the_cleanup_confirms_it()
    {
        Harness harness = new();
        harness.Paste.Observation = SuccessfulPasteObservation(3) with { Count = 99 };

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(
            "PowerPoint pasted shapes that did not match the chart. The pasted shapes were removed; try again.",
            outcome.Message);
        Assert.Equal(["paste", "delete"], harness.Paste.Calls);
    }

    /// <summary>
    /// R7.5's <c>MapCopyResult</c> exists for this command: a clipboard-unavailable copy
    /// refusal presents the transfer taxonomy's line, not an Excel-side sentence.
    /// </summary>
    [Fact]
    public void Run_maps_a_clipboard_unavailable_copy_refusal_through_the_transfer_taxonomy()
    {
        Harness harness = new()
        {
            NextLifecycle = new StagingLifecycleOutcome(
                null,
                null,
                new StagingCopyResult(false, "detail", StagedCopyRefusal.ClipboardUnavailable),
                StagingLifecycleRefusal.CopyRefused,
                []),
        };

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.Equal(SendToPowerPointRefusal.CopyRefused, outcome.Refusal);
        Assert.Equal(
            "The clipboard is busy or unavailable. Close the application holding the clipboard and try again.",
            outcome.Message);
    }

    [Fact]
    public void Run_does_not_report_success_when_the_transfer_seam_never_ran()
    {
        Harness harness = new()
        {
            NextLifecycle = new StagingLifecycleOutcome(null, null, null, null, []),
        };

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(SendToPowerPointRefusal.NothingTransferred, outcome.Refusal);
        Assert.Empty(harness.Successes);
    }

    [Fact]
    public void Run_refuses_when_the_preflight_refuses()
    {
        Harness harness = new()
        {
            ReadOutcome = GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing),
        };

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(SendToPowerPointRefusal.TableReadRefused, outcome.Refusal);
        Assert.Contains("Initialise sheet", Assert.Single(harness.Failures), StringComparison.Ordinal);
        Assert.Equal(["read"], harness.Order);
    }

    /// <summary>The post-resolution guard refuses a zero-extent chart before any host call.</summary>
    [Fact]
    public void Run_refuses_a_zero_extent_resolution_before_the_lifecycle()
    {
        Harness harness = new()
        {
            NextBounds = new ExportBoundsOutcome(
                new ExportBounds(new RectD(0, 0, 0, 50)),
                null,
                null),
        };

        SendToPowerPointOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(SendToPowerPointRefusal.BoundsRefused, outcome.Refusal);
        Assert.Contains("no area to send", Assert.Single(harness.Failures), StringComparison.Ordinal);
        Assert.DoesNotContain("lifecycle", harness.Order);
    }

    [Fact]
    public void Run_never_lets_a_throwing_presenter_escape()
    {
        Harness harness = new();

        SendToPowerPointOutcome outcome = SendToPowerPointCommand.Run(
            harness.Build(),
            _ => throw new InvalidOperationException("presenter failed"),
            _ => { });

        Assert.True(outcome.Succeeded, outcome.Message);
    }

    [Theory]
    [InlineData("Lifecycle", nameof(SendToPowerPointRefusal.MissingLifecycle))]
    [InlineData("StateScope", nameof(SendToPowerPointRefusal.MissingStateScope))]
    [InlineData("StagingWriter", nameof(SendToPowerPointRefusal.MissingStagingWriter))]
    [InlineData("GroupingPort", nameof(SendToPowerPointRefusal.MissingGroupingPort))]
    [InlineData("CopyPort", nameof(SendToPowerPointRefusal.MissingCopyPort))]
    [InlineData("ClipboardReader", nameof(SendToPowerPointRefusal.MissingClipboardReader))]
    [InlineData("PowerPointApplication", nameof(SendToPowerPointRefusal.MissingPowerPointApplication))]
    [InlineData("PowerPointPaste", nameof(SendToPowerPointRefusal.MissingPowerPointPaste))]
    [InlineData("PowerPointPlacement", nameof(SendToPowerPointRefusal.MissingPowerPointPlacement))]
    [InlineData("TableReader", nameof(SendToPowerPointRefusal.MissingTableReader))]
    [InlineData("ConfigReader", nameof(SendToPowerPointRefusal.MissingConfigReader))]
    [InlineData("GridMeasurement", nameof(SendToPowerPointRefusal.MissingGridMeasurement))]
    [InlineData("ExportRequestFactory", nameof(SendToPowerPointRefusal.MissingExportRequestFactory))]
    [InlineData("BoundsResolver", nameof(SendToPowerPointRefusal.MissingBoundsResolver))]
    public void Run_refuses_each_missing_dependency(string member, string expectedRefusal)
    {
        SendToPowerPointRefusal expected = Enum.Parse<SendToPowerPointRefusal>(expectedRefusal);
        Harness harness = new();
        SendToPowerPointRequest request = member switch
        {
            "Lifecycle" => harness.Build() with { Lifecycle = null },
            "StateScope" => harness.Build() with { StateScope = null },
            "StagingWriter" => harness.Build() with { StagingWriter = null },
            "GroupingPort" => harness.Build() with { GroupingPort = null },
            "CopyPort" => harness.Build() with { CopyPort = null },
            "ClipboardReader" => harness.Build() with { ClipboardReader = null },
            "PowerPointApplication" => harness.Build() with { PowerPointApplication = null },
            "PowerPointPaste" => harness.Build() with { PowerPointPaste = null },
            "PowerPointPlacement" => harness.Build() with { PowerPointPlacement = null },
            "TableReader" => harness.Build() with { TableReader = null },
            "ConfigReader" => harness.Build() with { ConfigReader = null },
            "GridMeasurement" => harness.Build() with { GridMeasurement = null },
            "ExportRequestFactory" => harness.Build() with { ExportRequestFactory = null },
            "BoundsResolver" => harness.Build() with { BoundsResolver = null },
            _ => throw new ArgumentOutOfRangeException(nameof(member), member, "Unknown dependency."),
        };

        SendToPowerPointOutcome outcome = SendToPowerPointCommand.Run(
            request,
            harness.Successes.Add,
            harness.Failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(expected, outcome.Refusal);
        Assert.Equal(SendToPowerPointCommand.ConstructionMessage, Assert.Single(harness.Failures));
        Assert.Empty(harness.Successes);
    }

    [Fact]
    public void Run_refuses_a_null_request()
    {
        List<string> failures = [];

        SendToPowerPointOutcome outcome = SendToPowerPointCommand.Run(null, _ => { }, failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SendToPowerPointRefusal.NullRequest, outcome.Refusal);
        Assert.Equal(SendToPowerPointCommand.ConstructionMessage, Assert.Single(failures));
    }

    [Fact]
    public void SuccessMessage_announces_a_created_presentation_and_the_dropped_rows()
    {
        string created = SendToPowerPointCommand.SuccessMessage(2, 0, true);
        string dropped = SendToPowerPointCommand.SuccessMessage(2, 1, false);

        Assert.Equal(
            "Sent 2 shapes to the active slide. PowerPoint was not open, so a new presentation was created.",
            created);
        Assert.Equal(
            "Sent 2 shapes to the active slide. 1 row(s) with errors were not sent — fix them and send again.",
            dropped);
    }

    private static PowerPointTargetObservation TargetObservation(
        int presentationCount = 1,
        bool readOnly = false) =>
        new(
            ReadSucceeded: true,
            HResult: null,
            PresentationCount: presentationCount,
            ProtectedViewWindowCount: 0,
            ActiveProtectedViewWindowPresent: false,
            ActivePresentationReadOnly: readOnly,
            ActiveWindowPresent: true,
            ActiveViewType: 9,
            SlideShowActive: false,
            SlideCount: 1,
            ActiveSlideIndex: 1);

    private static PowerPointPasteObservation SuccessfulPasteObservation(int childCount) =>
        new(
            PasteSucceeded: true,
            HResult: null,
            Count: 1 + childCount,
            GroupCount: 1,
            GroupShapeType: 6,
            GroupChildCount: childCount,
            GroupName: "GanttCreatorStaging:export-group",
            GroupId: 42,
            WidthPt: 100,
            HeightPt: 50);

    private static GanttRowDto ValidSpan(int rowNumber) =>
        new(
            rowNumber,
            GanttRowId.New().Value,
            null,
            null,
            "As-Planned Activity",
            "Build",
            _start,
            _finish,
            null,
            null,
            null,
            null,
            null,
            null,
            null);

    private static PanelCellGrid Grid() =>
        PanelCellGrid.TryCreate(
            [new PanelColumn("Type", 60.0)],
            [15.0],
            15.0,
            null).Grid!;

    private static ExportBoundsOutcome OkBounds() =>
        new(new ExportBounds(new RectD(0, 0, 100, 50)), null, null);

    /// <summary>
    /// The deterministic fakes and captured observations one fact drives. The lifecycle fake
    /// invokes the command's transfer hook exactly as the landed lifecycle does after a
    /// verified copy, so the real R7.2/R7.3/R7.4 stages and the R7.5 mapper run host-free.
    /// </summary>
    private sealed class Harness
    {
        public Harness()
        {
            Application = new(Order.Add);
            Paste = new(Order.Add);
            Placement = new(Order.Add);
        }

        public List<string> Order { get; } = [];

        public List<string> Successes { get; } = [];

        public List<string> Failures { get; } = [];

        public List<string> Records { get; } = [];

        public IReadOnlyList<string> StagedIds { get; set; } =
            ["panel-cell:0", "panel-cell:1", "panel-cell:2"];

        public IReadOnlyList<GanttRowDto> Rows { get; set; } = [ValidSpan(1)];

        public IReadOnlyDictionary<string, string> Settings { get; set; } =
            new Dictionary<string, string>();

        public GanttTableReadOutcome? ReadOutcome { get; set; }

        public ConfigReadOutcome? ConfigOutcome { get; set; }

        public PanelGridOutcome? GridOutcome { get; set; }

        public SceneBuildRequestOutcome NextExport { get; set; } =
            SceneBuildRequestOutcome.Ok(new SceneBuildRequest());

        public ExportBoundsOutcome NextBounds { get; set; } = OkBounds();

        public StagingLifecycleOutcome? NextLifecycle { get; set; }

        public FakePowerPointApplication Application { get; }

        public FakePowerPointPaste Paste { get; }

        public FakePowerPointPlacement Placement { get; }

        public Mock<IApplicationStateScope> Scope { get; } = new();

        public Mock<IShapeWritePort> Writer { get; } = new();

        public Mock<IShapeGroupingPort> Grouping { get; } = new();

        public Mock<IShapeCopyPort> Copier { get; } = new();

        public Mock<IClipboardFormatReader> Clipboard { get; } = new();

        public StagingLifecycleRequest? SeenLifecycleRequest { get; private set; }

        public SendToPowerPointRequest Build()
        {
            return new SendToPowerPointRequest(
                new FakeTableReader(this),
                new FakeConfigReader(this),
                new FakeMeasurement(this),
                (events, settings, registry, grid, metrics) =>
                {
                    Order.Add("export");
                    return NextExport;
                },
                (request, panelTheme, selection) =>
                {
                    Order.Add("resolve");
                    return NextBounds;
                },
                RunLifecycle,
                Scope.Object,
                Writer.Object,
                Grouping.Object,
                Copier.Object,
                Clipboard.Object,
                Application,
                Paste,
                Placement,
                Records.Add);
        }

        public SendToPowerPointOutcome Run() =>
            SendToPowerPointCommand.Run(Build(), Successes.Add, Failures.Add);

        private StagingLifecycleOutcome RunLifecycle(StagingLifecycleRequest? request)
        {
            Order.Add("lifecycle");
            SeenLifecycleRequest = request;

            if (NextLifecycle is not null)
            {
                return NextLifecycle;
            }

            StagedGroupOutcome group = new(
                "GanttCreatorStaging:export-group",
                StagedIds,
                new RectD(0, 0, 100, 50),
                null,
                null);
            StagingCopyResult copy = StagingCopyResult.Success;
            StagingTransferResult? transfer = request?.Transfer?.Invoke(group, copy);
            StagingLifecycleRefusal? refusal = transfer is { Succeeded: false }
                ? StagingLifecycleRefusal.TransferRefused
                : null;
            StagingCompositionOutcome composition = new(
                StagedIds,
                new RectD(0, 0, 100, 50),
                null,
                null,
                null);

            return new StagingLifecycleOutcome(composition, group, copy, refusal, [], transfer);
        }

        private sealed class FakeTableReader(Harness harness) : IGanttTableReader
        {
            public GanttTableReadOutcome Read()
            {
                harness.Order.Add("read");
                return harness.ReadOutcome ?? GanttTableReadOutcome.Ok(harness.Rows);
            }

            public GanttTableReadOutcome Read(IReadOnlyList<int> rowNumbers) => Read();
        }

        private sealed class FakeConfigReader(Harness harness) : IConfigCatalogueReader
        {
            public ConfigReadOutcome Read()
            {
                harness.Order.Add("config");
                return harness.ConfigOutcome ?? ConfigReadOutcome.Ok(
                    "workbook-id",
                    harness.Settings,
                    GanttStyleRegistry.Empty,
                    new Dictionary<string, double>());
            }
        }

        private sealed class FakeMeasurement(Harness harness) : IPanelGridMeasurementPort
        {
            public PanelGridOutcome Measure(
                IReadOnlyList<string> columns,
                IReadOnlyList<double>? rowHeights = null)
            {
                harness.Order.Add("measure");
                return harness.GridOutcome ?? PanelGridOutcome.Ok(Grid());
            }
        }
    }

    private sealed class FakePowerPointApplication(Action<string> record) : IPowerPointApplicationPort
    {
        public PowerPointApplicationFacts ApplicationFacts { get; set; } =
            new(true, false, 123, 1, "16.0", null);

        public PowerPointTargetObservation TargetObservation { get; set; } = TargetObservation();

        public PowerPointPresentationCreation Creation { get; set; } = new(true, 1, null);

        public List<string> Calls { get; } = [];

        public PowerPointApplicationFacts AttachOrStart()
        {
            Calls.Add("attach");
            record("target");
            return ApplicationFacts;
        }

        public PowerPointPresentationFacts ReadPresentationFacts() =>
            new(true, null, true, false, true, 1, 9, false, 1);

        public PowerPointTargetObservation ReadTargetObservation()
        {
            Calls.Add("observe");
            return TargetObservation;
        }

        public PowerPointPresentationCreation CreatePresentationWithBlankSlide()
        {
            Calls.Add("create");
            return Creation;
        }
    }

    private sealed class FakePowerPointPaste(Action<string> record) : IPowerPointPastePort
    {
        public PowerPointPasteObservation Observation { get; set; } = SuccessfulPasteObservation(3);

        public PowerPointGroupDeleteResult DeleteResult { get; set; } =
            new(PowerPointGroupDeleteOutcome.Deleted, null);

        public List<string> Calls { get; } = [];

        public int? LastDeletedGroupId { get; private set; }

        public PowerPointPasteObservation PasteSpecialShape()
        {
            Calls.Add("paste");
            record("paste");
            return Observation;
        }

        public PowerPointGroupDeleteResult DeletePastedGroup(int groupId)
        {
            Calls.Add("delete");
            LastDeletedGroupId = groupId;
            return DeleteResult;
        }
    }

    private sealed class FakePowerPointPlacement(Action<string> record) : IPowerPointPlacementPort
    {
        public PowerPointSlideSizeFacts SlideSize { get; set; } = new(true, null, 960, 540);

        public PowerPointPlacementObservation? NextObservation { get; set; }

        public List<string> Calls { get; } = [];

        public PowerPointSlideSizeFacts ReadSlideSize()
        {
            Calls.Add("read-size");
            record("read-size");
            return SlideSize;
        }

        public PowerPointPlacementObservation ApplyPlacement(int groupId, SlidePlacement placement)
        {
            Calls.Add("place");
            record("place");
            return NextObservation
                ?? new PowerPointPlacementObservation(
                    PowerPointPlacementApplyOutcome.Applied,
                    null,
                    placement.Left,
                    placement.Top,
                    placement.Width,
                    placement.Height);
        }
    }
}
