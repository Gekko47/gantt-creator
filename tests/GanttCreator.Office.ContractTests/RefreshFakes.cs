using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// The fakes the R4.8A orchestrator tests run against: one per collaborator, each
/// recording the order it was called in and able to refuse at a chosen step.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why each fake is a separate type rather than one configurable double.</b> D3 is
/// a claim about <em>which</em> step fails and what has been mutated by then. A
/// single "fail everything" fake could answer "does a failure prevent mutation" but
/// not "does a failure in the <em>scene build</em> prevent mutation", and the second
/// is the one that matters: the scene build sits after the worksheet writes and
/// before the first shape write, so it is the step where a careless reordering
/// would first become visible.
/// </para>
/// <para>
/// Every fake appends to one shared <c>Steps</c> list, which is what makes the
/// ordering assertion possible without any fake knowing about the others.
/// </para>
/// </remarks>
internal sealed class RefreshFakes
{
    /// <summary>The shared ordered call log every fake appends to.</summary>
    public List<string> Steps { get; } = [];

    /// <summary>The rows the table reader returns.</summary>
    public IReadOnlyList<GanttRowDto> Rows { get; set; } = [];

    /// <summary>Whether the table read refuses.</summary>
    public bool TableReadRefused { get; set; }

    /// <summary>Whether the configuration read refuses.</summary>
    public bool ConfigReadRefused { get; set; }

    /// <summary>Whether the protection guard reports a protected target.</summary>
    public ProtectionGuardOutcome Protection { get; set; } = ProtectionGuardOutcome.NotProtected;

    /// <summary>Whether the panel measurement refuses.</summary>
    public bool MeasurementRefused { get; set; }

    /// <summary>Whether the scene-request factory refuses.</summary>
    public bool FactoryRefused { get; set; }

    /// <summary>Whether the duration write refuses.</summary>
    public bool DurationRefused { get; set; }

    /// <summary>Whether the outline write refuses.</summary>
    public bool OutlineRefused { get; set; }

    /// <summary>Whether the row-height normalisation refuses.</summary>
    public bool RowHeightRefused { get; set; }

    /// <summary>The managed-row target the orchestrator last passed, or null if never called.</summary>
    public double? LastManagedHeightPt { get; set; }

    /// <summary>The splitter-row target the orchestrator last passed, or null if never called.</summary>
    public double? LastSplitterHeightPt { get; set; }

    /// <summary>The spacer-row target the orchestrator last passed, or null if never called.</summary>
    public double? LastSpacerHeightPt { get; set; }

    /// <summary>Whether the reconciliation refuses on its first operation.</summary>
    public bool ReconcileRefused { get; set; }

    /// <summary>The shape port, so a test can assert on its recorded calls.</summary>
    public FakeShapeWritePort Shapes { get; } = new();

    /// <summary>The table reader fake.</summary>
    public IGanttTableReader TableReader => new FakeTableReader(this);

    /// <summary>The configuration reader fake.</summary>
    public IConfigCatalogueReader ConfigReader => new FakeConfigReader(this);

    /// <summary>The protection guard fake.</summary>
    public IWorksheetProtectionGuard Guard => new FakeGuard(this);

    /// <summary>The panel measurement fake.</summary>
    public IPanelGridMeasurementPort Panel => new FakePanelMeasurement(this);

    /// <summary>The duration writer fake.</summary>
    public IDurationWritePort Duration => new FakeDurationWriter(this);

    /// <summary>The row-height normaliser fake.</summary>
    public IRowHeightNormalisationPort RowHeights => new FakeRowHeightNormaliser(this);

    /// <summary>The outline writer fake.</summary>
    public IOutlineGroupPort Outline => new FakeOutlineWriter(this);

    /// <summary>The scene-request factory fake.</summary>
    public ISceneBuildRequestFactory Factory => new FakeFactory(this);

    /// <summary>The managed-column classification restorer fake.</summary>
    public IColumnPresentationPort ColumnPresentation => new FakeColumnPresentation(this);

    /// <summary>The row-identity repairer fake.</summary>
    public IGanttRowIdentityRepairer Identity => new FakeIdentityRepairer(this);

    /// <summary>Whether the column-classification restore refuses.</summary>
    public bool ColumnPresentationRefused { get; set; }

    /// <summary>Whether the row-identity repair refuses.</summary>
    public bool IdentityRefused { get; set; }

    /// <summary>Builds the orchestrator over these fakes.</summary>
    /// <returns>The composed orchestrator.</returns>
    public GanttRefreshOrchestrator BuildOrchestrator() =>
        new(
            TableReader,
            ConfigReader,
            Guard,
            Panel,
            Duration,
            RowHeights,
            Outline,
            Factory,
            Shapes,
            ColumnPresentation,
            Identity);

    /// <summary>Builds the orchestrator with a specific application-state scope.</summary>
    /// <param name="scope">The scope to record against.</param>
    /// <returns>The composed orchestrator.</returns>
    public GanttRefreshOrchestrator BuildOrchestrator(IApplicationStateScope scope) =>
        new(
            TableReader,
            ConfigReader,
            Guard,
            Panel,
            Duration,
            RowHeights,
            Outline,
            Factory,
            Shapes,
            ColumnPresentation,
            Identity,
            scope);

    private sealed class FakeTableReader(RefreshFakes owner) : IGanttTableReader
    {
        public GanttTableReadOutcome Read()
        {
            owner.Steps.Add("Read");
            return owner.TableReadRefused
                ? GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing)
                : GanttTableReadOutcome.Ok(owner.Rows);
        }
    }

    private sealed class FakeConfigReader(RefreshFakes owner) : IConfigCatalogueReader
    {
        public ConfigReadOutcome Read()
        {
            owner.Steps.Add("Config");
            return owner.ConfigReadRefused
                ? ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing)

                // A registry with one real style. An EMPTY registry is refused by the
                // scene builder as UnresolvableStyle, which is correct — but it means a
                // test using one would never reach the reconciliation, and the D3 and
                // ordering claims are about what happens after composition.
                : ConfigReadOutcome.Ok(
                    "wb",
                    new Dictionary<string, string>(),
                    new GanttStyleRegistry(
                        [
                            new GanttStyleDefinition(
                                "AsPlannedActivity",
                                new HashSet<GanttLabelPosition>
                                {
                                    GanttLabelPosition.Inside,
                                    GanttLabelPosition.DataPanelLeft,
                                },
                                EntityColourCapability.Fill | EntityColourCapability.Stroke,
                                DefaultLabelPosition: GanttLabelPosition.Inside,
                                FillColour: "#DDEBF7",
                                StrokeColour: "#2E75B6",
                                TextColour: "#000000",

                                // HatchPattern is given explicitly as None. Its
                                // parameter type is a nullable enum, so leaving it
                                // defaulted arrives as a null the formatter validator
                                // refuses — a null hatch is not the same claim as "no
                                // hatch", which is what None means.
                                HatchPattern: GanttHatchPattern.None,
                                StandardOutlinePt: 1,
                                ActivityHeightPt: 12),
                        ]));
        }
    }

    private sealed class FakeGuard(RefreshFakes owner) : IWorksheetProtectionGuard
    {
        public ProtectionGuardOutcome Query()
        {
            owner.Steps.Add("Guard");
            return owner.Protection;
        }

        public ProtectionGuardOutcome QueryTarget(object? worksheet) => Query();
    }

    private sealed class FakePanelMeasurement(RefreshFakes owner) : IPanelGridMeasurementPort
    {
        public PanelGridOutcome Measure(IReadOnlyList<string> includedColumns)
        {
            owner.Steps.Add("Measure");
            return owner.MeasurementRefused
                ? PanelGridOutcome.Refused(PanelGridRefusalReason.TableMissing)
                : PanelGridOutcome.Ok(
                    // A realistic text-panel width. The plot resolver refuses a panel
                    // that leaves too little paper for a readable plot, so a fixture
                    // with a toy 80pt panel would exercise the refusal path instead of
                    // the composition this pipeline is about. Realistic here means
                    // "narrow enough to be believable, wide enough for A4 landscape to
                    // leave a readable plot".
                    PanelCellGrid.TryCreate(
                        [new PanelColumn("Id", 120)],
                        [15.0],
                        15,
                        ["Id"]).Grid!);
        }
    }

    private sealed class FakeDurationWriter(RefreshFakes owner) : IDurationWritePort
    {
        public DurationWriteOutcome Write(DurationWritePlan plan)
        {
            owner.Steps.Add("Duration");
            return owner.DurationRefused
                ? DurationWriteOutcome.Refused(DurationWriteRefusalReason.WriteFailed)
                : DurationWriteOutcome.Ok(plan.WriteCount, plan.WriteCount);
        }
    }

    private sealed class FakeRowHeightNormaliser(RefreshFakes owner) : IRowHeightNormalisationPort
    {
        public RowHeightNormalisationOutcome Normalise(double managed, double splitter, double spacer)
        {
            owner.Steps.Add("RowHeights");

            // The three targets are recorded, not discarded. This fake previously
            // ignored all three arguments and returned Ok(0), so it could not fail on
            // a wrong value: the orchestrator's hardcoded (15, 6, 6) passed every
            // refresh test while contradicting the catalogue's 18 / 18 / 9. A fake
            // that discards the value under test is not a test double, it is a
            // rubber stamp.
            owner.LastManagedHeightPt = managed;
            owner.LastSplitterHeightPt = splitter;
            owner.LastSpacerHeightPt = spacer;

            return owner.RowHeightRefused
                ? RowHeightNormalisationOutcome.Refused(RowHeightNormalisationRefusalReason.TargetProtected)
                : RowHeightNormalisationOutcome.Ok(0);
        }
    }

    private sealed class FakeColumnPresentation(RefreshFakes owner) : IColumnPresentationPort
    {
        public ColumnPresentationOutcome EnsureClassification()
        {
            owner.Steps.Add("Columns");
            return owner.ColumnPresentationRefused
                ? ColumnPresentationOutcome.Refused(ColumnPresentationRefusalReason.TargetProtected)
                : ColumnPresentationOutcome.Ok(0);
        }
    }

    private sealed class FakeIdentityRepairer(RefreshFakes owner) : IGanttRowIdentityRepairer
    {
        public GanttRowIdentityRepairOutcome Repair()
        {
            owner.Steps.Add("Identity");
            return owner.IdentityRefused
                ? GanttRowIdentityRepairOutcome.Refused(GanttRowIdentityRepairRefusalReason.WriteFailed)
                : GanttRowIdentityRepairOutcome.Ok(0);
        }
    }

    private sealed class FakeOutlineWriter(RefreshFakes owner) : IOutlineGroupPort
    {
        public OutlineGroupOutcome Apply(IReadOnlyList<GanttEvent> events)
        {
            owner.Steps.Add("Outline");
            return owner.OutlineRefused
                ? OutlineGroupOutcome.Refused(OutlineGroupRefusalReason.TargetProtected)
                : OutlineGroupOutcome.Ok(0, 0);
        }
    }

    /// <summary>
    /// A factory that returns the REAL <see cref="ExcelSceneBuildRequestFactory"/>,
    /// because a stubbed request would let the orchestrator's later steps be tested
    /// against a scene that could never be built.
    /// </summary>
    private sealed class FakeFactory(RefreshFakes owner) : ISceneBuildRequestFactory
    {
        public SceneBuildRequestOutcome Create(
            IReadOnlyList<GanttEvent> events,
            IReadOnlyDictionary<string, string> settings,
            GanttStyleRegistry registry,
            PanelCellGrid grid)
        {
            owner.Steps.Add("Factory");
            if (owner.FactoryRefused)
            {
                return SceneBuildRequestOutcome.Refused(
                    SceneBuildRequestRefusal.InvalidSetting,
                    "injected");
            }

            // The real factory, so the request the orchestrator builds from is the one
            // production would build from. Core already publishes a deterministic
            // text-metrics implementation for hosts with no real font yet, so the
            // composition is identical on every run without a second fake here.
            return new ExcelSceneBuildRequestFactory(new FakeTextMetrics())
                .Create(events, settings, registry, grid);
        }
    }
}
