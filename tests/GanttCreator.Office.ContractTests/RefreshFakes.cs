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

    /// <summary>
    /// How many Duration cells the fake writer reports as written.
    /// </summary>
    /// <remarks>
    /// Settable so a test can assert that the orchestrator passes the writer's own
    /// count through to the outcome. Deriving it from the plan instead would make an
    /// assertion about pass-through indistinguishable from an assertion about the
    /// planner.
    /// </remarks>
    public int DurationCellsWritten { get; set; }

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

    /// <summary>
    /// The header-row target the orchestrator last passed (ADR-0030 D5), or null if
    /// never called. The header row is the period band's row, so this height is what
    /// makes the period band's bottom coincide with the first body row's top.
    /// </summary>
    public double? LastHeaderHeightPt { get; set; }

    /// <summary>
    /// The reserved-row target the orchestrator last passed (ADR-0030 D4), or null if
    /// never called. That row carries the table title and the year band.
    /// </summary>
    public double? LastReservedRowHeightPt { get; set; }

    /// <summary>
    /// The chart-padding-row target the orchestrator last passed (ADR-0031 D2), or
    /// null if never called. These rows are the chart's top and bottom margin, so
    /// this figure is the height the user's chart frame ends up with.
    /// </summary>
    public double? LastPaddingRowHeightPt { get; set; }

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

    /// <summary>
    /// The panel measurement fake.
    /// </summary>
    /// <remarks>
    /// Cached rather than constructed per access. The fake records the columns it was
    /// asked to measure, so the orchestrator and a test asserting on that record must
    /// share one instance; a fresh instance per property read would leave the test
    /// looking at an object the orchestrator never called.
    /// </remarks>
    public IPanelGridMeasurementPort Panel => PanelFake;

    /// <summary>
    /// The columns the last measurement was asked for, so a test can assert the set
    /// that reaches the port.
    /// </summary>
    public IReadOnlyList<string> LastMeasuredColumns => PanelFake.LastMeasuredColumns;

    /// <summary>
    /// The concrete panel fake, for the tests that assert on what it recorded.
    /// </summary>
    private FakePanelMeasurement PanelFake => _panel ??= new FakePanelMeasurement(this);

    private FakePanelMeasurement? _panel;

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
        /// <summary>
        /// The columns the orchestrator last asked to measure, so a test can assert
        /// which set reaches the port rather than only that the port was reached.
        /// </summary>
        public IReadOnlyList<string> LastMeasuredColumns { get; private set; } = [];

        public PanelGridOutcome Measure(IReadOnlyList<string> includedColumns)
        {
            owner.Steps.Add("Measure");
            LastMeasuredColumns = includedColumns;
            return owner.MeasurementRefused
                ? PanelGridOutcome.Refused(PanelGridRefusalReason.TableMissing)
                : MeasureLikeExcel(includedColumns);
        }

        /// <summary>
        /// Reproduces what the live host reports, so this fake can fail for the same
        /// real reason production did.
        /// </summary>
        /// <remarks>
        /// <para>
        /// A HIDDEN column measures <c>0</c> — that is Excel's behaviour, and it is the
        /// whole defect. The old fake ignored <paramref name="includedColumns"/> and
        /// returned a hardcoded grid containing <c>Id</c>, so it passed a column set
        /// that the live host refuses and the bug stayed invisible to every
        /// orchestrator test.
        /// </para>
        /// <para>
        /// Hidden columns are therefore given width <c>0</c> and the grid is built
        /// through the real <see cref="PanelCellGrid.TryCreate"/>, which refuses a
        /// non-positive width exactly as it does in production.
        /// </para>
        /// </remarks>
        private static PanelGridOutcome MeasureLikeExcel(IReadOnlyList<string> includedColumns)
        {
            HashSet<string> hidden = new(
                GanttTableSchema.Default.Columns.Where(c => c.IsHidden).Select(c => c.Name),
                StringComparer.Ordinal);

            List<PanelColumn> columns =
            [
                .. includedColumns.Select(name =>
                    new PanelColumn(name, hidden.Contains(name) ? 0d : 24d)),
            ];

            // Total width of 120pt across the visible columns keeps the plot resolver
            // from refusing for insufficient width, so these tests exercise composition
            // rather than the shortfall path.
            PanelCellGridCreationOutcome created = PanelCellGrid.TryCreate(
                columns,
                [15.0],
                15,
                includedColumns);

            return created.Succeeded && created.Grid is not null
                ? PanelGridOutcome.Ok(created.Grid)
                : PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
        }
    }

    private sealed class FakeDurationWriter(RefreshFakes owner) : IDurationWritePort
    {
        public DurationWriteOutcome Write(DurationWritePlan plan)
        {
            owner.Steps.Add("Duration");
            return owner.DurationRefused
                ? DurationWriteOutcome.Refused(DurationWriteRefusalReason.WriteFailed)

                // CellsWritten is the orchestrator-reported count (first parameter);
                // writeCount is the planner's own. Injected so a test can prove the
                // orchestrator passes the former through.
                : DurationWriteOutcome.Ok(owner.DurationCellsWritten, plan.WriteCount);
        }
    }

    private sealed class FakeRowHeightNormaliser(RefreshFakes owner) : IRowHeightNormalisationPort
    {
        public RowHeightNormalisationOutcome Normalise(
            double managed,
            double splitter,
            double spacer,
            double header,
            double reservedRow,
            double paddingRow)
        {
            owner.Steps.Add("RowHeights");

            // The targets are recorded, not discarded. This fake previously
            // ignored all three arguments and returned Ok(0), so it could not fail on
            // a wrong value: the orchestrator's hardcoded (15, 6, 6) passed every
            // refresh test while contradicting the catalogue's 18 / 18 / 9. A fake
            // that discards the value under test is not a test double, it is a
            // rubber stamp. The layout-row targets are recorded for the same
            // reason: the header height is what makes the period band's bottom meet
            // the first body row's top (ADR-0030 D5), and the padding-row height is
            // the chart's top and bottom margin (ADR-0031 D2).
            owner.LastManagedHeightPt = managed;
            owner.LastSplitterHeightPt = splitter;
            owner.LastSpacerHeightPt = spacer;
            owner.LastHeaderHeightPt = header;
            owner.LastReservedRowHeightPt = reservedRow;
            owner.LastPaddingRowHeightPt = paddingRow;

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
