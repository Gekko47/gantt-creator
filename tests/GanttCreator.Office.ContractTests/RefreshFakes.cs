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

    /// <summary>
    /// The refusal the FIRST table read reports, or <see langword="null"/> for
    /// <see cref="GanttTableReadRefusalReason.TableMissing"/>.
    /// </summary>
    /// <remarks>
    /// The first read has its own refusal mapping in the orchestrator — an unsupported
    /// 1904 workbook is a table-read refusal with its own message, not a missing table —
    /// and a fake that could only produce one reason made every other message
    /// unreachable.
    /// </remarks>
    public GanttTableReadRefusalReason? FirstReadRefusal { get; set; }

    /// <summary>Whether the configuration read refuses.</summary>
    public bool ConfigReadRefused { get; set; }

    /// <summary>
    /// The metrics the configuration read returns, so a test can prove the
    /// workbook's own values reach the row-height normalisation.
    /// </summary>
    /// <remarks>
    /// The Theme dialog writes the user's row, splitter and spacer heights to the
    /// metrics table, and the orchestrator used to normalise the sheet to the
    /// token catalogue's DEFAULTS instead - so the setting was stored and then
    /// ignored on every Refresh. A fake that returned an empty map could not
    /// tell that apart from a workbook that never carried the metric.
    /// </remarks>
    public IReadOnlyDictionary<string, double> Metrics { get; set; } = new Dictionary<string, double>();

    /// <summary>
    /// The settings the configuration read returns, so a test can supply the
    /// chart title and its visibility.
    /// </summary>
    public IReadOnlyDictionary<string, string> Settings { get; set; } = new Dictionary<string, string>();

    /// <summary>
    /// The grid the panel measurement returns, or null for the Excel-like
    /// default.
    /// </summary>
    /// <remarks>
    /// The default grid models a sheet with no reserved row, so a test that
    /// needs the title to land in one injects a grid that has it.
    /// </remarks>
    public PanelCellGrid? MeasuredGrid { get; set; }

    /// <summary>Whether the protection guard reports a protected target.</summary>
    public ProtectionGuardOutcome Protection { get; set; } = ProtectionGuardOutcome.NotProtected;

    /// <summary>The protection port the refresh's arming step writes through (ADR-0055).</summary>
    public FakeWorksheetProtectionPort SheetProtection { get; } = new();

    /// <summary>
    /// The cross-refresh arming memory, shared by every orchestrator a test builds over these fakes,
    /// which is what the live composition does with a new orchestrator per refresh.
    /// </summary>
    public ProtectionArmingState SheetProtectionArming { get; } = new();

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

    /// <summary>
    /// How many outline groups the fake reports as applied, so a test can prove that a
    /// structural outline write forces a full shape update.
    /// </summary>
    public int OutlineGroupsApplied { get; set; }

    /// <summary>How many stale outline groups the fake reports as cleared.</summary>
    public int OutlineRowsUngrouped { get; set; }

    /// <summary>Whether the row-height normalisation refuses.</summary>
    public bool RowHeightRefused { get; set; }

    /// <summary>
    /// How many row heights the fake reports as written, so a test can prove that a
    /// structural height write forces a full shape update.
    /// </summary>
    public int RowHeightsWritten { get; set; }

    /// <summary>
    /// The body heights the fake normaliser reports as the ones it measured, so a test can
    /// prove the orchestrator passes them to the panel measurement.
    /// </summary>
    public IReadOnlyList<double> RowHeightBodyHeights { get; set; } = [];

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

    /// <summary>
    /// The reserved anchor-row target the orchestrator last passed (ADR-0038 D1), or
    /// null if never called. This row is the plot-spanning shapes' bottom cell anchor,
    /// so it must be a sub-row height and must NOT equal the padding-row height -- a
    /// margin-height anchor row would be a visible strip of sheet.
    /// </summary>
    public double? LastAnchorRowHeightPt { get; set; }

    /// <summary>
    /// The <c>Type</c> texts the orchestrator last handed to the height normaliser, or
    /// <see langword="null"/> when it supplied none.
    /// </summary>
    /// <remarks>
    /// Recorded rather than discarded for the same reason as the height targets: the
    /// whole point of passing the bulk-read column is that the normaliser does not read
    /// it again, and a fake that dropped the argument could not distinguish a caller
    /// that passes the texts from one that does not.
    /// </remarks>
    public IReadOnlyList<string?>? LastTypeTexts { get; set; }

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
    /// The body heights the orchestrator handed to the panel measurement, or
    /// <see langword="null"/> when it supplied none.
    /// </summary>
    /// <remarks>
    /// Surfaces the measurement fake's record so a test can prove the orchestrator
    /// reuses the heights its own normalisation just measured instead of making the
    /// measurement read every body row again.
    /// </remarks>
    public IReadOnlyList<double>? LastMeasuredBodyHeightsPt => PanelFake.LastBodyHeightsPt;

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

    /// <summary>The row/cell warning highlighter fake.</summary>
    public IRowWarningHighlighter Highlighter => new FakeRowWarningHighlighter(this);

    /// <summary>Whether the column-classification restore refuses.</summary>
    public bool ColumnPresentationRefused { get; set; }

    /// <summary>
    /// Every issue list the reporter was handed, in call order, so a test can assert
    /// what the orchestrator reported rather than only how many times it called.
    /// </summary>
    public List<IReadOnlyList<HighlightTarget>> HighlightReconciliations { get; } = [];

    /// <summary>Whether the notes reporter refuses.</summary>
    public bool WarningHighlightsRefused { get; set; }

    /// <summary>Whether the row-identity repair refuses.</summary>
    public bool IdentityRefused { get; set; }

    /// <summary>
    /// How many Id cells the fake repairer reports as written. Zero models the normal
    /// case, in which the orchestrator must not re-read the whole table.
    /// </summary>
    public int IdentityRepairedCount { get; set; }

    /// <summary>
    /// How many engine-hidden cells the fake type-change normaliser reports as written.
    /// Zero models a no-op pass.
    /// </summary>
    public int TypeCellsWritten { get; set; }

    /// <summary>
    /// The rows the table reader returns AFTER identity repair has run, or
    /// <see langword="null"/> to leave <see cref="Rows"/> unchanged.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A real repairer WRITES new Id values into the worksheet, so the table read
    /// after it must see different rows than the read before it. A fake that changed
    /// nothing let the orchestrator validate its stale step-1 snapshot and still pass,
    /// which is exactly the defect the re-read after repair fixes. Setting this
    /// models the write.
    /// </para>
    /// <para>
    /// A second read can also fail, which the production code has to handle; see
    /// <see cref="TableReadRefusedAfterRepair"/>.
    /// </para>
    /// </remarks>
    public IReadOnlyList<GanttRowDto>? RowsAfterRepair { get; set; }

    /// <summary>Whether the table re-read performed after identity repair refuses.</summary>
    public bool TableReadRefusedAfterRepair { get; set; }

    /// <summary>
    /// The refusal the fake table reader reports on every read AFTER the first.
    /// </summary>
    /// <remarks>
    /// Settable so a test can distinguish the two distinct re-read steps — the one after
    /// identity repair and the one after type-change normalisation — and can drive each
    /// to a different refusal. Both steps map a failed re-read onto the same two
    /// refusals, but only the reasons are shared; the code paths are separate, so a test
    /// that could only produce <see cref="GanttTableReadRefusalReason.TableMissing"/>
    /// would leave the no-workbook arm of the normalisation path unproven.
    /// </remarks>
    public GanttTableReadRefusalReason RereadRefusalReason { get; set; } =
        GanttTableReadRefusalReason.TableMissing;

    /// <summary>Whether the StyleKey write refuses.</summary>
    public bool StyleKeyWriteRefused { get; set; }

    /// <summary>How many StyleKey cells the fake writer reports as written.</summary>
    public int StyleKeyCellsWritten { get; set; }

    /// <summary>How many rows the last StyleKey plan asked to write.</summary>
    /// <remarks>
    /// Recorded from the plan rather than derived from the rows, so "the planner had
    /// nothing to do" and "the orchestrator skipped the write" stay distinguishable
    /// claims. Deriving it would make the assertion a restatement of the planner.
    /// </remarks>
    public int StyleKeyPlanWriteCount { get; private set; }

    /// <summary>How many times the type-change normaliser was asked to run.</summary>
    public int NormaliseCount { get; private set; }

    /// <summary>How many times the table reader was called.</summary>
    public int TableReadCount { get; private set; }

    /// <summary>How many unsafe changed-row reads were requested.</summary>
    public int TargetedTableReadCount { get; private set; }

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
            Identity,
            null,
            null,
            Highlighter,
            null);

    /// <summary>Builds the orchestrator with the ADR-0055 arming step wired.</summary>
    /// <param name="trace">The diagnostic sink, or null for no trace.</param>
    /// <returns>The composed orchestrator.</returns>
    /// <remarks>
    /// The ownership source and the arming memory are deliberately the fakes' shared instances:
    /// arming is a session fact that outlives one orchestrator, so a test that wants to see the
    /// second pass has to build a second orchestrator over the same memory.
    /// </remarks>
    public GanttRefreshOrchestrator BuildOrchestratorWithProtection(Action<string>? trace = null) =>
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
            null,
            null,
            Highlighter,
            null,
            trace: trace,
            protectionArming: new WorksheetProtectionArming(SheetProtection, Shapes, SheetProtectionArming, trace));

    /// <summary>Builds the orchestrator with a pending worksheet change batch.</summary>
    public GanttRefreshOrchestrator BuildOrchestratorWithChanges(ChangeTracker tracker) =>
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
            null,
            null,
            Highlighter,
            null,
            changeTracker: tracker);

    /// <summary>Builds an orchestrator using the injected payload-signature cache.</summary>
    /// <param name="cache">The cache to consult and commit.</param>
    /// <param name="useDeltaCache">Whether this run may skip unchanged payload writes.</param>
    /// <param name="trace">The diagnostic sink, or null for no trace.</param>
    public GanttRefreshOrchestrator BuildOrchestratorWithDeltaCache(
        ShapeDeltaCache cache,
        bool useDeltaCache = true,
        Action<string>? trace = null) =>
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
            null,
            null,
            Highlighter,
            null,
            shapeDeltaCache: cache,
            useDeltaCache: useDeltaCache,
            trace: trace);

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
            scope,
            null,
            Highlighter,
            null);

    /// <summary>
    /// Builds the orchestrator with an explicit warning highlighter, so a test can prove
    /// the dependency is genuinely optional rather than merely defaulted.
    /// </summary>
    /// <param name="warningHighlighter">
    /// The highlighter to wire, or <see langword="null"/> to omit it entirely.
    /// </param>
    /// <returns>The composed orchestrator.</returns>
    public GanttRefreshOrchestrator BuildOrchestratorWithWarningHighlighter(IRowWarningHighlighter? warningHighlighter) =>
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
            null,
            null,
            warningHighlighter,
            null);

    /// <summary>
    /// Builds the orchestrator with ONLY the type-change normaliser wired.
    /// </summary>
    /// <remarks>
    /// One optional collaborator at a time, so a test that observes an extra table read
    /// can attribute it to the normaliser rather than to any other step that also
    /// re-reads.
    /// </remarks>
    /// <returns>The composed orchestrator.</returns>
    public GanttRefreshOrchestrator BuildOrchestratorWithTypeChangeNormaliser() =>
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
            columnPresentation: null,

            // Identity repair is deliberately OMITTED. It re-reads the table too, so a
            // wired identity repairer would consume the injected re-read refusal first
            // and leave the normalisation path unproven — the test would pass through a
            // different branch than the one it names.
            identityRepairer: null,
            stateScope: null,
            styleKeyWriter: null,
            rowWarningHighlighter: Highlighter,
            typeChangeNormaliser: new FakeTypeChangeNormaliser(this));

    /// <summary>
    /// Builds the orchestrator with ONLY the StyleKey writer wired.
    /// </summary>
    /// <returns>The composed orchestrator.</returns>
    public GanttRefreshOrchestrator BuildOrchestratorWithStyleKeyWriter() =>
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
            columnPresentation: null,
            identityRepairer: null,
            stateScope: null,
            styleKeyWriter: new FakeStyleKeyWriter(this),
            rowWarningHighlighter: Highlighter,
            typeChangeNormaliser: null);

    private sealed class FakeTableReader(RefreshFakes owner) : IGanttTableReader
    {
        public GanttTableReadOutcome Read()
        {
            owner.Steps.Add("Read");
            owner.TableReadCount++;

            // The second and later reads happen after identity repair has written to
            // the worksheet, so they see the REPAIRED rows and can fail independently
            // of the first read. A first read still uses the original rows.
            if (owner.TableReadCount > 1)
            {
                if (owner.TableReadRefusedAfterRepair)
                {
                    return GanttTableReadOutcome.Refused(owner.RereadRefusalReason);
                }

                return GanttTableReadOutcome.Ok(owner.RowsAfterRepair ?? owner.Rows);
            }

            return owner.TableReadRefused
                ? GanttTableReadOutcome.Refused(
                    owner.FirstReadRefusal ?? GanttTableReadRefusalReason.TableMissing)
                : GanttTableReadOutcome.Ok(owner.Rows);
        }

        public GanttTableReadOutcome Read(IReadOnlyList<int> rowNumbers)
        {
            owner.TargetedTableReadCount++;
            return GanttTableReadOutcome.Ok([.. owner.Rows.Where(row => rowNumbers.Contains(row.RowNumber))]);
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
                    owner.Settings,
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
                        ]),
                    owner.Metrics);
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

        /// <summary>
        /// The body heights the orchestrator handed over, or <see langword="null"/> when it
        /// supplied none.
        /// </summary>
        /// <remarks>
        /// Recorded so a test can prove the orchestrator reuses the heights its own
        /// normalisation just measured instead of making this port read every row again.
        /// </remarks>
        public IReadOnlyList<double>? LastBodyHeightsPt { get; private set; }

        public PanelGridOutcome Measure(
            IReadOnlyList<string> includedColumns,
            IReadOnlyList<double>? bodyRowHeightsPt = null)
        {
            owner.Steps.Add("Measure");
            LastMeasuredColumns = includedColumns;
            LastBodyHeightsPt = bodyRowHeightsPt;
            if (owner.MeasurementRefused)
            {
                return PanelGridOutcome.Refused(PanelGridRefusalReason.TableMissing);
            }

            return owner.MeasuredGrid is { } grid
                ? PanelGridOutcome.Ok(grid)
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

            // The grid models the LIVE sheet's vertical layout, because a grid
            // without the reserved row is a sheet that cannot exist (ADR-0030
            // D4/D5): the top padding row, then the reserved row that carries the
            // table title and the year band, then the header row that IS the
            // period band's row, then the body. The old fake had a header and a
            // body but no reserved row, so the year band would have been drawn
            // over the header - and a live chart whose content starts above row 1
            // is now refused rather than silently misplaced (ADR-0047).
            //
            // Total width of 120pt across the visible columns keeps the plot resolver
            // from refusing for insufficient width, so these tests exercise composition
            // rather than the shortfall path.
            double reservedRowHeightPt = GanttCatalogues.MetricDefault("YearBandHeightPt");
            double headerHeightPt = GanttCatalogues.MetricDefault("PeriodBandHeightPt");
            PanelCellGridCreationOutcome created = PanelCellGrid.TryCreate(
                columns,
                [GanttCatalogues.MetricDefault("GanttRowHeightPt")],
                headerHeightPt,
                includedColumns,
                originTopPt: reservedRowHeightPt + headerHeightPt,
                originLeftPt: 0d,
                topPaddingHeightPt: 0d);

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
            double paddingRow,
            double anchorRow,
            IReadOnlyList<string?>? typeTexts = null)
        {
            owner.Steps.Add("RowHeights");

            // The texts are recorded, not discarded: passing the bulk-read column is the
            // whole point of this parameter, and only a fake that keeps it can prove the
            // orchestrator supplies it.
            owner.LastTypeTexts = typeTexts;

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

            // The anchor row is recorded for the same reason, and it is the one figure
            // that must differ from the padding height (ADR-0038 D1).
            owner.LastAnchorRowHeightPt = anchorRow;

            return owner.RowHeightRefused
                ? RowHeightNormalisationOutcome.Refused(RowHeightNormalisationRefusalReason.TargetProtected)
                : RowHeightNormalisationOutcome.Ok(owner.RowHeightsWritten, owner.RowHeightBodyHeights);
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
                : GanttRowIdentityRepairOutcome.Ok(owner.IdentityRepairedCount);
        }
    }

    /// <summary>A type-change normaliser that records that it ran.</summary>
    /// <remarks>
    /// The step takes no input and reports only a count, so there is nothing else to
    /// model. What matters to the orchestrator is that the write happened at all and
    /// that the rows are re-read afterwards, which is what the tests assert.
    /// </remarks>
    private sealed class FakeTypeChangeNormaliser(RefreshFakes owner) : ITypeChangeNormaliser
    {
        public GanttTypeChangeNormaliseOutcome Normalise()
        {
            owner.Steps.Add("Normalise");
            owner.NormaliseCount++;
            return new GanttTypeChangeNormaliseOutcome(owner.TypeCellsWritten);
        }
    }

    /// <summary>A StyleKey writer that records the plan it was asked to write.</summary>
    /// <remarks>
    /// <c>WriteCount</c> comes from the plan, not from a settable field, so a test can
    /// tell "the planner found nothing to write" from "the orchestrator never called the
    /// writer" without either assertion restating the other.
    /// </remarks>
    private sealed class FakeStyleKeyWriter(RefreshFakes owner) : IStyleKeyWritePort
    {
        public StyleKeyWriteOutcome Write(StyleKeyWritePlan plan)
        {
            owner.Steps.Add("StyleKey");
            owner.StyleKeyPlanWriteCount = plan.WriteCount;

            return owner.StyleKeyWriteRefused
                ? StyleKeyWriteOutcome.Refused(StyleKeyWriteRefusalReason.WriteFailed)
                : StyleKeyWriteOutcome.Ok(owner.StyleKeyCellsWritten, plan.WriteCount);
        }
    }

    private sealed class FakeRowWarningHighlighter(RefreshFakes owner) : IRowWarningHighlighter
    {
        public GanttValidationReportOutcome HighlightRow(int rowNumber, GanttValidationSeverity severity) =>
            GanttValidationReportOutcome.Ok(0);

        public GanttValidationReportOutcome HighlightCell(int rowNumber, int column, GanttValidationSeverity severity) =>
            GanttValidationReportOutcome.Ok(0);

        public GanttValidationReportOutcome ClearHighlight(int rowNumber, int? column = null) =>
            GanttValidationReportOutcome.Ok(0);

        public GanttValidationReportOutcome ReconcileHighlights(
            IReadOnlyList<HighlightTarget> currentTargets,
            IReadOnlyList<int>? rowsToReconcile = null)
        {
            owner.Steps.Add("Highlights");
            owner.HighlightReconciliations.Add(currentTargets);
            return owner.WarningHighlightsRefused
                ? GanttValidationReportOutcome.Refused(GanttValidationReportRefusalReason.TargetProtected)
                : GanttValidationReportOutcome.Ok(currentTargets.Count);
        }
    }

    private sealed class FakeOutlineWriter(RefreshFakes owner) : IOutlineGroupPort
    {
        public OutlineGroupOutcome Apply(IReadOnlyList<GanttEvent> events)
        {
            owner.Steps.Add("Outline");
            return owner.OutlineRefused
                ? OutlineGroupOutcome.Refused(OutlineGroupRefusalReason.TargetProtected)
                : OutlineGroupOutcome.Ok(owner.OutlineGroupsApplied, owner.OutlineRowsUngrouped);
        }

        public OutlineCollapseOutcome SetCollapsed(OutlineCollapseDirection direction)
        {
            owner.Steps.Add("Collapse");
            return owner.OutlineRefused
                ? OutlineCollapseOutcome.Refused(OutlineCollapseRefusalReason.TargetProtected)
                : OutlineCollapseOutcome.Ok(0);
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
            PanelCellGrid grid,
            IReadOnlyDictionary<string, double> themedMetrics)
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
                .Create(events, settings, registry, grid, themedMetrics);
        }
    }
}
