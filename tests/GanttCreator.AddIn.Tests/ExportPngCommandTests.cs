using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using GanttCreator.Raster;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// R8.8 D1: one command, the verified pipeline. The pins are the order (every
/// input-dependent refusal precedes the dialog, the write follows the render) and the
/// one-line outcomes (a cancel is silent, a refusal names one actionable fact, a success
/// states counts and the path only).
/// </summary>
/// <remarks>
/// The render and encode stages run the production R8.4/R8.5 entries against an empty
/// scene, so the pins exercise real bytes; the dialog, the write and the pre-flight's own
/// reader/measurement seams are stubs. The independent byte-level assertions on the
/// written file are the Phase-E suite's job.
/// </remarks>
public sealed class ExportPngCommandTests
{
    /// <summary>A 2:1 chart bounds; the zero-origin extent is what the pixel budget reads.</summary>
    private static readonly RectD Bounds = new(0, 0, 800, 400);

    private static GanttScene EmptyScene() =>
        GanttScene.TryCreate(Bounds, Bounds, [], []).Scene!;

    private static PanelCellGrid Grid() =>
        PanelCellGrid.TryCreate(
            [new PanelColumn("Type", 60.0)],
            [15.0],
            15.0,
            null).Grid!;

    /// <summary>Builds a pre-flight request whose every stage succeeds host-free.</summary>
    private static EditableExportPreflightRequest Preflight(
        IReadOnlyDictionary<string, string>? settings = null,
        GanttScene? scene = null,
        ConfigReadRefusalReason? configRefusal = null)
    {
        GanttScene built = scene ?? EmptyScene();
        return new EditableExportPreflightRequest(
            new StubTableReader(),
            new StubConfigReader(settings ?? new Dictionary<string, string>(), configRefusal),
            new StubMeasurement(),
            (_, _, _, _, _) => SceneBuildRequestOutcome.Ok(new SceneBuildRequest()),
            (_, _, _) => new ExportBoundsOutcome(new ExportBounds(Bounds), null, null, built));
    }

    /// <summary>A scene carrying one text primitive with the supplied (or missing) family.</summary>
    private static GanttScene SceneWithText(string? family)
    {
        SceneText text = new(
            "chart:title-text",
            SceneOwnerId.Chart,
            ZLayer.Title,
            "T",
            new RectD(10, 10, 50, 10),
            new SceneStyle("TitleStyle", fontFamily: family, fontSizePt: 9, bold: false),
            GanttTextAlignment.Centre);
        return GanttScene.TryCreate(Bounds, Bounds, [text], []).Scene!;
    }

    private static ExportPngRequest Request(
        PngExportTarget target = PngExportTarget.File,
        Action? onRender = null,
        ISaveFileDialogPort? dialog = null,
        PngClipboardPlace? clipboard = null,
        PngWrite? write = null,
        EditableExportPreflightRequest? preflight = null,
        Action<RasterBackground>? onBackground = null,
        Action<string>? technicalRecord = null,
        bool chartRenderCurrent = true,
        Action? ensureCurrent = null,
        string? initialDirectory = null,
        Action<string?>? persistFolder = null,
        PngFontGate.FamilyResolver? fontResolver = null) =>
        new(
            preflight ?? Preflight(),
            target,
            chartRenderCurrent,
            ensureCurrent,
            dialog ?? new StubDialog("out.png"),
            clipboard ?? (_ => PngClipboardOutcome.Ok()),
            initialDirectory,
            persistFolder,
            (scene, dimensions, background) =>
            {
                onBackground?.Invoke(background);
                onRender?.Invoke();
                return SceneRasterRenderer.Render(scene, dimensions, background);
            },
            PngExportEncoder.Encode,
            write ?? ((png, path) => new PngWriteOutcome(true, null, "written", "record")),
            technicalRecord,
            fontResolver ?? Resolves);

    /// <summary>
    /// A resolver that reports every family and weight as renderable, so a test that does not
    /// exercise the font gate never touches the pinned faces on disk (host-free default).
    /// </summary>
    private static PinnedFontResolution Resolves(string fontFamily, bool _) =>
        PinnedFontResolution.Resolved(
            null!,
            new PinnedFontObservation(fontFamily, fontFamily, string.Empty, double.NaN, 1000, null, string.Empty));

    [Fact]
    public void Run_writes_and_reports_the_path_and_the_exact_dimensions()
    {
        var successes = new List<string>();
        var failures = new List<string>();
        ExportPngRequest captured = Request();
        string? writtenPath = null;
        PngExportBytes? writtenBytes = null;
        ExportPngRequest request = captured with
        {
            Write = (png, path) =>
            {
                writtenPath = path;
                writtenBytes = png;
                return new PngWriteOutcome(true, null, "written", "record");
            },
        };

        ExportPngOutcome outcome = ExportPngCommand.Run(request, successes.Add, failures.Add);

        Assert.True(outcome.Succeeded);
        Assert.False(outcome.Cancelled);
        Assert.Null(outcome.Refusal);
        Assert.Equal("out.png", writtenPath);
        Assert.Equal(outcome.Dimensions, writtenBytes!.Dimensions);
        Assert.Single(successes);
        Assert.Empty(failures);
        Assert.Contains("px to out.png", successes[0], StringComparison.Ordinal);
        Assert.DoesNotContain("Build", successes[0], StringComparison.Ordinal);
    }

    /// <summary>
    /// The dimension budget is checked against the FRESH composition bounds and precedes
    /// the dialog, so an over-sized chart never makes the user choose a path first (D1/D7).
    /// </summary>
    [Fact]
    public void Run_refuses_an_over_budget_composition_before_the_dialog()
    {
        var dialog = new StubDialog("out.png");
        var failures = new List<string>();
        RectD huge = new(0, 0, 100_000, 400);
        GanttScene hugeScene = GanttScene.TryCreate(huge, huge, [], []).Scene!;
        EditableExportPreflightRequest preflight = Preflight() with
        {
            BoundsResolver = (_, _, _) =>
                new ExportBoundsOutcome(new ExportBounds(huge), null, null, hugeScene),
        };

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(preflight: preflight, dialog: dialog),
            _ => { },
            failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ExportPngRefusal.DimensionsRefused, outcome.Refusal);
        Assert.Equal(0, dialog.Calls);
        Assert.Contains("65535", failures[0], StringComparison.Ordinal);
        Assert.DoesNotContain("Parameter", failures[0], StringComparison.Ordinal);
    }

    /// <summary>A cancelled dialog is silence: nothing is rendered, written or presented.</summary>
    [Fact]
    public void Run_treats_a_cancelled_dialog_as_silence()
    {
        var rendered = 0;
        var written = 0;
        var successes = new List<string>();
        var failures = new List<string>();

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(
                dialog: new StubDialog(null),
                onRender: () => rendered++,
                write: (_, _) =>
                {
                    written++;
                    return new PngWriteOutcome(true, null, "written", "record");
                }),
            successes.Add,
            failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.True(outcome.Cancelled);
        Assert.Null(outcome.Refusal);
        Assert.Equal(0, rendered);
        Assert.Equal(0, written);
        Assert.Empty(successes);
        Assert.Empty(failures);
    }

    /// <summary>A write refusal presents the writer's one actionable line and never a success.</summary>
    [Fact]
    public void Run_reports_a_write_refusal_with_the_writers_message()
    {
        var successes = new List<string>();
        var failures = new List<string>();

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(write: (_, _) => new PngWriteOutcome(
                false,
                PngWriteRefusal.UnwritablePath,
                "The folder could not be written. Choose another folder and export again.",
                "record")),
            successes.Add,
            failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ExportPngRefusal.WriteRefused, outcome.Refusal);
        Assert.Empty(successes);
        Assert.Single(failures);
        Assert.Contains("another folder", failures[0], StringComparison.Ordinal);
    }

    /// <summary>
    /// The canvas background comes from the same settings key and catalogue fallback the
    /// scene uses, so the image's edge pixels carry the themed chart background (D3).
    /// </summary>
    [Fact]
    public void Run_passes_the_themed_chart_background_to_the_renderer()
    {
        string key = GanttChromeKey();
        RasterBackground? seen = null;

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(
                preflight: Preflight(new Dictionary<string, string> { [key] = "#123456" }),
                onBackground: background => seen = background),
            _ => { },
            _ => { });

        Assert.True(outcome.Succeeded);
        Assert.Equal(RasterBackground.FromColour(ColourHex.Parse("#123456")), seen);
    }

    /// <summary>A pre-flight refusal is reported with this command's own line and stops the run.</summary>
    [Fact]
    public void Run_reports_a_preflight_refusal_and_stops()
    {
        var dialog = new StubDialog("out.png");
        var failures = new List<string>();
        EditableExportPreflightRequest refusing = Preflight() with
        {
            TableReader = new StubTableReader(GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing)),
        };

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(preflight: refusing, dialog: dialog),
            _ => { },
            failures.Add);

        Assert.Equal(ExportPngRefusal.TableReadRefused, outcome.Refusal);
        Assert.Equal(0, dialog.Calls);
        Assert.Contains("table could not be read", failures[0], StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// A configuration the reader cannot accept is reported with the remedy, not as a dead end.
    /// The live 2026-11-24 report was a schema-27 workbook whose export said only that the
    /// configuration could not be read, so the user discovered Fix through Refresh instead; the
    /// export now names Fix itself and records the reader's reason.
    /// </summary>
    [Fact]
    public void Run_names_the_repair_remedy_when_the_configuration_cannot_be_read()
    {
        var failures = new List<string>();
        var records = new List<string>();

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(
                preflight: Preflight(configRefusal: ConfigReadRefusalReason.RowCountMismatch),
                technicalRecord: records.Add),
            _ => { },
            failures.Add);

        Assert.Equal(ExportPngRefusal.ConfigReadRefused, outcome.Refusal);
        Assert.Contains("Run Fix", failures[0], StringComparison.Ordinal);
        Assert.Contains(
            records,
            record => string.Equals(
                record,
                "gantt.export config refused=RowCountMismatch",
                StringComparison.Ordinal));
    }

    /// <summary>An absent workbook is not sent to Fix: that command cannot create one.</summary>
    [Fact]
    public void Run_does_not_send_an_absent_workbook_to_the_repair_remedy()
    {
        var failures = new List<string>();

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(preflight: Preflight(configRefusal: ConfigReadRefusalReason.NoActiveWorkbook)),
            _ => { },
            failures.Add);

        Assert.Equal(ExportPngRefusal.ConfigReadRefused, outcome.Refusal);
        Assert.DoesNotContain("Run Fix", failures[0], StringComparison.Ordinal);
        Assert.Contains("no active workbook", failures[0], StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>The stages run in the contract's order: read, dialog, render, write.</summary>
    [Fact]
    public void Run_runs_the_stages_in_the_contract_order()
    {
        var order = new List<string>();
        EditableExportPreflightRequest preflight = Preflight() with
        {
            TableReader = new StubTableReader(order: order),
        };

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(
                preflight: preflight,
                dialog: new StubDialog("out.png", () => order.Add("dialog")),
                onRender: () => order.Add("render"),
                write: (_, _) =>
                {
                    order.Add("write");
                    return new PngWriteOutcome(true, null, "written", "record");
                }),
            _ => { },
            _ => { });

        Assert.True(outcome.Succeeded);
        Assert.Equal(["read", "dialog", "render", "write"], order);
    }

    /// <summary>The background falls back to the catalogue token when the setting is absent.</summary>
    [Fact]
    public void ChartBackground_falls_back_to_the_catalogue_token()
    {
        RasterBackground background = ExportPngCommand.ChartBackground(new Dictionary<string, string>());

        Assert.Equal(
            RasterBackground.FromColour(
                ColourHex.Parse(GanttChartChrome.DefaultHexFor(GanttChartChromeRole.Background))),
            background);
    }

    /// <summary>
    /// The live defect of 2026-11-23, now a regression pin. A text primitive declared with
    /// no font family used to make the raster renderer refuse the whole export
    /// (<c>FamilyMismatch</c> with an empty family). The scene names the label family at
    /// the primitive, so the export now renders, and the trace states the family the
    /// renderer will resolve — the two facts that together prove the fix and keep the
    /// diagnosis available if another family-less style ever appears.
    /// </summary>
    [Fact]
    public void A_text_primitive_declared_without_a_family_renders_with_the_label_family()
    {
        var records = new List<string>();
        var failures = new List<string>();

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(
                preflight: Preflight(scene: SceneWithText(family: null)),
                technicalRecord: records.Add) with
            {
                // Host-free render: the pinned label family is a proprietary Office font, so a
                // hosted runner cannot load it. This test asserts the export's orchestration (the
                // primitive is rendered and the trace names the family), not the host's fonts.
                Render = (_, dimensions, _) => new RasterRenderResult(dimensions, RasterDeterminism.Pinned),
            },
            _ => { },
            failures.Add);

        Assert.True(outcome.Succeeded, failures.Count > 0 ? failures[0] : null);
        Assert.Contains(
            records,
            record => record.StartsWith("gantt.export text-families", StringComparison.Ordinal)
                && record.Contains(
                    "TitleStyle=" + GanttCatalogues.LabelFontFamily,
                    StringComparison.Ordinal));
    }

    /// <summary>
    /// A family the engine resolves adds no caveat and no font record: the success line is the
    /// plain one (the positive counterpart of the substitution arm).
    /// </summary>
    [Fact]
    public void Run_adds_no_font_caveat_when_every_family_resolves()
    {
        var successes = new List<string>();
        var records = new List<string>();

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(
                preflight: Preflight(scene: SceneWithText(family: GanttCatalogues.LabelFontFamily)),
                technicalRecord: records.Add) with
            {
                // Host-free render, for the same reason as the family-less pin above.
                Render = (_, dimensions, _) => new RasterRenderResult(dimensions, RasterDeterminism.Pinned),
            },
            successes.Add,
            _ => { });

        Assert.True(outcome.Succeeded);
        Assert.DoesNotContain("drawn with", successes[0], StringComparison.Ordinal);
        Assert.DoesNotContain(
            records,
            record => record.StartsWith("gantt.export font-fallback", StringComparison.Ordinal));
    }

    /// <summary>
    /// A family the engine cannot resolve renders with the pinned label family, and the status
    /// line and the technical record both name the substitution (D10). The render receives the
    /// substituted scene, so the production renderer never sees the unresolvable family.
    /// </summary>
    [Fact]
    public void Run_reports_a_font_substitution_in_the_success_line_and_record()
    {
        var successes = new List<string>();
        var records = new List<string>();

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(
                preflight: Preflight(scene: SceneWithText(family: "Comic Sans MS")),
                technicalRecord: records.Add,
                fontResolver: (family, _) => string.Equals(
                    family, GanttCatalogues.LabelFontFamily, StringComparison.Ordinal)
                        ? Resolves(family, false)
                        : PinnedFontResolution.Refused(PinnedFontRefusal.FontFileMissing, null)) with
            {
                // The substitution TARGET is the pinned label family, which the renderer can only
                // load where that proprietary face is installed. Render host-free so the assertion
                // stays about the substitution, not about the runner's font inventory.
                Render = (_, dimensions, _) => new RasterRenderResult(dimensions, RasterDeterminism.Pinned),
            },
            successes.Add,
            _ => { });

        Assert.True(outcome.Succeeded);
        Assert.Contains(
            "drawn with " + GanttCatalogues.LabelFontFamily,
            successes[0],
            StringComparison.Ordinal);
        Assert.Contains(
            records,
            record => record.StartsWith("gantt.export font-fallback", StringComparison.Ordinal)
                && record.Contains("families=Comic Sans MS", StringComparison.Ordinal)
                && record.Contains(
                    "substituted=" + GanttCatalogues.LabelFontFamily,
                    StringComparison.Ordinal));
    }

    /// <summary>
    /// When the pinned faces themselves cannot load there is nothing to substitute with, so the
    /// renderer's typed refusal stands and the run reports it as a render refusal (never a
    /// silent swap and never a whole-export refusal).
    /// </summary>
    [Fact]
    public void Run_refuses_the_render_when_the_pinned_faces_are_missing()
    {
        var successes = new List<string>();
        var failures = new List<string>();
        ExportPngRequest request = Request(
            preflight: Preflight(scene: SceneWithText(family: "Comic Sans MS")),
            fontResolver: (_, _) => PinnedFontResolution.Refused(PinnedFontRefusal.FontFileMissing, null)) with
        {
            Render = (_, _, _) => throw new PinnedFontException(
                PinnedFontRefusal.FontFileMissing,
                GanttCatalogues.LabelFontFamily,
                null),
        };

        ExportPngOutcome outcome = ExportPngCommand.Run(request, successes.Add, failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ExportPngRefusal.RenderRefused, outcome.Refusal);
        Assert.Empty(successes);
        Assert.Single(failures);
    }

    private static string GanttChromeKey() =>
        GanttChartChrome.SettingKeyFor(GanttChartChromeRole.Background);

    /// <summary>
    /// A clipboard-target run copies the rendered PNG, presents the clipboard success line,
    /// reports the delivery, and never opens the dialog (R8.8 D8).
    /// </summary>
    [Fact]
    public void Run_copies_to_the_clipboard_by_default_and_reports_it()
    {
        var successes = new List<string>();
        var failures = new List<string>();
        int placements = 0;
        PixelDimensions placed = default;
        var dialog = new StubDialog("unused.png");

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(
                target: PngExportTarget.Clipboard,
                dialog: dialog,
                clipboard: png =>
                {
                    placements++;
                    placed = png.Dimensions;
                    return PngClipboardOutcome.Ok();
                }),
            successes.Add,
            failures.Add);

        Assert.True(outcome.Succeeded);
        Assert.True(outcome.DeliveredToClipboard);
        Assert.Null(outcome.WrittenPath);
        Assert.Equal(1, placements);
        Assert.Equal(outcome.Dimensions, placed);
        Assert.Equal(0, dialog.Calls);
        Assert.Single(successes);
        Assert.Contains("clipboard", successes[0], StringComparison.OrdinalIgnoreCase);
        Assert.Empty(failures);
    }

    /// <summary>A refused clipboard placement is the typed refusal and never a success (D8).</summary>
    [Fact]
    public void Run_reports_a_clipboard_refusal_with_the_ports_message()
    {
        var successes = new List<string>();
        var failures = new List<string>();

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(
                target: PngExportTarget.Clipboard,
                clipboard: _ => PngClipboardOutcome.Fail("The clipboard is busy, so the PNG was not copied.")),
            successes.Add,
            failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ExportPngRefusal.ClipboardRefused, outcome.Refusal);
        Assert.Empty(successes);
        Assert.Single(failures);
        Assert.Contains("clipboard is busy", failures[0], StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>A clipboard target without a clipboard port is a construction refusal (D8).</summary>
    [Fact]
    public void Run_refuses_a_clipboard_target_without_a_clipboard_port()
    {
        var failures = new List<string>();
        ExportPngRequest request = Request(target: PngExportTarget.Clipboard) with { Clipboard = null };

        ExportPngOutcome outcome = ExportPngCommand.Run(request, _ => { }, failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ExportPngRefusal.MissingClipboard, outcome.Refusal);
        Assert.Single(failures);
    }

    /// <summary>
    /// A successful file export remembers the chosen folder (D9), so the next dialog opens
    /// there; the initial directory the command resolved is passed to the dialog.
    /// </summary>
    [Fact]
    public void Run_persists_the_chosen_folder_and_passes_the_initial_directory()
    {
        var dialog = new StubDialog("out.png");
        var persisted = new List<string?>();
        const string initial = @"C:\remembered";

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(dialog: dialog, initialDirectory: initial, persistFolder: persisted.Add),
            _ => { },
            _ => { });

        Assert.True(outcome.Succeeded);
        Assert.Equal(initial, dialog.LastInitialDirectory);
        string? recorded = Assert.Single(persisted);
        Assert.Equal(Path.GetDirectoryName("out.png"), recorded);
    }

    /// <summary>
    /// A stale chart runs the freshness step before the read (D10); a current chart does not.
    /// </summary>
    [Theory]
    [InlineData(false, 1)]
    [InlineData(true, 0)]
    public void Run_ensures_freshness_only_when_the_chart_is_stale(bool chartRenderCurrent, int expectedCalls)
    {
        var order = new List<string>();
        EditableExportPreflightRequest preflight = Preflight() with
        {
            TableReader = new StubTableReader(order: order),
        };

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(
                preflight: preflight,
                chartRenderCurrent: chartRenderCurrent,
                ensureCurrent: () => order.Add("ensure")),
            _ => { },
            _ => { });

        Assert.True(outcome.Succeeded);
        Assert.Equal(
            expectedCalls,
            order.Count(entry => string.Equals(entry, "ensure", StringComparison.Ordinal)));
        if (expectedCalls == 1)
        {
            Assert.Equal(["ensure", "read"], order);
        }
    }

    private sealed class StubDialog(string? path, Action? onPick = null) : ISaveFileDialogPort
    {
        internal int Calls { get; private set; }

        /// <summary>The initial directory the command supplied, for the folder-default pin.</summary>
        internal string? LastInitialDirectory { get; private set; }

        public string? PickPath(string suggestedFileName, string? initialDirectory)
        {
            Calls++;
            LastInitialDirectory = initialDirectory;
            onPick?.Invoke();
            return path;
        }
    }

    private sealed class StubTableReader(
        GanttTableReadOutcome? outcome = null,
        List<string>? order = null) : IGanttTableReader
    {
        public GanttTableReadOutcome Read()
        {
            order?.Add("read");
            return outcome ?? GanttTableReadOutcome.Ok([]);
        }

        public GanttTableReadOutcome Read(IReadOnlyList<int> rowNumbers) => Read();
    }

    private sealed class StubConfigReader(
        IReadOnlyDictionary<string, string> settings,
        ConfigReadRefusalReason? refusal = null) : IConfigCatalogueReader
    {
        public ConfigReadOutcome Read() => refusal is { } reason
            ? ConfigReadOutcome.Refused(reason)
            : ConfigReadOutcome.Ok(
                "wb-id",
                settings,
                GanttStyleRegistry.Empty,
                new Dictionary<string, double>());
    }

    private sealed class StubMeasurement : IPanelGridMeasurementPort
    {
        public PanelGridOutcome Measure(
            IReadOnlyList<string> includedColumns,
            IReadOnlyList<double>? rowHeights = null) =>
            PanelGridOutcome.Ok(Grid());
    }
}
