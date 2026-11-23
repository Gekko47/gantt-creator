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
        IReadOnlyDictionary<string, string>? settings = null)
    {
        GanttScene scene = EmptyScene();
        return new EditableExportPreflightRequest(
            new StubTableReader(),
            new StubConfigReader(settings ?? new Dictionary<string, string>()),
            new StubMeasurement(),
            (_, _, _, _, _) => SceneBuildRequestOutcome.Ok(new SceneBuildRequest()),
            (_, _, _) => new ExportBoundsOutcome(new ExportBounds(Bounds), null, null, scene));
    }

    private static ExportPngRequest Request(
        Action? onRender = null,
        string? widthText = "10",
        ExportUnit unit = ExportUnit.Centimetres,
        ISaveFileDialogPort? dialog = null,
        PngWrite? write = null,
        EditableExportPreflightRequest? preflight = null,
        Action<RasterBackground>? onBackground = null) =>
        new(
            preflight ?? Preflight(),
            widthText,
            unit,
            dialog ?? new StubDialog("out.png"),
            (scene, dimensions, background) =>
            {
                onBackground?.Invoke(background);
                onRender?.Invoke();
                return SceneRasterRenderer.Render(scene, dimensions, background);
            },
            PngExportEncoder.Encode,
            write ?? ((png, path) => new PngWriteOutcome(true, null, "written", "record")),
            null);

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
    /// Every input-dependent refusal precedes the dialog, so a refused export never makes
    /// the user choose a path first (D1).
    /// </summary>
    [Fact]
    public void Run_refuses_an_invalid_width_before_the_dialog()
    {
        var dialog = new StubDialog("out.png");
        var failures = new List<string>();

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(widthText: "abc", dialog: dialog),
            _ => { },
            failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ExportPngRefusal.InvalidWidth, outcome.Refusal);
        Assert.Equal(0, dialog.Calls);
        Assert.Single(failures);
        Assert.False(outcome.Cancelled);
    }

    /// <summary>The pixel budget is checked against the FRESH bounds, and names the limit.</summary>
    [Fact]
    public void Run_refuses_a_width_over_the_limit_before_the_dialog()
    {
        var dialog = new StubDialog("out.png");
        var failures = new List<string>();

        ExportPngOutcome outcome = ExportPngCommand.Run(
            Request(widthText: "100000", dialog: dialog),
            _ => { },
            failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ExportPngRefusal.WidthRefused, outcome.Refusal);
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

    private static string GanttChromeKey() =>
        GanttChartChrome.SettingKeyFor(GanttChartChromeRole.Background);

    private sealed class StubDialog(string? path, Action? onPick = null) : ISaveFileDialogPort
    {
        internal int Calls { get; private set; }

        public string? PickPath(string suggestedFileName)
        {
            Calls++;
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

    private sealed class StubConfigReader(IReadOnlyDictionary<string, string> settings)
        : IConfigCatalogueReader
    {
        public ConfigReadOutcome Read() => ConfigReadOutcome.Ok(
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
