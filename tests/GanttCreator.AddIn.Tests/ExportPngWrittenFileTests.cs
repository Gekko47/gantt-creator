using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using GanttCreator.Raster;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// R8.8 D2/D3: the file the command leaves on disk, verified by an independent reader
/// (<see cref="PngFileInspection"/>) rather than by the writer's own helpers. The pixels
/// come from the production R8.4/R8.5/R8.6 entries, so the assertions cover the whole
/// chain: the pre-flight's bounds size the canvas, the encoder stamps the density, and the
/// atomic writer's output is what the reader sees.
/// </summary>
public sealed class ExportPngWrittenFileTests
{
    private static readonly RectD Bounds = new(0, 0, 800, 400);

    [Fact]
    public void The_written_file_carries_the_dimensions_the_budget_computed()
    {
        string path = TempPath();
        try
        {
            ExportPngOutcome outcome = ExportPngCommand.Run(
                Request(path),
                _ => { },
                message => Assert.Fail(message));

            Assert.True(outcome.Succeeded);
            Assert.True(File.Exists(path));

            PngInspection inspection = PngFileInspection.Read(path);

            Assert.Equal(outcome.Dimensions.PixelWidth, inspection.Width);
            Assert.Equal(outcome.Dimensions.PixelHeight, inspection.Height);

            // The dimensions are the chart's own bounds at 300 DPI, so the file cannot be
            // sized from any cached value: 800 pt x 400 pt → 3333 x 1667 px.
            PixelDimensions expected = ExportSize.ToPixelsFromPoints(Bounds.Width, Bounds.Height);
            Assert.Equal(expected, outcome.Dimensions);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void The_written_file_carries_the_300_dpi_metadata()
    {
        string path = TempPath();
        try
        {
            Assert.True(ExportPngCommand.Run(Request(path), _ => { }, message => Assert.Fail(message)).Succeeded);

            PngInspection inspection = PngFileInspection.Read(path);

            // 11811 pixels per metre is exactly 300 DPI; unit 1 is metre.
            Assert.Equal(11811, inspection.PixelsPerMetreX);
            Assert.Equal(11811, inspection.PixelsPerMetreY);
            Assert.Equal(1, inspection.PhysUnit);
            Assert.True(inspection.DataLength > 0, "The file carries no image data.");
            Assert.Equal("IEND", inspection.ChunkTypes[^1]);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void The_written_file_is_a_valid_png_by_an_independent_crc_check()
    {
        string path = TempPath();
        try
        {
            Assert.True(ExportPngCommand.Run(Request(path), _ => { }, message => Assert.Fail(message)).Succeeded);

            // Read() throws InvalidDataException on a bad signature, a bad CRC or a missing
            // IHDR; reaching this line is the assertion.
            PngInspection inspection = PngFileInspection.Read(path);

            Assert.Equal("IHDR", inspection.ChunkTypes[0]);
            Assert.Equal(8, inspection.BitDepth);
        }
        finally
        {
            Delete(path);
        }
    }

    private static string TempPath() =>
        Path.Combine(Path.GetTempPath(), "gantt-export-" + Guid.NewGuid().ToString("N") + ".png");

    private static void Delete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
            // A leftover temp file is not a test failure; the assertion already ran.
        }
    }

    private static ExportPngRequest Request(string path)
    {
        GanttScene scene = GanttScene.TryCreate(Bounds, Bounds, [], []).Scene!;
        return new ExportPngRequest(
            new EditableExportPreflightRequest(
                new StubTableReader(),
                new StubConfigReader(),
                new StubMeasurement(),
                (_, _, _, _, _) => SceneBuildRequestOutcome.Ok(new SceneBuildRequest()),
                (_, _, _) => new ExportBoundsOutcome(new ExportBounds(Bounds), null, null, scene)),
            PngExportTarget.File,
            ChartRenderCurrent: true,
            EnsureCurrent: null,
            new FixedDialog(path),
            Clipboard: null,
            InitialDirectory: null,
            PersistFolder: null,
            (s, dimensions, background) => SceneRasterRenderer.Render(s, dimensions, background),
            PngExportEncoder.Encode,
            (png, targetPath) => AtomicPngFileWriter.Write(png, targetPath),
            null);
    }

    private sealed class FixedDialog(string path) : ISaveFileDialogPort
    {
        public string? PickPath(string suggestedFileName, string? initialDirectory) => path;
    }

    private sealed class StubTableReader : IGanttTableReader
    {
        public GanttTableReadOutcome Read() => GanttTableReadOutcome.Ok([]);

        public GanttTableReadOutcome Read(IReadOnlyList<int> rowNumbers) => Read();
    }

    private sealed class StubConfigReader : IConfigCatalogueReader
    {
        public ConfigSchemaObservation? ObserveSchema() => null;

        public ConfigReadOutcome Read() => ConfigReadOutcome.Ok(
            "wb-id",
            new Dictionary<string, string>(),
            GanttStyleRegistry.Empty,
            new Dictionary<string, double>());
    }

    private sealed class StubMeasurement : IPanelGridMeasurementPort
    {
        public PanelGridOutcome Measure(
            IReadOnlyList<string> includedColumns,
            IReadOnlyList<double>? rowHeights = null) =>
            PanelGridOutcome.Ok(
                PanelCellGrid.TryCreate(
                    [new PanelColumn("Type", 60.0)],
                    [15.0],
                    15.0,
                    null).Grid!);
    }
}
