using System.Diagnostics;
using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using GanttCreator.Raster;

namespace GanttCreator.AddIn;

/// <summary>Renders a scene to a pixel surface (the R8.4 entry, injectable for tests).</summary>
/// <param name="scene">The immutable scene to draw.</param>
/// <param name="dimensions">The validated pixel dimensions.</param>
/// <param name="background">The explicit canvas background.</param>
/// <returns>The owned pixel result.</returns>
internal delegate RasterRenderResult PngRender(
    GanttScene scene,
    PixelDimensions dimensions,
    RasterBackground background);

/// <summary>Encodes a rendered raster as PNG bytes (the R8.5 entry, injectable for tests).</summary>
/// <param name="result">The rendered pixel result.</param>
/// <returns>The encoded bytes with their dimensions.</returns>
internal delegate PngExportBytes PngEncode(RasterRenderResult result);

/// <summary>Writes and validates the bytes atomically (the R8.6 entry, injectable for tests).</summary>
/// <param name="png">The encoded bytes.</param>
/// <param name="targetPath">The chosen file path.</param>
/// <returns>The write outcome; success only after the writer re-read and validated the file.</returns>
internal delegate PngWriteOutcome PngWrite(PngExportBytes png, string targetPath);

/// <summary>Why the PNG export produced no file (R8.8).</summary>
/// <remarks>
/// Every member is a typed, expected outcome rather than an exception; the construction
/// members exist so a dependency the command cannot run without is reported, never
/// dereferenced. A user-cancelled dialog is <em>not</em> a refusal: it is the silent
/// <see cref="ExportPngOutcome.Cancelled"/> state.
/// </remarks>
internal enum ExportPngRefusal
{
    /// <summary>No command request was supplied.</summary>
    NullRequest = 0,

    /// <summary>No shared pre-flight request was supplied.</summary>
    MissingPreflight = 1,

    /// <summary>No save-file dialog port was supplied.</summary>
    MissingDialog = 2,

    /// <summary>No render seam was supplied.</summary>
    MissingRender = 3,

    /// <summary>No encode seam was supplied.</summary>
    MissingEncode = 4,

    /// <summary>No write seam was supplied.</summary>
    MissingWrite = 5,

    /// <summary>The table could not be read, so no file was written.</summary>
    TableReadRefused = 6,

    /// <summary>The workbook configuration could not be read, so no file was written.</summary>
    ConfigReadRefused = 7,

    /// <summary>The live panel grid could not be measured, so no file was written.</summary>
    MeasurementRefused = 8,

    /// <summary>The export scene request was refused, so no file was written.</summary>
    ExportRequestRefused = 9,

    /// <summary>The export bounds or scene were refused, so no file was written.</summary>
    BoundsRefused = 10,

    /// <summary>The composition's bounds convert to pixels beyond the safety limits.</summary>
    DimensionsRefused = 11,

    /// <summary>The render refused (a pinned font could not be loaded).</summary>
    RenderRefused = 12,

    /// <summary>The atomic write or its post-write validation refused.</summary>
    WriteRefused = 13,
}

/// <summary>The typed result of one PNG export run (R8.8).</summary>
/// <param name="Succeeded">Whether the file was written and independently validated.</param>
/// <param name="Cancelled">Whether the user cancelled the save dialog; nothing was written.</param>
/// <param name="WrittenPath">The written path on success, or null.</param>
/// <param name="Dimensions">The exported pixel dimensions on success, or the default.</param>
/// <param name="DroppedRowCount">
/// How many body rows carried a blocking validation error and were left out of the export,
/// or zero when nothing was dropped.
/// </param>
/// <param name="Refusal">The typed refusal, or <see langword="null"/> on success or cancellation.</param>
/// <param name="Message">The success sentence, the one actionable refusal line, or empty on cancellation.</param>
internal sealed record ExportPngOutcome(
    bool Succeeded,
    bool Cancelled,
    string? WrittenPath,
    PixelDimensions Dimensions,
    int DroppedRowCount,
    ExportPngRefusal? Refusal,
    string Message)
{
    /// <summary>Creates the success outcome.</summary>
    /// <param name="writtenPath">The path the validated file was written to.</param>
    /// <param name="dimensions">The exported pixel dimensions.</param>
    /// <param name="droppedRows">The blocking-error row count left out of the export.</param>
    /// <param name="message">The success sentence.</param>
    /// <returns>The success outcome.</returns>
    internal static ExportPngOutcome Ok(
        string writtenPath,
        PixelDimensions dimensions,
        int droppedRows,
        string message) =>
        new(true, false, writtenPath, dimensions, droppedRows, null, message);

    /// <summary>Creates the silent cancellation outcome: nothing was written, nothing is presented.</summary>
    /// <param name="droppedRows">The blocking-error row count the pre-flight had already counted.</param>
    /// <returns>The cancellation outcome.</returns>
    internal static ExportPngOutcome CancelledOutcome(int droppedRows) =>
        new(false, true, null, default, droppedRows, null, string.Empty);

    /// <summary>Creates a refusal outcome.</summary>
    /// <param name="refusal">The typed refusal.</param>
    /// <param name="message">The one actionable line.</param>
    /// <param name="droppedRows">The blocking-error row count, when known.</param>
    /// <returns>The refusal outcome.</returns>
    internal static ExportPngOutcome Refuse(
        ExportPngRefusal refusal,
        string message,
        int droppedRows = 0) =>
        new(false, false, null, default, droppedRows, refusal, message);
}

/// <summary>The inputs one PNG export run consumes (R8.8).</summary>
/// <param name="Preflight">
/// The shared pre-flight request (table reader, configuration reader, grid measurement,
/// export request factory, bounds resolver). Required.
/// </param>
/// <param name="Dialog">The save-file dialog port. Required.</param>
/// <param name="Render">The render entry. Required.</param>
/// <param name="Encode">The encode entry. Required.</param>
/// <param name="Write">The atomic write-and-validate entry. Required.</param>
/// <param name="TechnicalRecord">Receives the export records, or null to discard them.</param>
internal sealed record ExportPngRequest(
    EditableExportPreflightRequest? Preflight,
    ISaveFileDialogPort? Dialog,
    PngRender? Render,
    PngEncode? Encode,
    PngWrite? Write,
    Action<string>? TechnicalRecord = null);

/// <summary>
/// R8.8 D1/D7: the Export PNG click. Runs the shared pre-flight, sizes the image from the
/// fresh composition bounds at 300 DPI, shows the save dialog, renders, encodes, and writes
/// atomically with post-write validation — reporting success only after the file was validated.
/// </summary>
/// <remarks>
/// <para>
/// <b>The order is the contract.</b> Every input-dependent refusal (pre-flight, dimension
/// budget) precedes the dialog, so a refused export never makes the user pick a path first;
/// the render precedes the write, and the write's own validator precedes success.
/// </para>
/// <para>
/// <b>The image is the chart's own size (owner ruling 2026-11-23).</b> The pixel dimensions
/// derive from the pre-flight's freshly resolved export-composition bounds in points, so the
/// PNG is what the user sees and a table changed since the last render can never size the
/// file for a chart that no longer exists. There is no width input.
/// </para>
/// <para>
/// <b>The render is synchronous on Excel's command thread</b> (D6). The
/// <see cref="ExportSize.MaxTotalPixels"/> budget bounds the legal worst case; cancellation
/// for long non-COM rendering is R9.1's row.
/// </para>
/// <para>
/// <b>A scene-contract fault is not caught here.</b> <see cref="SceneRasterRenderer.Render"/>
/// throws <see cref="ArgumentException"/> when a text primitive carries no resolvable family
/// or size, which is a defect rather than a user state; it escapes to the ribbon
/// <see cref="CommandBoundary"/>'s generic arm, where every unexpected fault is contained.
/// A missing pinned font is an environment state and is presented as its own message.
/// </para>
/// </remarks>
internal static class ExportPngCommand
{
    /// <summary>The one actionable line used when a required dependency is absent.</summary>
    internal const string ConstructionMessage =
        "The export could not start. Try again, or run Refresh chart if the problem persists.";

    /// <summary>The file name the save dialog is pre-filled with.</summary>
    internal const string SuggestedFileName = "gantt.png";

    /// <summary>
    /// Runs the live PNG export in the current Excel session: wires the production seams,
    /// runs once, and restores the add-in tab.
    /// </summary>
    /// <remarks>
    /// <b>No protection window and no application-state scope.</b> The pre-flight reads and
    /// measures, the renderer writes pixels in memory and the writer writes a file; no shape
    /// and no cell is touched, so ADR-0055's release-and-restore window and ADR-0020's
    /// screen-updating scope would only add COM traffic.
    /// </remarks>
    internal static void RunForExcel()
    {
        object? application = ExcelDna.Integration.ExcelDnaUtil.Application;
        Action<string> trace = CommandBoundary.Instance.WriteTechnicalRecord;

        // D9-5: the status bar is the non-modal channel; a user who hid it gets the dialog.
        Action<string> successPresenter = CopyEditableCommand.SelectSuccessPresenter(
            RibbonStateService.Instance.IsStatusBarVisible(),
            RibbonStatusBar.Warn,
            CommandErrorDialog.Show);

        EditableExportPreflightRequest preflight = new(
            new ExcelGanttTableReader(application),
            new ExcelConfigCatalogueReader(application),
            new ExcelPanelGridMeasurement(application),
            EditableExportRequest.Create,
            ExportBoundsResolver.TryResolve,
            trace);

        ExportPngRequest request = new(
            preflight,
            new WindowsSaveFileDialogPort(),
            RenderLive,
            EncodeLive,
            WriteLive,
            trace);

        _ = Run(request, successPresenter, CommandErrorDialog.Show);

        RibbonStateService.Instance.ActivateGanttTab();
    }

    /// <summary>Runs one PNG export pipeline and presents its one-line result.</summary>
    /// <param name="request">The run's inputs; see <see cref="ExportPngRequest"/>.</param>
    /// <param name="successPresenter">Receives the success sentence.</param>
    /// <param name="failurePresenter">Receives the refusal sentence.</param>
    /// <returns>The typed outcome; the presenters are the only side effect.</returns>
    internal static ExportPngOutcome Run(
        ExportPngRequest? request,
        Action<string> successPresenter,
        Action<string> failurePresenter)
    {
        ArgumentNullException.ThrowIfNull(successPresenter);
        ArgumentNullException.ThrowIfNull(failurePresenter);

        if (request is null)
        {
            return Refuse(ExportPngRefusal.NullRequest, ConstructionMessage, failurePresenter);
        }

        // Construction members first, so a missing dependency is reported, never dereferenced.
        if (request.Preflight is not { } preflightRequest)
        {
            return Refuse(ExportPngRefusal.MissingPreflight, ConstructionMessage, failurePresenter);
        }

        if (request.Dialog is not { } dialog)
        {
            return Refuse(ExportPngRefusal.MissingDialog, ConstructionMessage, failurePresenter);
        }

        if (request.Render is not { } render)
        {
            return Refuse(ExportPngRefusal.MissingRender, ConstructionMessage, failurePresenter);
        }

        if (request.Encode is not { } encode)
        {
            return Refuse(ExportPngRefusal.MissingEncode, ConstructionMessage, failurePresenter);
        }

        if (request.Write is not { } write)
        {
            return Refuse(ExportPngRefusal.MissingWrite, ConstructionMessage, failurePresenter);
        }

        long started = Stopwatch.GetTimestamp();

        // Stage 1 — the shared pre-flight, unchanged (D5).
        EditableExportPreflightOutcome preflight = EditableExportPreflight.Run(preflightRequest);
        if (!preflight.Succeeded || preflight.Preparation is not { } preparation)
        {
            return Refuse(
                TranslatePreflightRefusal(preflight.Refusal),
                DescribePreflightRefusal(preflight),
                failurePresenter,
                preflight.DroppedRowCount);
        }

        ExportBoundsOutcome resolution = preparation.Resolution;
        if (resolution.Scene is not { } scene || resolution.Bounds is not { } bounds)
        {
            // The pre-flight's guard already refused this state, so the arm is unreachable;
            // it is still named rather than dereferenced.
            return Refuse(
                ExportPngRefusal.BoundsRefused,
                "The chart could not be measured for export, so no PNG was written.",
                failurePresenter,
                preflight.DroppedRowCount);
        }

        // Stage 2 — pixel dimensions from the FRESH composition bounds at 300 DPI (D7).
        PixelDimensions dimensions;
        try
        {
            dimensions = ExportSize.ToPixelsFromPoints(bounds.Bounds.Width, bounds.Bounds.Height);
        }
        catch (ArgumentOutOfRangeException ex)
        {
            return Refuse(
                ExportPngRefusal.DimensionsRefused,
                $"{TrimParameterSuffix(ex.Message)} Reduce the chart's size and export again.",
                failurePresenter,
                preflight.DroppedRowCount);
        }

        // Stage 3 — the dialog is the last gate before the expensive render, and a cancel is
        // silent: nothing is written and nothing is presented.
        string? path = dialog.PickPath(SuggestedFileName);
        if (string.IsNullOrWhiteSpace(path))
        {
            return ExportPngOutcome.CancelledOutcome(preflight.DroppedRowCount);
        }

        // Diagnosis hook (live defect 2026-11-23): the raster renderer refuses a text
        // primitive whose style names no font family. The trace names the style keys and
        // families BEFORE the refusal, so a live export says which family the scene is
        // missing instead of only that one is missing.
        TraceTextFamilies(scene, request.TechnicalRecord);

        // Stage 4 + 5 — render, then encode. The render owns the surface; the bytes outlive it.
        PngExportBytes png;
        try
        {
            using RasterRenderResult rendered = render(scene, dimensions, ChartBackground(preparation.Settings));
            png = encode(rendered);
        }
        catch (PinnedFontException ex)
        {
            return Refuse(
                ExportPngRefusal.RenderRefused,
                ex.Message,
                failurePresenter,
                preflight.DroppedRowCount);
        }

        // Stage 6 — write, and validate from disk before success (R8.6 D2).
        PngWriteOutcome written = write(png, path);
        if (!string.IsNullOrWhiteSpace(written.TechnicalRecord))
        {
            request.TechnicalRecord?.Invoke(written.TechnicalRecord);
        }

        if (!written.Succeeded)
        {
            return Refuse(
                ExportPngRefusal.WriteRefused,
                written.ActionableMessage,
                failurePresenter,
                preflight.DroppedRowCount);
        }

        // Stage 7 — report. Counts and sizes only; never schedule content.
        string message = string.Create(
            CultureInfo.InvariantCulture,
            $"Exported {dimensions.PixelWidth} \u00d7 {dimensions.PixelHeight} px to {path}.");
        successPresenter(message);

        request.TechnicalRecord?.Invoke(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.export png path={path} width={dimensions.PixelWidth} height={dimensions.PixelHeight} rows-dropped={preflight.DroppedRowCount} ms={Stopwatch.GetElapsedTime(started).TotalMilliseconds:0.#}"));

        return ExportPngOutcome.Ok(path, dimensions, preflight.DroppedRowCount, message);
    }

    /// <summary>The canvas background: the workbook's chart background, token fallback (R8.8 D1).</summary>
    /// <param name="settings">The workbook's effective settings map.</param>
    /// <returns>The explicit background so the edge pixels carry the chart background.</returns>
    /// <remarks>
    /// Reads the same key and the same catalogue fallback the scene uses
    /// (<see cref="GanttChartChrome.SettingKeyFor"/> / <see cref="GanttChartChrome.DefaultHexFor"/>),
    /// so a themed background cannot differ between the sheet and the image.
    /// </remarks>
    internal static RasterBackground ChartBackground(IReadOnlyDictionary<string, string> settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        string key = GanttChartChrome.SettingKeyFor(GanttChartChromeRole.Background);
#pragma warning disable IDE0046 // The short-circuit is the point: the value is only read when the key is present.
        if (settings.TryGetValue(key, out string? stored)
            && ColourHex.TryParse(stored, out ColourHex? parsed)
            && parsed is not null)
        {
            return RasterBackground.FromColour(parsed);
        }
#pragma warning restore IDE0046

        return RasterBackground.FromColour(
            ColourHex.Parse(GanttChartChrome.DefaultHexFor(GanttChartChromeRole.Background)));
    }

    /// <summary>
    /// Records the distinct style/family pairs among the scene's text primitives, so a
    /// font-refusal is diagnosable from the log. Counts and style keys only; a family-less
    /// style is named <c>&lt;none&gt;</c>, which is exactly the state the raster
    /// renderer refuses.
    /// </summary>
    /// <param name="scene">The export scene about to be rendered.</param>
    /// <param name="trace">The technical-record sink, or null to discard the line.</param>
    private static void TraceTextFamilies(GanttScene scene, Action<string>? trace)
    {
        if (trace is null)
        {
            return;
        }

        SortedDictionary<string, int> groups = new(StringComparer.Ordinal);
        foreach (SceneText text in scene.Primitives.OfType<SceneText>())
        {
            string family = string.IsNullOrWhiteSpace(text.Style.FontFamily)
                ? "<none>"
                : text.Style.FontFamily;
            string key = text.Style.StyleKey + "=" + family;
            groups[key] = groups.TryGetValue(key, out int count) ? count + 1 : 1;
        }

        trace("gantt.export text-families "
            + string.Join(";", groups.Select(pair => $"{pair.Key}x{pair.Value}")));
    }

    /// <summary>Maps a pre-flight refusal to this command's typed refusal.</summary>
    /// <param name="refusal">The pre-flight refusal, or null when unprobeable.</param>
    /// <returns>The command's refusal member.</returns>
    private static ExportPngRefusal TranslatePreflightRefusal(
        EditableExportPreflightRefusal? refusal) => refusal switch
        {
            EditableExportPreflightRefusal.TableReadRefused => ExportPngRefusal.TableReadRefused,
            EditableExportPreflightRefusal.ConfigReadRefused => ExportPngRefusal.ConfigReadRefused,
            EditableExportPreflightRefusal.MeasurementRefused => ExportPngRefusal.MeasurementRefused,
            EditableExportPreflightRefusal.ExportRequestRefused => ExportPngRefusal.ExportRequestRefused,
            EditableExportPreflightRefusal.BoundsRefused => ExportPngRefusal.BoundsRefused,
            EditableExportPreflightRefusal.NullRequest => ExportPngRefusal.MissingPreflight,
            EditableExportPreflightRefusal.MissingTableReader => ExportPngRefusal.MissingPreflight,
            EditableExportPreflightRefusal.MissingConfigReader => ExportPngRefusal.MissingPreflight,
            EditableExportPreflightRefusal.MissingGridMeasurement => ExportPngRefusal.MissingPreflight,
            EditableExportPreflightRefusal.MissingExportRequestFactory => ExportPngRefusal.MissingPreflight,
            EditableExportPreflightRefusal.MissingBoundsResolver => ExportPngRefusal.MissingPreflight,
            _ => ExportPngRefusal.MissingPreflight,
        };

    /// <summary>Builds this command's one actionable line for a pre-flight refusal.</summary>
    /// <param name="preflight">The refused pre-flight outcome.</param>
    /// <returns>The user-facing sentence; the shared pre-flight owns no text of its own.</returns>
    private static string DescribePreflightRefusal(EditableExportPreflightOutcome preflight) =>
        preflight.Refusal switch
        {
            EditableExportPreflightRefusal.TableReadRefused =>
                "The table could not be read, so no PNG was written.",
            EditableExportPreflightRefusal.ConfigReadRefused =>
                "The workbook configuration could not be read, so no PNG was written.",
            EditableExportPreflightRefusal.MeasurementRefused =>
                "The table could not be measured, so no PNG was written.",
            EditableExportPreflightRefusal.ExportRequestRefused =>
                preflight.ExportRequestMessage
                ?? "The chart could not be prepared for export, so no PNG was written.",
            EditableExportPreflightRefusal.BoundsRefused =>
                "The chart could not be measured for export, so no PNG was written.",
            EditableExportPreflightRefusal.NullRequest => ConstructionMessage,
            EditableExportPreflightRefusal.MissingTableReader => ConstructionMessage,
            EditableExportPreflightRefusal.MissingConfigReader => ConstructionMessage,
            EditableExportPreflightRefusal.MissingGridMeasurement => ConstructionMessage,
            EditableExportPreflightRefusal.MissingExportRequestFactory => ConstructionMessage,
            EditableExportPreflightRefusal.MissingBoundsResolver => ConstructionMessage,
            _ => ConstructionMessage,
        };

    /// <summary>Removes .NET's parameter-name suffix from an out-of-range message.</summary>
    /// <param name="message">The exception message.</param>
    /// <returns>The message a user can act on.</returns>
    private static string TrimParameterSuffix(string message)
    {
        int suffix = message.IndexOf(" (Parameter '", StringComparison.Ordinal);
        return suffix >= 0 ? message[..suffix] : message;
    }

    /// <summary>Presents one refusal and returns it.</summary>
    /// <param name="refusal">The typed refusal.</param>
    /// <param name="message">The one actionable line.</param>
    /// <param name="failurePresenter">The presenter that receives the line.</param>
    /// <param name="droppedRows">The blocking-error row count, when known.</param>
    /// <returns>The refusal outcome.</returns>
    private static ExportPngOutcome Refuse(
        ExportPngRefusal refusal,
        string message,
        Action<string> failurePresenter,
        int droppedRows = 0)
    {
        failurePresenter(message);
        return ExportPngOutcome.Refuse(refusal, message, droppedRows);
    }

    /// <summary>Renders through the production R8.4 entry.</summary>
    /// <param name="scene">The scene to draw.</param>
    /// <param name="dimensions">The validated pixel dimensions.</param>
    /// <param name="background">The explicit canvas background.</param>
    /// <returns>The owned pixel result.</returns>
    private static RasterRenderResult RenderLive(
        GanttScene scene,
        PixelDimensions dimensions,
        RasterBackground background) =>
        SceneRasterRenderer.Render(scene, dimensions, background);

    /// <summary>Encodes through the production R8.5 entry.</summary>
    /// <param name="result">The rendered pixel result.</param>
    /// <returns>The encoded PNG bytes.</returns>
    private static PngExportBytes EncodeLive(RasterRenderResult result) => PngExportEncoder.Encode(result);

    /// <summary>Writes and validates through the production R8.6 entry.</summary>
    /// <param name="png">The encoded bytes.</param>
    /// <param name="targetPath">The chosen path.</param>
    /// <returns>The write outcome.</returns>
    private static PngWriteOutcome WriteLive(PngExportBytes png, string targetPath) =>
        AtomicPngFileWriter.Write(png, targetPath);
}
