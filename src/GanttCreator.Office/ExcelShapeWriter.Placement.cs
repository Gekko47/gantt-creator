using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

public partial class ExcelShapeWriter
{
    /// <summary>
    /// Writes a geometry onto an existing shape at the single rounding boundary, unless
    /// the live box already matches it.
    /// </summary>
    /// <param name="shape">The shape to update.</param>
    /// <param name="kind">The shape kind, which selects which members are written.</param>
    /// <param name="geometry">The scene geometry, in points.</param>
    /// <param name="hostOwnsExtent">
    /// Whether the host owns the vertical extent (ADR-0052 D6–D8), which widens the
    /// width/height acceptance exactly as <see cref="GeometryMatches"/> does.
    /// </param>
    /// <returns><see langword="true"/> when the box matches or every member was written.</returns>
    /// <remarks>
    /// <b>See <see cref="ApplyPlacement"/> for why the shape is placed with
    /// <c>xlMove</c> and not <c>xlMoveAndSize</c>.</b>
    /// </remarks>
    /// <remarks>
    /// <para>
    /// R4.1 D3: this is the <em>only</em> place a scene point becomes a host
    /// <c>Single</c>, via <c>GeometryMath.SnapToDisplayPrecision</c>. A caller
    /// never pre-rounds, and there is no second conversion anywhere on the path.
    /// </para>
    /// <para>
    /// Every renderable kind is rewritten in place, milestones included: a
    /// diamond auto-shape is positioned by <c>Left</c>/<c>Top</c>/<c>Width</c>/
    /// <c>Height</c> like any other auto-shape, because the host derives its
    /// vertices from that box. (When milestones were freeforms this was not
    /// expressible - a freeform exposes its vertices as a read-only
    /// <c>Vertices</c> collection - which is one of the reasons the diamond
    /// auto-shape is preferred. An <em>undefined</em> kind is the only case that
    /// returns <see langword="false"/>, making the limitation explicit rather
    /// than silently ignoring the requested geometry.
    /// </para>
    /// </remarks>
    private static bool ApplyGeometry(
        Excel.Shape shape,
        OfficeShapeKind kind,
        OfficeShapeGeometry geometry,
        bool hostOwnsExtent)
    {
        ArgumentNullException.ThrowIfNull(geometry);

        if (!HasRequiredGeometry(kind, geometry))
        {
            return false;
        }

        double left;
        double top;
        double width;
        double height;
        if (kind == OfficeShapeKind.Line)
        {
            // Probed 2026-09-27: a created line carries its absolute endpoints,
            // but the Shape surface exposes only the bounding box, so an update
            // is expressed as left/top/width/height. A negative-width line
            // (right-to-left) is normalised here rather than rejected, because
            // the host's own bounding box cannot represent it.
            PointD from = geometry.From!.Value;
            PointD to = geometry.To!.Value;
            left = Math.Min(from.X, to.X);
            top = Math.Min(from.Y, to.Y);
            width = Math.Abs(to.X - from.X);
            height = Math.Abs(to.Y - from.Y);
        }
        else if (kind is OfficeShapeKind.Rectangle or OfficeShapeKind.TextBox or OfficeShapeKind.Diamond)
        {
            // A diamond auto-shape is positioned exactly like any other
            // auto-shape, so an in-place update only has to move and resize it.
            // The vertices are derived from the box by the host, so there is no
            // vertex collection to rewrite.
            RectD bounds = geometry.Bounds!.Value;
            left = bounds.X;
            top = bounds.Y;
            width = bounds.Width;
            height = bounds.Height;
        }
        else
        {
            // An undefined kind is not expressible. Returning false here makes
            // the limitation explicit rather than silently ignoring the request.
            return false;
        }

        // Read-back verification (ADR-0054 amendment 2026-11-20): a converged shape pays
        // four reads and no write. The tolerances are the same ones GeometryMatches uses,
        // including the host-owned extent's wider acceptance, so the skip decision and the
        // write path cannot disagree about "already correct".
        bool hostOwnsWidth = hostOwnsExtent && kind == OfficeShapeKind.Line;
        bool matches = MatchesLive(shape.Left, left)
            && MatchesLive(shape.Top, top)
            && (hostOwnsWidth
                ? MatchesLiveWithin(shape.Width, width, HostOwnedGeometry.HostResizeTolerancePt)
                : MatchesLive(shape.Width, width))
            && (hostOwnsExtent
                ? MatchesLiveWithin(shape.Height, height, HostOwnedGeometry.HostResizeTolerancePt)
                : MatchesLive(shape.Height, height));
        if (matches)
        {
            return true;
        }

        shape.Left = GeometryMath.SnapToDisplayPrecision(left);
        shape.Top = GeometryMath.SnapToDisplayPrecision(top);
        shape.Width = GeometryMath.SnapToDisplayPrecision(width);
        shape.Height = GeometryMath.SnapToDisplayPrecision(height);
        return true;
    }

    /// <summary>
    /// Places the shape with the ownership the request's class requires: <c>xlMove</c>
    /// for a scene-owned shape, <c>xlMoveAndSize</c> for a host-owned one.
    /// </summary>
    /// <param name="shape">The shape to place.</param>
    /// <param name="hostOwnsExtent">
    /// Whether the shape spans the plot's full vertical extent and the host therefore
    /// owns its height (ADR-0052 D6).
    /// </param>
    /// <remarks>
    /// <para>
    /// <b>Why <c>xlMove</c> for a scene-owned shape and not the host's default
    /// <c>xlMoveAndSize</c>.</b> The scene owns the geometry, so the host resizing a shape
    /// is the renderer losing control of its own output. Excel creates shapes with
    /// <c>xlMoveAndSize</c>, which ties the displayed height to the rows the shape spans,
    /// so hiding, un-hiding or re-heighting a row inside that span resizes the shape
    /// without any scene change. Measured 2026-11-03: the scene declared the splitter band
    /// at 18pt on every refresh while the live band was four rows tall in one state and a
    /// thin line in another — a rewritten 18pt rect did not survive a row-visibility
    /// change.
    /// </para>
    /// <para>
    /// <c>xlMove</c> preserves the single host behaviour the row inserter depends on:
    /// an inserted worksheet row still moves the shape down with its cell (measured
    /// 2026-10-03, <c>scripts/probe-rowinsert-anchoring.ps1</c> measured the move; the
    /// sizing half is what this removes). Placement is re-applied on every update, so a
    /// workbook whose shapes were created under the old rule heals on its next Refresh.
    /// </para>
    /// <para>
    /// <b>Why the host-owned class gets the opposite answer (ADR-0052 D6–D7, amended
    /// 2026-11-08).</b> A primitive that spans the plot's full vertical extent has a
    /// height the scene derives from the same measured rows Excel uses, so host resizing
    /// is not the renderer losing control — it is the chart following its own rows, which
    /// is what makes a delineator or a band stretch live on collapse/expand before any
    /// refresh runs. Verification accepts that geometry for this class only, with the
    /// wider height tolerance D8 states.
    /// </para>
    /// </remarks>
    private static void ApplyPlacement(Excel.Shape shape, bool hostOwnsExtent) =>
        shape.Placement = hostOwnsExtent
            ? Excel.XlPlacement.xlMoveAndSize
            : Excel.XlPlacement.xlMove;

    /// <summary>The placement a request's class requires in the host.</summary>
    private static Excel.XlPlacement RequiredPlacement(OfficeShapeRequest request) =>
        request.HostOwnsExtent ? Excel.XlPlacement.xlMoveAndSize : Excel.XlPlacement.xlMove;

    /// <summary>Whether the identifier names a splitter band primitive.</summary>
    /// <param name="primitiveId">The stable primitive identifier.</param>
    /// <returns><see langword="true"/> for a splitter band.</returns>
    private static bool IsSplitterBand(string primitiveId) =>
        primitiveId.EndsWith(":splitter-band", StringComparison.Ordinal);

    /// <summary>
    /// Describes the band's live placement and geometry before a write, or
    /// <see langword="null"/> when this shape is not traced.
    /// </summary>
    /// <param name="shape">The live shape.</param>
    /// <param name="primitiveId">The stable primitive identifier.</param>
    /// <returns>The description, or null.</returns>
    private string? DescribeBandShape(Excel.Shape shape, string primitiveId) =>
        _trace is null || !IsSplitterBand(primitiveId) ? null : BandGeometry(shape);

    /// <summary>
    /// Writes the band's live placement and geometry around one create or update, so a
    /// live log can tell "the host ignored the height we wrote" from "the scene asked
    /// for the wrong height".
    /// </summary>
    /// <param name="action">The write kind, for the log.</param>
    /// <param name="shape">The live shape after the write.</param>
    /// <param name="primitiveId">The stable primitive identifier.</param>
    /// <param name="before">The description captured before the write, or null.</param>
    private void TraceBandShape(string action, Excel.Shape shape, string primitiveId, string? before)
    {
        if (_trace is null || !IsSplitterBand(primitiveId))
        {
            return;
        }

        _trace(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.shape write={action} id={primitiveId} before=[{before ?? "n/a"}] after=[{BandGeometry(shape)}]"));
    }

    /// <summary>
    /// Logs one inventory line for a chart-shaped or untagged shape, so an artifact the
    /// reconciler cannot see is identifiable rather than inferred from the pixels.
    /// </summary>
    /// <param name="shape">The live shape.</param>
    /// <param name="name">The shape's name.</param>
    /// <param name="ownedTag">Whether the shape carries a valid ownership tag.</param>
    private void TraceInventoryShape(Excel.Shape shape, string name, bool ownedTag)
    {
        if (_trace is null)
        {
            return;
        }

        _trace(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.shape inventory name={name} owned={ownedTag} place={shape.Placement} l={shape.Left:0.###} t={shape.Top:0.###} w={shape.Width:0.###} h={shape.Height:0.###}"));
    }

    /// <summary>Formats a shape's placement, top and height as one invariant fragment.</summary>
    /// <param name="shape">The live shape.</param>
    /// <returns>The fragment.</returns>
    private static string BandGeometry(Excel.Shape shape) =>
        string.Create(
            CultureInfo.InvariantCulture,
            $"place={shape.Placement} t={shape.Top:0.###} h={shape.Height:0.###}");

    // ---- Test seams (internal virtual, per the ExcelWorkbookInitialiser pattern) ----

}
