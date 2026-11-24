using System.Diagnostics;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// The state one shape-write batch accumulates: the hoisted shapes collection, the proxy
/// and tag memos, the template cache, and the per-step create timings.
/// </summary>
/// <remarks>
/// <para>
/// One instance lives for the writer's whole life and is opened and closed by
/// <see cref="Open"/> and <see cref="Close"/>, so the fields stay non-nullable and a
/// caller reads them exactly as it read the writer's own fields before the extraction.
/// Opening resets every member; closing reports the timing record and clears them.
/// </para>
/// <para>
/// The class holds no COM proxy of its own beyond the shapes the writer resolved; it never
/// calls the host and never consults a protection guard. <see cref="EnsureNameIndex"/>
/// reaches the host only through the two seams the writer passes in, which is what keeps
/// the writer's test subclasses on the path.
/// </para>
/// </remarks>
internal sealed class ShapeWriteBatch(Action<string>? trace)
{
    /// <summary>Whether a batch is open (ADR-0054).</summary>
    internal bool Active { get; private set; }

    /// <summary>
    /// The shapes collection resolved once for the open batch, or <see langword="null"/>
    /// outside a batch or after a refused resolution.
    /// </summary>
    internal Excel.Shapes? Shapes { get; set; }

    /// <summary>
    /// Shape proxies resolved during the open batch, keyed by exact name. A name can still
    /// be absent, which the complete index proves is a true miss.
    /// </summary>
    internal Dictionary<string, Excel.Shape>? Proxies { get; set; }

    /// <summary>
    /// Ownership answers proved during the open batch, keyed by exact shape name. An owned
    /// tag cannot change while the batch runs, so one read per name is enough.
    /// </summary>
    internal Dictionary<string, bool> Tags { get; } = new(StringComparer.Ordinal);

    /// <summary>
    /// The batch's first shape per appearance, or <see langword="null"/> when the writer
    /// did not opt into template reuse.
    /// </summary>
    internal Dictionary<ExcelShapeWriter.ShapeTemplateKey, Excel.Shape>? Templates { get; set; }

    /// <summary>Whether the complete name index has been built for this batch.</summary>
    internal bool NameIndexed { get; set; }

    /// <summary>The batch's accumulated create timings (R5.18n).</summary>
    internal ShapeCreateTimings Timings { get; } = new();

    /// <summary>
    /// Opens the batch and resets every member it accumulates.
    /// </summary>
    /// <param name="reuseTemplates">
    /// Whether this batch may cache the first shape per appearance as a duplication
    /// template.
    /// </param>
    internal void Open(bool reuseTemplates)
    {
        Active = true;
        Shapes = null;
        Proxies = new Dictionary<string, Excel.Shape>(StringComparer.Ordinal);
        Tags.Clear();
        Templates = reuseTemplates ? [] : null;
        Timings.Reset();
        NameIndexed = false;
    }

    /// <summary>
    /// Closes the batch: reports the create-step record, then clears the state.
    /// </summary>
    /// <remarks>
    /// The report runs before the state is cleared and only when the caller asked for
    /// trace records: it is a measurement of the pass, not a product behaviour.
    /// </remarks>
    internal void Close()
    {
        ReportCreateSteps();
        Active = false;
        Shapes = null;
        Proxies = null;
        Tags.Clear();
        Templates = null;
        NameIndexed = false;
    }

    /// <summary>
    /// Builds the open batch's complete name index exactly once, so an ABSENT name can be
    /// answered without asking the host.
    /// </summary>
    /// <param name="shapes">The batch's resolved shapes collection.</param>
    /// <param name="count">The writer's shape-count seam.</param>
    /// <param name="shapeAt">The writer's one-based shape seam.</param>
    /// <remarks>
    /// <para>
    /// <b>Measured cost of not having this.</b> Excel's name indexer RAISES for a name it
    /// does not hold, and a COM exception costs ~26 ms to cross the boundary. The first
    /// live plot after Initialise writes 52 shapes onto a sheet with none, so every
    /// create paid a miss: <c>gantt.shape.create-steps n=52 lookup=1355.0 add=119.3
    /// style=31.6 ownership=13.2</c> — 1355 ms of a 1529 ms write, and 1984 ms of a
    /// click-to-plot the user reported as "feels like four seconds". Creating the
    /// shapes was never the expensive part.
    /// </para>
    /// <para>
    /// After one scan, a name that is not in the index is <em>proven</em> absent: events
    /// are off for the whole refresh (ADR-0020), so nothing can add a shape behind this
    /// pass. Creates add themselves through <see cref="Remember"/>, so a later lookup of
    /// a shape this batch created still hits.
    /// </para>
    /// </remarks>
    internal void EnsureNameIndex(
        Excel.Shapes shapes,
        Func<Excel.Shapes, int> count,
        Func<Excel.Shapes, int, Excel.Shape?> shapeAt)
    {
        // Held in a local: the seam calls below can invalidate the property's null state,
        // and the writer's own rule is that a proxy-bearing member is read once anyway.
        Dictionary<string, Excel.Shape>? proxies = Proxies;
        if (proxies is null || NameIndexed)
        {
            return;
        }

        int total = count(shapes);
        for (int index = 1; index <= total; index++)
        {
            Excel.Shape? shape = shapeAt(shapes, index);
            if (shape is null)
            {
                continue;
            }

            string name = shape.Name;
            if (!string.IsNullOrWhiteSpace(name))
            {
                _ = proxies.TryAdd(name, shape);
            }
        }

        NameIndexed = true;
    }

    /// <summary>Remembers a resolved proxy for the rest of the batch, when one is open.</summary>
    /// <param name="name">The exact shape name.</param>
    /// <param name="shape">The resolved shape, or <see langword="null"/>.</param>
    /// <returns>The shape, so a lookup can return it directly.</returns>
    internal Excel.Shape? Remember(string name, Excel.Shape? shape)
    {
        if (shape is not null && Proxies is not null)
        {
            Proxies[name] = shape;
        }

        return shape;
    }

    /// <summary>
    /// Writes one record splitting the batch's create cost by COM step, so the next
    /// attempt at making the first plot faster is aimed at the expensive step instead of
    /// at whichever one looks likely.
    /// </summary>
    private void ReportCreateSteps()
    {
        if (trace is null || Timings.Count == 0)
        {
            return;
        }

        static double Ms(long ticks)
        {
            return ticks * 1000d / Stopwatch.Frequency;
        }

        trace(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.shape.create-steps n={Timings.Count}"
            + $" lookup={Ms(Timings.Lookup):F1}"
            + $" add={Ms(Timings.Add):F1}"
            + $" style={Ms(Timings.Style):F1}"
            + $" ownership={Ms(Timings.Ownership):F1}"));
    }
}

/// <summary>
/// The accumulated create cost of one batch, in <see cref="Stopwatch"/> ticks, split by
/// COM step (R5.18n).
/// </summary>
/// <remarks>
/// Accumulated, never per shape, so a 52-shape pass still writes one record. The live
/// trace reports one <c>write=</c> figure for the whole pass, which cannot distinguish
/// "Excel is slow at AddShape" from "we issue too many round trips per shape".
/// </remarks>
internal sealed class ShapeCreateTimings
{
    /// <summary>Ticks spent proving a primitive name absent.</summary>
    internal long Lookup;

    /// <summary>Ticks spent in the host's shape creation call.</summary>
    internal long Add;

    /// <summary>Ticks spent writing the style.</summary>
    internal long Style;

    /// <summary>Ticks spent writing the ownership stamp.</summary>
    internal long Ownership;

    /// <summary>The number of creates the timings cover.</summary>
    internal int Count;

    /// <summary>Resets every member to zero.</summary>
    internal void Reset()
    {
        Lookup = 0;
        Add = 0;
        Style = 0;
        Ownership = 0;
        Count = 0;
    }
}
