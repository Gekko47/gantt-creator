using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// The staging adapters' shared name lookup prefers the host's name indexer, so the
/// grouping preflight is one indexed lookup per staged name instead of a full collection
/// scan per name.
/// </summary>
/// <remarks>
/// Measured 2026-11-19 on a 20-activity export: the old scan cost 29.1 s of a 39.4 s copy
/// (340 lookups over a 438-shape collection). These facts pin the three outcomes the
/// indexer can produce, including the measured raising miss, and the degraded fallback
/// for a host or test double whose indexer resolves nothing without raising.
/// </remarks>
public sealed class StagingShapeAdapterLookupTests
{
    private const string _name = "GanttCreatorStaging:chart:bar:0";

    [Fact]
    public void An_indexed_hit_returns_the_shape_without_scanning_the_collection()
    {
        Mock<Excel.Shape> shape = new();
        LookupGrouper grouper = new(indexed: shape.Object);

        Excel.Shape? found = grouper.FindShapeByName(new Mock<Excel.Shapes>().Object, _name);

        Assert.Same(shape.Object, found);
        Assert.Equal(0, grouper.ScanReads);
    }

    [Fact]
    public void A_raising_indexed_miss_is_absent_without_scanning_the_collection()
    {
        LookupGrouper grouper = new(
            miss: new ArgumentException("The item with the specified name wasn't found."));

        Excel.Shape? found = grouper.FindShapeByName(new Mock<Excel.Shapes>().Object, _name);

        Assert.Null(found);
        Assert.Equal(0, grouper.ScanReads);
    }

    [Fact]
    public void An_invalid_operation_miss_is_absent_without_scanning_the_collection()
    {
        LookupGrouper grouper = new(miss: new InvalidOperationException("The item was not found."));

        Excel.Shape? found = grouper.FindShapeByName(new Mock<Excel.Shapes>().Object, _name);

        Assert.Null(found);
        Assert.Equal(0, grouper.ScanReads);
    }

    [Fact]
    public void A_degraded_indexer_falls_back_to_the_name_scan()
    {
        Mock<Excel.Shape> shape = new();
        _ = shape.SetupGet(candidate => candidate.Name).Returns(_name);
        LookupGrouper grouper = new(indexed: null, scanned: [shape.Object]);

        Excel.Shape? found = grouper.FindShapeByName(new Mock<Excel.Shapes>().Object, _name);

        Assert.Same(shape.Object, found);
        Assert.Equal(1, grouper.ScanReads);
    }

    /// <summary>
    /// A grouper whose only host interactions are the three lookup outcomes under test.
    /// </summary>
    private sealed class LookupGrouper : ExcelShapeGrouper
    {
        private readonly Excel.Shape? _indexed;
        private readonly Exception? _miss;
        private readonly IReadOnlyList<Excel.Shape> _scanned;

        internal LookupGrouper(
            Excel.Shape? indexed = null,
            Exception? miss = null,
            IReadOnlyList<Excel.Shape>? scanned = null)
            : base(application: null, protectionGuard: null)
        {
            _indexed = indexed;
            _miss = miss;
            _scanned = scanned ?? [];
        }

        internal int ScanReads { get; private set; }

        internal override Excel.Shape? GetShapeByExactName(Excel.Shapes shapes, string name) =>
            _miss is null ? _indexed : throw _miss;

        internal override int GetShapeCount(Excel.Shapes shapes) => _scanned.Count;

        internal override Excel.Shape? GetShapeAt(Excel.Shapes shapes, int index)
        {
            ScanReads++;
            return index >= 1 && index <= _scanned.Count ? _scanned[index - 1] : null;
        }
    }
}
