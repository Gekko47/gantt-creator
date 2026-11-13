using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The second half of the live callback contract: not what Excel passes to a
/// handler, but whether the values a control is handed can be resolved.
/// </summary>
/// <remarks>
/// Two live defects on 2026-11-13 both lived here rather than in the signature:
/// a dropdown whose chosen item id was never forwarded, and a
/// <c>getSelectedItemIndex</c> that answered -1 for a control that has no such
/// state. <see cref="RibbonCallbackSignatureTests"/> pins the arity; this pins
/// the values, which is the half a signature check cannot see.
/// <para>
/// The rules are the ones Excel enforces: every item id the dropdown can send
/// must be accepted by that dropdown's <c>onAction</c> and land on the same
/// index the dropdown shows; and every index a list can report must resolve to a
/// non-empty, unique item.
/// </para>
/// </remarks>
public class RibbonDropdownContractTests
{
    private static readonly XNamespace _ns =
        XNamespace.Get("http://schemas.microsoft.com/office/2009/07/customui");

    /// <summary>
    /// The item ids of a dropDown, in the shipped XML's order, resolved through
    /// <c>GetCustomUI</c> so the test reads the document Excel reads.
    /// </summary>
    private static List<string> ItemIds(string dropDownId)
    {
        GanttRibbon ribbon = new();
        XDocument document = XDocument.Parse(ribbon.GetCustomUI("Microsoft.Excel.Workbook"));
        return
        [
            .. document.Descendants(_ns + "dropDown")
                .Where(dropDown => (string?)dropDown.Attribute("id") == dropDownId)
                .SelectMany(dropDown => dropDown.Elements(_ns + "item"))
                .Select(item => (string?)item.Attribute("id") ?? string.Empty),
        ];
    }

    private static RibbonStateService ArmedService()
    {
        // A service with no workbook, no table reader and no selection: the state
        // every getter must survive, because Excel asks before the user has done
        // anything. The catalogue writer is removed so nothing reaches COM.
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        return service;
    }

    [Fact]
    public void Every_margin_item_id_lands_on_its_own_index()
    {
        List<string> ids = ItemIds(RibbonControlIds.Margin);
        Assert.NotEmpty(ids);
        var service = ArmedService();

        foreach ((string id, int expectedIndex) in ids.Select((id, index) => (id, index)))
        {
            GanttRibbon.OnMarginChange(null, id, service);

            // The id the dropdown sends must be the key the handler accepts: a
            // mismatch here is the "selecting an item does nothing" defect.
            Assert.Equal(expectedIndex, GanttRibbon.GetMargin(null, service));
        }
    }

    [Fact]
    public void Every_time_scale_item_id_lands_on_its_own_index()
    {
        List<string> ids = ItemIds(RibbonControlIds.PlotTimeScale);
        Assert.NotEmpty(ids);
        var service = ArmedService();

        foreach ((string id, int expectedIndex) in ids.Select((id, index) => (id, index)))
        {
            GanttRibbon.OnTimeScaleChange(null, id, service);

            Assert.Equal(expectedIndex, GanttRibbon.GetTimeScale(null, service));
        }
    }

    [Fact]
    public void Every_size_item_id_lands_on_its_own_index()
    {
        List<string> ids = ItemIds(RibbonControlIds.SizeDropdown);
        Assert.NotEmpty(ids);
        var service = ArmedService();

        foreach ((string id, int expectedIndex) in ids.Select((id, index) => (id, index)))
        {
            GanttRibbon.OnSizeChange(null, id, service);

            Assert.Equal(expectedIndex, GanttRibbon.GetSize(null, service));
        }
    }

    /// <summary>
    /// The dynamic lists must be coherent before any workbook, table or
    /// selection exists: a non-negative count, a selected index that lies inside
    /// that count, and a non-empty unique id for every index the count promises.
    /// A list that reports three items but an empty id at index 2 hands Excel a
    /// value it cannot resolve.
    /// </summary>
    [Fact]
    public void The_dynamic_dropdowns_are_coherent_without_a_selection()
    {
        var service = ArmedService();

        AssertListIsCoherent(
            "description",
            () => GanttRibbon.GetDescriptionLabelCount(null, service),
            index => GanttRibbon.GetDescriptionLabelItemID(null, index, service),
            index => GanttRibbon.GetDescriptionLabelItemLabel(null, index, service),
            () => GanttRibbon.GetDescriptionLabelSelectedIndex(null, service));

        AssertListIsCoherent(
            "date",
            () => GanttRibbon.GetDateLabelCount(null, service),
            index => GanttRibbon.GetDateLabelItemID(null, index, service),
            index => GanttRibbon.GetDateLabelItemLabel(null, index, service),
            () => GanttRibbon.GetDateLabelSelectedIndex(null, service));

        AssertListIsCoherent(
            "custom style",
            () => GanttRibbon.GetCustomStyleCount(null, service),
            index => GanttRibbon.GetCustomStyleItemID(null, index, service),
            index => GanttRibbon.GetCustomStyleItemLabel(null, index, service),
            () => GanttRibbon.GetCustomStyleSelectedIndex(null, service));
    }

    private static void AssertListIsCoherent(
        string name,
        Func<int> count,
        Func<int, string> itemId,
        Func<int, string> itemLabel,
        Func<int> selectedIndex)
    {
        int itemCount = count();
        Assert.True(itemCount >= 0, $"{name}: a count cannot be negative, read {itemCount}.");

        int selected = selectedIndex();
        Assert.True(
            selected >= -1,
            $"{name}: the selected index must be -1 or a real index, read {selected}.");
        if (itemCount == 0)
        {
            // Excel tolerates "nothing selected" only when there is nothing to
            // select, and it does not ask for item ids in that state.
            return;
        }

        Assert.True(
            selected < itemCount,
            $"{name}: selected index {selected} is outside the {itemCount} item(s) offered.");

        List<string> ids = [.. Enumerable.Range(0, itemCount).Select(itemId)];
        Assert.All(ids, id => Assert.False(
            string.IsNullOrWhiteSpace(id),
            $"{name}: an item id in 0..{itemCount - 1} is empty, so Excel cannot resolve it."));
        Assert.Equal(ids.Count, ids.Distinct(StringComparer.Ordinal).Count());

        List<string> labels = [.. Enumerable.Range(0, itemCount).Select(itemLabel)];
        Assert.All(labels, label => Assert.False(
            string.IsNullOrWhiteSpace(label),
            $"{name}: an item label in 0..{itemCount - 1} is empty."));
    }
}
