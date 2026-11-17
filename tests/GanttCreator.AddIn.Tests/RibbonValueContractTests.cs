namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The value half of the ribbon contract for the controls that are not
/// dropdowns: the text readouts and the pressed-state toggles.
/// </summary>
/// <remarks>
/// The live defects of 2026-11-13 were both about a callback handing Excel a
/// value it could not use. The same four questions apply here: can a
/// <c>getText</c>/<c>getLabel</c> callback return null, and does a
/// <c>getPressed</c> callback read back exactly what its <c>onAction</c> just
/// set? A checkbox that wrote one thing and reported another looks broken while
/// working perfectly, which is worse than an error.
/// </remarks>
public class RibbonValueContractTests
{
    /// <summary>
    /// A service with no workbook, no reader and no selection: the state Excel
    /// asks about before the user has done anything, and the state the -1 defect
    /// was found in.
    /// </summary>
    private static RibbonStateService ArmedService()
    {
        RibbonStateService service = new();
        service.SetCatalogueWriter(null);

        // The setting writers report a failed persist through the command
        // boundary, whose production presenter is a modal dialog. Installing a
        // capturing boundary is what makes these paths runnable headlessly at
        // all: without it, a failing persist waits forever for a dialog nobody
        // can dismiss. Observed live on 2026-11-13, when this file's first
        // version hung the test host for hours.
        service.SetCommandBoundary(new CommandBoundary(presenter: _ => { }));
        return service;
    }

    /// <summary>
    /// Every string the ribbon puts in a control must be a string. A null from a
    /// <c>get*</c> callback is the same class of defect as a negative index: the
    /// host has been handed something it cannot render.
    /// </summary>
    [Fact]
    public void Every_text_and_label_callback_returns_a_string_without_a_workbook()
    {
        RibbonStateService service = ArmedService();

        Assert.NotNull(GanttRibbon.GetPlotStartDate(null, service));
        Assert.NotNull(GanttRibbon.GetPlotFinishDate(null, service));
        Assert.NotNull(GanttRibbon.GetMarginCm(null, service));
        Assert.NotNull(GanttRibbon.GetPlotWidthLabel(null, service));
        Assert.NotNull(GanttRibbon.GetPlotHeightLabel(null, service));
    }

    /// <summary>
    /// The two Auto checkboxes: the pressed state Excel reads must be the state
    /// the click just stored. A getter that lags or inverts makes the checkbox
    /// spring back the moment it is clicked.
    /// </summary>
    [Fact]
    public void The_auto_checkboxes_read_back_exactly_what_they_set()
    {
        RibbonStateService service = ArmedService();

        GanttRibbon.OnPlotStartAutoClick(null, true, service);
        Assert.True(GanttRibbon.GetPlotStartAuto(null, service));

        GanttRibbon.OnPlotStartAutoClick(null, false, service);
        Assert.False(GanttRibbon.GetPlotStartAuto(null, service));

        GanttRibbon.OnPlotFinishAutoClick(null, true, service);
        Assert.True(GanttRibbon.GetPlotFinishAuto(null, service));

        GanttRibbon.OnPlotFinishAutoClick(null, false, service);
        Assert.False(GanttRibbon.GetPlotFinishAuto(null, service));
    }

    /// <summary>
    /// The two Auto checkboxes are independent settings, not one shared flag: a
    /// pair that moved together would silently re-date the other end of the
    /// chart whenever either was cleared.
    /// </summary>
    [Fact]
    public void The_two_auto_checkboxes_are_independent()
    {
        RibbonStateService service = ArmedService();
        GanttRibbon.OnPlotStartAutoClick(null, true, service);
        GanttRibbon.OnPlotFinishAutoClick(null, true, service);

        GanttRibbon.OnPlotStartAutoClick(null, false, service);

        Assert.False(GanttRibbon.GetPlotStartAuto(null, service));
        Assert.True(GanttRibbon.GetPlotFinishAuto(null, service));
    }
}
