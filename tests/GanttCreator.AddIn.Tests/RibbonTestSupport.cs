using System.Globalization;
using System.Reflection;
using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The fixtures and validators the ribbon contract facts share.
/// </summary>
/// <remarks>
/// Every member here needs only a ribbon XML document and the ribbon type, so the topic classes
/// share one copy instead of carrying their own. A test class reaches them with
/// <c>using static</c>, which leaves every call site unchanged.
/// </remarks>
internal static class RibbonTestSupport
{
    /// <summary>The workbook ribbon id the add-in answers to.</summary>
    internal const string WorkbookRibbonId = "Microsoft.Excel.Workbook";

    /// <summary>The RibbonX 2009 namespace.</summary>
    internal const string NamespaceCustomUI2010 =
        "http://schemas.microsoft.com/office/2009/07/customui";

    /// <summary>One ribbon instance for the read-only contract facts.</summary>
    internal static readonly GanttRibbon Ribbon = new();

    /// <summary>
    /// The RibbonX attribute names that dispatch to a callback method on the ribbon class. Any of
    /// these in the document must resolve to a public method with a compatible signature.
    /// </summary>
    internal static readonly HashSet<string> CallbackAttributeNames =
    [
        "onLoad",
        "onAction",
        "onChange",
        "getPressed",
        "getEnabled",
        "getVisible",
        "getImage",
        "getLabel",
        "getSize",
        "getScreentip",
        "getSupertip",
        "getDescription",
        "getKeyTip",
        "getText",
        "getContent",
        "getShowLabel",
        "getShowImage",
        "getSelectedItemIndex",

        // R5.5: the dynamic dropDown item callbacks. Excel calls getItemCount
        // first, then getItemID/getItemLabel for each index, then
        // getSelectedItemIndex — all four must resolve to a public method.
        "getItemCount",
        "getItemID",
        "getItemLabel",
    ];

    private static int _nextPlotBoundsId;

    /// <summary>Groups that do not declare the <c>getVisible</c> callback (R5.16).</summary>
    /// <param name="xml">The RibbonX document to validate.</param>
    /// <returns>The ids of the offending groups.</returns>
    internal static IReadOnlyList<string> GroupsMissingVisibleCallback(string xml)
    {
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = doc.Root!.Name.Namespace;
        return
        [
            .. doc.Descendants(ns + "group")
                .Where(group => (string?)group.Attribute("getVisible") != "GetVisible")
                .Select(group => (string?)group.Attribute("id") ?? string.Empty),
        ];
    }

    /// <summary>Buttons that carry neither a static label nor a projected one (R5.16).</summary>
    /// <param name="xml">The RibbonX document to validate.</param>
    /// <returns>The ids of the offending buttons.</returns>
    internal static IReadOnlyList<string> ButtonsMissingVisibleLabel(string xml)
    {
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;
        return
        [
            .. doc.Descendants(ns + "button")
                .Where(button => string.IsNullOrWhiteSpace(button.Attribute("label")?.Value)
                    && string.IsNullOrWhiteSpace(button.Attribute("getLabel")?.Value))
                .Select(button => (string?)button.Attribute("id") ?? string.Empty),
        ];
    }

    /// <summary>The control ids the service invalidated, in call order.</summary>
    /// <param name="ribbon">The mocked ribbon handle.</param>
    /// <returns>The invalidated ids.</returns>
    internal static List<string> RibbonInvalidatedIds(Mock<IRibbonUI> ribbon) =>
    [
        .. ribbon.Invocations
            .Where(invocation => invocation.Method.Name == nameof(IRibbonUI.InvalidateControl))
            .Select(invocation => (string)invocation.Arguments[0]!),
    ];

    /// <summary>One table body row carrying the supplied plot bounds.</summary>
    /// <param name="row">The row number.</param>
    /// <param name="start">The start date as <c>dd/MM/yyyy</c>.</param>
    /// <param name="finish">The finish date as <c>dd/MM/yyyy</c>.</param>
    /// <returns>The row.</returns>
    internal static GanttRowDto PlotBoundsRow(int row, string start, string finish) =>
        new(
            row,
            PlotBoundsId(),
            PlotBoundsId(),
            null,
            "As-Planned Activity",
            $"Row {row}",
            DateOnly.ParseExact(start, "dd/MM/yyyy", CultureInfo.InvariantCulture),
            DateOnly.ParseExact(finish, "dd/MM/yyyy", CultureInfo.InvariantCulture),
            null,
            null,
            null,
            null,
            null,
            true,
            null);

    /// <summary>
    /// Every callback attribute in the document whose method does not resolve to a public method
    /// with a compatible signature.
    /// </summary>
    /// <param name="xml">The RibbonX document to validate.</param>
    /// <param name="ribbonType">The ribbon class type.</param>
    /// <returns>The unresolved callback method names, if any.</returns>
    internal static List<string> ValidateCallbacks(string xml, Type ribbonType)
    {
        XDocument doc = XDocument.Parse(xml);
        List<string> unresolved = [];

        foreach (XElement element in doc.Descendants())
        {
            foreach (XAttribute attribute in element.Attributes())
            {
                if (!CallbackAttributeNames.Contains(attribute.Name.LocalName))
                {
                    continue;
                }

                string methodName = attribute.Value;
                if (!CallbackResolves(ribbonType, methodName))
                {
                    unresolved.Add(methodName);
                }
            }
        }

        return unresolved;
    }

    /// <summary>Whether a callback method name resolves to a compatible public method.</summary>
    /// <param name="ribbonType">The ribbon class type.</param>
    /// <param name="methodName">The method name the XML names.</param>
    /// <returns><see langword="true"/> when Excel can dispatch it.</returns>
    private static bool CallbackResolves(Type ribbonType, string methodName)
    {
        MethodInfo? method = ribbonType.GetMethod(
            methodName,
            BindingFlags.Instance | BindingFlags.Public);

        if (method is null)
        {
            return false;
        }

        ParameterInfo[] parameters = method.GetParameters();
        if (parameters.Length == 0)
        {
            return true;
        }

        if (parameters.Length == 1)
        {
            Type paramType = parameters[0].ParameterType;
            return paramType == typeof(IRibbonControl)
                || paramType == typeof(IRibbonUI);
        }

        if (parameters.Length == 2)
        {
            // The RibbonX editBox onChange callback passes the control plus the
            // edited text: public void OnChange(IRibbonControl, string).
            if (parameters[0].ParameterType == typeof(IRibbonControl)
                && parameters[1].ParameterType == typeof(string))
            {
                return true;
            }

            // The RibbonX dynamic dropDown item callbacks (R5.5) pass the
            // control plus the zero-based item index:
            // public string GetItemID(IRibbonControl, int) and
            // public string GetItemLabel(IRibbonControl, int). Excel
            // dispatches both by reflection, so a missing int parameter
            // throws at dropdown-population time.
            if (parameters[0].ParameterType == typeof(IRibbonControl)
                && parameters[1].ParameterType == typeof(int))
            {
                return true;
            }

            // The RibbonX checkBox onAction callback passes the control plus
            // the pressed state: public void OnClick(IRibbonControl, bool).
            return parameters[0].ParameterType == typeof(IRibbonControl)
                && parameters[1].ParameterType == typeof(bool);
        }

        if (parameters.Length == 3)
        {
            // The RibbonX dropDown onAction callback passes the control, the
            // selected item id, and the selected index:
            // public void OnChange(IRibbonControl, string, int). Excel
            // dispatches all three by reflection, so a two-parameter handler
            // throws TargetParameterCountException at click time (live defect,
            // found in Excel 2026-10-10).
            return parameters[0].ParameterType == typeof(IRibbonControl)
                && parameters[1].ParameterType == typeof(string)
                && parameters[2].ParameterType == typeof(int);
        }

        return false;
    }

    /// <summary>A valid identity for a seeded table row; the counter keeps them unique.</summary>
    /// <returns>The next row identity.</returns>
    private static string PlotBoundsId() => "G-" + (++_nextPlotBoundsId).ToString("x32", CultureInfo.InvariantCulture);
}
