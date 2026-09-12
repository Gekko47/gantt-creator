using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;

// CA2000: test fixtures are deliberately not disposed — the code under test
// (GanttRibbon) does not own disposable resources, and these are plain strings
// / reflection results. Scoped to this file.
#pragma warning disable CA2000

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for <see cref="GanttRibbon"/>: the Excel-DNA reflection
/// contract (public, ComVisible, ExcelRibbon-derived, parameterless
/// constructor), the <see cref="GetCustomUI"/> behaviour, and the RibbonX
/// namespace / IDs / callback contract. None of these require Excel.
/// </summary>
public class GanttRibbonTests
{
    private const string WorkbookRibbonId = "Microsoft.Excel.Workbook";

    private static readonly GanttRibbon Ribbon = new();

    // The set of RibbonX attribute names that dispatch to a callback method on
    // the ribbon class. Any of these appearing in the document must resolve to a
    // public method with a compatible signature (the callback contract).
    private static readonly HashSet<string> CallbackAttributeNames =
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
    ];

    private const string NamespaceCustomUI2010 =
        "http://schemas.microsoft.com/office/2009/07/customui";

    [Fact]
    public void GanttRibbon_is_public_ComVisible_ExcelRibbon_derived_and_has_parameterless_ctor()
    {
        Type type = typeof(GanttRibbon);

        Assert.True(type.IsPublic, "Excel-DNA requires a public ribbon type.");
        ComVisibleAttribute? ComVisibleAttribute =
            type.GetCustomAttribute<ComVisibleAttribute>(inherit: true);
        Assert.True(
            ComVisibleAttribute is not null && ComVisibleAttribute.Value == true,
            "The ribbon type must be [ComVisible(true)] for COM registration.");
        Assert.True(
            typeof(ExcelRibbon).IsAssignableFrom(type),
            "Excel-DNA only auto-registers ExcelRibbon descendants (IsRibbonType).");
        System.Reflection.ConstructorInfo? ctor = type.GetConstructor(Type.EmptyTypes);
        Assert.NotNull(ctor);
        Assert.True(
            ctor!.IsPublic,
            "Excel-DNA instantiates the ribbon via Activator.CreateInstance, which requires a public parameterless constructor.");
    }

    [Fact]
    public void GetCustomUI_returns_the_ribbon_xml_for_the_workbook_id()
    {
        string? xml = Ribbon.GetCustomUI(WorkbookRibbonId);

        Assert.NotNull(xml);
        Assert.NotEmpty(xml);
        Assert.Contains("customUI", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void GetCustomUI_returns_null_for_any_other_ribbon_id()
    {
        Assert.Null(Ribbon.GetCustomUI("Microsoft.Excel.Dataset"));
        Assert.Null(Ribbon.GetCustomUI(string.Empty));
        Assert.Null(Ribbon.GetCustomUI("Other"));
    }

    [Fact]
    public void Ribbon_xml_parses_and_uses_the_office_2009_namespace()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;

        XDocument doc = XDocument.Parse(xml);
        XElement? root = doc.Root;
        Assert.NotNull(root);
        Assert.Equal("customUI", root.Name.LocalName, ignoreCase: false);

        // The 2009/07 namespace targets Excel 2010+, the supported baseline.
        string ns = root.Name.NamespaceName;
        Assert.Equal(NamespaceCustomUI2010, ns);
    }

    [Fact]
    public void Ribbon_declares_exactly_one_gantt_creator_tab_with_expected_id_and_label()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);

        // Elements live in the CustomUI 2009/07 default namespace, so queries
        // must be namespace-qualified to match.
        XNamespace ns = NamespaceCustomUI2010;
        List<XElement> tabs = doc.Descendants(ns + "tab").ToList();
        Assert.Single(tabs);

        XElement tab = tabs[0];
        Assert.Equal(
            "tabGanttCreator",
            tab.Attribute("id")?.Value,
            ignoreCase: false);
        Assert.Equal(
            "Gantt Creator",
            tab.Attribute("label")?.Value,
            ignoreCase: false);
    }

    [Fact]
    public void Callback_contract_all_callbacks_in_the_real_xml_resolve_to_public_methods()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;

        List<string> unresolved = ValidateCallbacks(xml, typeof(GanttRibbon));

        Assert.Empty(unresolved);
    }

    [Fact]
    public void Callback_contract_reports_a_missing_callback_method_positive_test()
    {
        // A deliberately broken document whose onAction references a method that
        // does not exist on the ribbon type. Proves the validator actually
        // catches violations rather than vacuously passing.
        const string brokenXml = @"<?xml version='1.0' encoding='utf-8'?>
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabGanttCreator' label='Gantt Creator'>
        <group id='grpPlaceholder' label='Gantt Creator'>
          <button id='btnBroken' label='Broken' onAction='DoesNotExist'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        List<string> unresolved = ValidateCallbacks(brokenXml, typeof(GanttRibbon));

        Assert.Contains("DoesNotExist", unresolved);
    }

    /// <summary>
    /// Parses the RibbonX <paramref name="xml"/> and returns the callback method
    /// names (the values of known callback attributes) that do not resolve to a
    /// public instance method on <paramref name="ribbonType"/> with a compatible
    /// signature: zero parameters, or one parameter of type
    /// <see cref="IRibbonControl"/> or <see cref="IRibbonUI"/>.
    /// </summary>
    /// <param name="xml">The RibbonX document to validate.</param>
    /// <param name="ribbonType">The ribbon class type.</param>
    /// <returns>The set of unresolved callback method names, if any.</returns>
    private static List<string> ValidateCallbacks(
        string xml, Type ribbonType)
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

    private static bool CallbackResolves(Type ribbonType, string methodName)
    {
        System.Reflection.MethodInfo? method = ribbonType.GetMethod(
            methodName,
            BindingFlags.Instance | BindingFlags.Public);

        if (method is null)
        {
            return false;
        }

        System.Reflection.ParameterInfo[] parameters = method.GetParameters();
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

        return false;
    }
}
