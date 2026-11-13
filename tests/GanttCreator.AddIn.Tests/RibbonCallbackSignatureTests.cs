using System.Reflection;
using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The Excel-DNA callback binding contract: every callback attribute declared in
/// <c>Ribbon.xml</c> must name a method whose parameter list is exactly what
/// Excel passes for that control and attribute.
/// </summary>
/// <remarks>
/// This exists because a wrong arity compiles, passes every other ribbon test,
/// and then fails only in the live host with
/// "Custom UI Runtime Error ... Callback signature mismatch", which is invisible
/// to the offline gates. Observed live on 2026-11-13 when the Size dropdown
/// raised it for <c>OnSizeChange</c>.
/// <para>
/// Excel-DNA binds by name from the <c>on*</c>/<c>get*</c> attribute and then
/// checks the signature; a method with an extra injectable parameter (the
/// <c>RibbonStateService</c> seam the unit tests use) is a mismatch in exactly
/// that way. The seam therefore lives in a differently named method.
/// </para>
/// </remarks>
public class RibbonCallbackSignatureTests
{
    /// <summary>
    /// One callback binding rule: the element that carries the attribute, the
    /// attribute, and the parameter and return types Excel requires.
    /// </summary>
    private sealed record CallbackRule(
        string Element,
        string Attribute,
        Type[] Parameters,
        Type ReturnType);

    /// <summary>
    /// The rules for the callbacks this ribbon declares. Excel passes exactly
    /// these arguments; anything else is a binding failure in the host.
    /// </summary>
    private static readonly CallbackRule[] _rules =
    [
        new("customUI", "onLoad", [typeof(IRibbonUI)], typeof(void)),

        // Control-state getters: control in, value out.
        new("*", "getEnabled", [typeof(IRibbonControl)], typeof(bool)),
        new("*", "getVisible", [typeof(IRibbonControl)], typeof(bool)),
        new("*", "getPressed", [typeof(IRibbonControl)], typeof(bool)),
        new("*", "getLabel", [typeof(IRibbonControl)], typeof(string)),
        new("editBox", "getText", [typeof(IRibbonControl)], typeof(string)),
        new("dropDown", "getSelectedItemIndex", [typeof(IRibbonControl)], typeof(int)),

        // Dynamic item lists: count, then per-index id and label.
        new("dropDown", "getItemCount", [typeof(IRibbonControl)], typeof(int)),
        new("dropDown", "getItemID", [typeof(IRibbonControl), typeof(int)], typeof(string)),
        new("dropDown", "getItemLabel", [typeof(IRibbonControl), typeof(int)], typeof(string)),

        // Actions: the argument list depends on the control.
        new("button", "onAction", [typeof(IRibbonControl)], typeof(void)),
        new("toggleButton", "onAction", [typeof(IRibbonControl), typeof(bool)], typeof(void)),
        new("checkBox", "onAction", [typeof(IRibbonControl), typeof(bool)], typeof(void)),
        new("dropDown", "onAction",
            [typeof(IRibbonControl), typeof(string), typeof(int)], typeof(void)),
        new("editBox", "onChange", [typeof(IRibbonControl), typeof(string)], typeof(void)),
    ];

    /// <summary>
    /// Every callback attribute in the shipped XML, with the element that
    /// carries it, resolved through <c>GetCustomUI</c> so the test reads the
    /// same document Excel reads.
    /// </summary>
    private static List<(string Element, string Attribute, string Callback)> DeclaredCallbacks(
        string xml)
    {
        XDocument document = XDocument.Parse(xml);
        XNamespace ns = document.Root!.Name.Namespace;
        List<(string Element, string Attribute, string Callback)> declared = [];
        foreach (CallbackRule rule in _rules)
        {
            IEnumerable<XElement> elements = rule.Element == "*"
                ? document.Descendants()
                : document.Descendants(ns + rule.Element);
            foreach (XElement element in elements)
            {
                string? callback = (string?)element.Attribute(rule.Attribute);
                if (!string.IsNullOrWhiteSpace(callback))
                {
                    declared.Add((element.Name.LocalName, rule.Attribute, callback));
                }
            }
        }

        return declared;
    }

    private static bool IsCompatible(CallbackRule rule, MethodInfo method)
    {
        ParameterInfo[] parameters = method.GetParameters();
        if (parameters.Length != rule.Parameters.Length
            || method.ReturnType != rule.ReturnType)
        {
            return false;
        }

        for (int index = 0; index < parameters.Length; index++)
        {
            Type expected = rule.Parameters[index];
            Type actual = parameters[index].ParameterType;
            // Excel passes an IRibbonControl; `object` is equally acceptable and
            // is what several existing callbacks declare.
            bool controlPosition = index == 0 && expected == typeof(IRibbonControl);
            if (actual != expected && !(controlPosition && actual == typeof(object)))
            {
                return false;
            }
        }

        return true;
    }

    private static List<string> MismatchedCallbacks(string xml)
    {
        MethodInfo[] methods = typeof(GanttRibbon).GetMethods(
            BindingFlags.Public | BindingFlags.NonPublic
            | BindingFlags.Instance | BindingFlags.Static);
        List<string> mismatches = [];
        foreach ((string element, string attribute, string callback) in DeclaredCallbacks(xml))
        {
            CallbackRule rule = _rules.First(candidate =>
                candidate.Attribute == attribute
                && (candidate.Element == element || candidate.Element == "*"));
            bool satisfied = methods
                .Where(method => method.Name == callback)
                .Any(method => IsCompatible(rule, method));
            if (!satisfied)
            {
                mismatches.Add(
                    $"<{element} {attribute}=\"{callback}\"> expects "
                    + $"({string.Join(", ", rule.Parameters.Select(p => p.Name))}) -> "
                    + rule.ReturnType.Name);
            }
        }

        return mismatches;
    }

    [Fact]
    public void Every_declared_callback_has_the_signature_excel_calls()
    {
        GanttRibbon ribbon = new();
        string xml = ribbon.GetCustomUI("Microsoft.Excel.Workbook");

        List<string> mismatches = MismatchedCallbacks(xml);

        // The message lists every violation: the host reports one at a time, and
        // a truncated assertion would hide the rest of the same defect class.
        Assert.True(
            mismatches.Count == 0,
            $"{mismatches.Count} callback(s) do not match the signature Excel calls:"
            + Environment.NewLine
            + string.Join(Environment.NewLine, mismatches));
    }

    [Fact]
    public void The_signature_validator_flags_a_callback_with_an_extra_parameter()
    {
        // Positive control: the live defect. A dropdown action that takes the
        // state-service seam is exactly what the host rejected, and the
        // validator must name it rather than accept it.
        MethodInfo method = typeof(SeamShapedCallbacks).GetMethod(
            nameof(SeamShapedCallbacks.OnSizeChangeWithSeam),
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)!;
        CallbackRule dropDownAction = _rules.First(rule =>
            rule.Element == "dropDown" && rule.Attribute == "onAction");

        Assert.False(IsCompatible(dropDownAction, method));
    }

    [Fact]
    public void The_signature_validator_accepts_the_correct_dropdown_signature()
    {
        // Positive control the other way: the canonical three-argument dropdown
        // action must be accepted, or the failing test above would pass for the
        // wrong reason.
        MethodInfo method = typeof(SeamShapedCallbacks).GetMethod(
            nameof(SeamShapedCallbacks.OnSizeChange),
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)!;
        CallbackRule dropDownAction = _rules.First(rule =>
            rule.Element == "dropDown" && rule.Attribute == "onAction");

        Assert.True(IsCompatible(dropDownAction, method));
    }

    /// <summary>
    /// Two shapes of the same callback so the validator can be proven in both
    /// directions without depending on production code staying wrong.
    /// </summary>
    private static class SeamShapedCallbacks
    {
        internal static void OnSizeChange(
            IRibbonControl control,
            string selectedId,
            int selectedIndex)
        {
            _ = control;
            _ = selectedId;
            _ = selectedIndex;
        }

        internal static void OnSizeChangeWithSeam(
            IRibbonControl control,
            RibbonStateService stateService)
        {
            _ = control;
            _ = stateService;
        }
    }
}
