using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live probe for the calculation mode ADR-0020 D1 removed and D2 said must be
/// re-probed before reinstating.
/// </summary>
/// <remarks>
/// <para>
/// ADR-0020 dropped calculation mode because every typed access on the reference host
/// raised <c>DISP_E_TYPEMISMATCH</c>. R5.18n measured the cost of that omission on
/// 2026-11-14: the first render after Initialise wrote 52 shapes in 6626.6 ms with
/// automatic calculation on and 80.5 ms with it manual — 127 ms versus 1.5 ms per
/// shape — so the setting now has a real consumer.
/// </para>
/// <para>
/// <b>Both in-process paths were probed and both fail.</b> Three attempts, recorded
/// because a fourth would be the same evidence:
/// </para>
/// <list type="number">
/// <item>typed <c>application.Calculation</c> — <c>COMException 0x80020005</c>
/// (<c>DISP_E_TYPEMISMATCH</c>), exactly as ADR-0020 D1 recorded;</item>
/// <item>late-bound <c>InvokeMember</c> with the invariant culture —
/// <c>TYPE_E_INVDATAREAD 0x80028018</c>, the documented wrong-thread-locale binder
/// defect rather than a refusal of the member;</item>
/// <item>late-bound <c>InvokeMember</c> with the thread culture set to <c>en-US</c> —
/// no exception, but the value returned is <c>0x800A03EA</c>, which is not a member
/// of <see cref="Excel.XlCalculation"/>.</item>
/// </list>
/// <para>
/// The assertion is therefore the negative one: the member is unusable in-process on
/// this build, so a reinstatement cannot be written as a typed property access. The
/// process-lifetime measurement that motivated the probe is recorded on
/// <c>docs/adr/0020-application-state-scope-five-settings.md</c>; the speed work needs
/// a different mechanism.
/// </para>
/// </remarks>
public class ApplicationStateScopeIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_Calculation_member_is_unusable_in_process_on_either_access_form()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Excel.Application application = fixture.Excel;
            _output.WriteLine($"Excel {application.Version}");

            // 1. The typed PIA path: what ExcelApplicationStateScope would have to use.
            COMException typed = Assert.Throws<COMException>(() =>
            {
                _ = application.Calculation;
            });
            _output.WriteLine($"typed Calculation refused: 0x{typed.HResult:X8} {typed.Message}");
            Assert.Equal(unchecked((int)0x80020005), typed.HResult);

            // 2. The same member through IDispatch. IDispatch binding is
            //    locale-sensitive, so the culture is part of the call shape: the
            //    invariant culture is the one that produced TYPE_E_INVDATAREAD.
            CultureInfo originalCulture = CultureInfo.CurrentCulture;
            _output.WriteLine($"thread culture = {originalCulture.Name}");
            object? lateBound;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
                lateBound = application.GetType().InvokeMember(
                    "Calculation",
                    BindingFlags.GetProperty,
                    binder: null,
                    target: application,
                    args: null,
                    culture: CultureInfo.GetCultureInfo("en-US"));
            }
            finally
            {
                CultureInfo.CurrentCulture = originalCulture;
            }

            _output.WriteLine($"late-bound Calculation = {lateBound} (0x{Convert.ToInt32(lateBound, CultureInfo.InvariantCulture):X8})");
            Assert.NotNull(lateBound);

            // The returned value is an HRESULT, not a mode. Asserting the NEGATIVE is
            // deliberate: if a future Office build starts returning a real
            // XlCalculation this test fails, which is precisely the prompt ADR-0020 D2
            // asks for ("re-probe before reinstating").
            int returned = Convert.ToInt32(lateBound, CultureInfo.InvariantCulture);
            Assert.False(
                Enum.IsDefined(typeof(Excel.XlCalculation), returned),
                $"late-bound Calculation returned the enum member {(Excel.XlCalculation)returned}; the member is usable again and ADR-0020 D2 must be revisited.");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
