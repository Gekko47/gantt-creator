using GanttCreator.Office;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The workbook-state handler is the only place that can learn the active workbook changed, so it
/// is the only place that can make ADR-0055's arming memory wrong-but-safe.
/// </summary>
/// <remarks>
/// The state these tests touch is deliberately the process-wide instance the production refresh
/// shares. That is the thing under test: a private copy would prove that <c>Invalidate</c> works,
/// which is already obvious from its one-line body, and not that the add-in actually calls it.
/// </remarks>
public class ProtectionArmingResetTests
{
    [Fact]
    public void Resetting_the_arming_forgets_it_even_after_the_add_in_armed_protection()
    {
        ProtectionArmingState.Shared.MarkArmed();
        try
        {
            RefreshSheetCommand.ResetProtectionArming();

            Assert.False(ProtectionArmingState.Shared.IsArmed);
        }
        finally
        {
            // Leave the shared instance as the rest of the suite expects to find it: unarmed, which
            // is also what a fresh Excel session starts with.
            ProtectionArmingState.Shared.Invalidate();
        }
    }

    [Fact]
    public void Resetting_the_arming_is_harmless_when_nothing_was_armed()
    {
        ProtectionArmingState.Shared.Invalidate();

        RefreshSheetCommand.ResetProtectionArming();

        Assert.False(ProtectionArmingState.Shared.IsArmed);
    }
}
