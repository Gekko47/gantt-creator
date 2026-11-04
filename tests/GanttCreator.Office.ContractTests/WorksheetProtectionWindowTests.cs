namespace GanttCreator.Office.ContractTests;

/// <summary>
/// The release window is what makes a protected sheet writable at all: the live measurement shows a
/// protected sheet refuses shape creation, movement and row insertion, so every write phase runs
/// inside one of these.
/// </summary>
/// <remarks>
/// The nesting tests are the ones that matter. A command opens the window, the refresh it triggers
/// opens it again, and only the outermost scope may re-protect — otherwise the refresh would lock the
/// sheet while the command that caused it still has a row to write.
/// </remarks>
public class WorksheetProtectionWindowTests
{
    [Fact]
    public void Opening_and_closing_an_armed_window_releases_then_protects_once()
    {
        FakeWorksheetProtectionPort port = new();
        ProtectionArmingState state = new();
        state.MarkArmed();
        WorksheetProtectionWindow window = new(port, state);

        using (window.Open())
        {
            Assert.False(port.State!.ContentsProtected);
        }

        Assert.Equal(["Unprotect", "Protect"], port.Calls);
        Assert.True(port.State!.ContentsProtected);
        Assert.True(state.IsArmed);
    }

    [Fact]
    public void An_unarmed_window_makes_no_host_call()
    {
        FakeWorksheetProtectionPort port = new();
        ProtectionArmingState state = new();

        using (new WorksheetProtectionWindow(port, state).Open())
        {
        }

        Assert.Empty(port.Calls);
    }

    /// <summary>
    /// The first refresh of a session arms INSIDE the window, so the window has to protect on the way
    /// out even though nothing was armed on the way in.
    /// </summary>
    [Fact]
    public void A_window_that_arms_inside_itself_protects_when_it_closes()
    {
        FakeWorksheetProtectionPort port = new();
        ProtectionArmingState state = new();
        WorksheetProtectionWindow window = new(port, state);

        using (window.Open())
        {
            state.MarkArmed();
        }

        Assert.Equal(["Protect"], port.Calls);
        Assert.True(port.State!.ContentsProtected);
    }

    [Fact]
    public void A_nested_window_releases_and_protects_exactly_once()
    {
        FakeWorksheetProtectionPort port = new();
        ProtectionArmingState state = new();
        state.MarkArmed();
        WorksheetProtectionWindow window = new(port, state);

        using (window.Open())
        {
            using (window.Open())
            {
                Assert.False(port.State!.ContentsProtected);
            }

            // The inner scope closing must NOT re-protect: the outer one is still writing.
            Assert.False(port.State!.ContentsProtected);
        }

        Assert.Equal(["Unprotect", "Protect"], port.Calls);
        Assert.True(port.State!.ContentsProtected);
    }

    [Fact]
    public void Disposing_a_scope_twice_re_protects_only_once()
    {
        FakeWorksheetProtectionPort port = new();
        ProtectionArmingState state = new();
        state.MarkArmed();
        WorksheetProtectionWindow window = new(port, state);

        IDisposable scope = window.Open();
        scope.Dispose();
        scope.Dispose();

        Assert.Equal(["Unprotect", "Protect"], port.Calls);
    }

    /// <summary>
    /// A refused release is recorded and the write continues; the host will refuse the write itself,
    /// which is a truer signal than this method hiding it.
    /// </summary>
    [Fact]
    public void A_refused_release_is_recorded_and_the_window_still_protects_on_close()
    {
        FakeWorksheetProtectionPort port = new()
        {
            UnprotectRefusal = ProtectionWriteRefusal.HostRejected,
        };
        ProtectionArmingState state = new();
        state.MarkArmed();
        List<string> traces = [];
        WorksheetProtectionWindow window = new(port, state, traces.Add);

        using (window.Open())
        {
        }

        Assert.Contains(
            traces,
            line => line.Contains("window release refused=HostRejected", StringComparison.Ordinal));
        Assert.Equal(["Unprotect", "Protect"], port.Calls);
    }

    /// <summary>
    /// Arming is invalidated while the window is open (the user turned the setting off), so the close
    /// must not re-protect a sheet the user now owns.
    /// </summary>
    [Fact]
    public void A_window_that_loses_the_arming_does_not_protect_on_close()
    {
        FakeWorksheetProtectionPort port = new();
        ProtectionArmingState state = new();
        state.MarkArmed();
        WorksheetProtectionWindow window = new(port, state);

        using (window.Open())
        {
            state.Invalidate();
        }

        Assert.Equal(["Unprotect"], port.Calls);
        Assert.False(port.State!.ContentsProtected);
    }
}
