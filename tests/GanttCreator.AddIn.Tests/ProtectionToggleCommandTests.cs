using GanttCreator.Office;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// ADR-0055 clause 7: the ribbon carries one CHART toggle and one CELL toggle, and each command sets
/// the state the button has already moved to, re-applies protection from it, and leaves the other
/// toggle alone.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why these are command tests rather than port tests.</b> The port mapping is proved in
/// <c>ExcelWorksheetProtectionPortTests</c>; what can still go wrong here is the wiring — a command
/// that flips its own flag (inverting the user's intent when Excel has already changed the button), or
/// that lets the two toggles interfere. Both are cheap to pin without a host, and both were live
/// defects during this feature's development.
/// </para>
/// <para>
/// The live end-to-end gate (<c>ProtectionToggleIntegrationTests</c>) covers what only Excel can
/// answer: whether the sheet and the shapes actually end up locked.
/// </para>
/// </remarks>
public class ProtectionToggleCommandTests
{
    [Fact]
    public void Locking_the_chart_locks_every_owned_shape_and_records_the_protection()
    {
        var shapes = new FakeShapes("G-a:bar", "G-b:bar");
        var protection = new FakeProtection();
        var arming = new ProtectionArmingState();
        var preferences = new ProtectionPreferences();
        List<string> messages = [];

        ToggleChartLockCommand.Run(shapes, protection, arming, preferences, locked: true, messages.Add);

        Assert.True(preferences.ChartLocked);
        Assert.Equal(1, shapes.LockCount);
        Assert.Empty(shapes.Unlocked);
        Assert.Equal(1, protection.ProtectCount);
        Assert.True(arming.IsArmed);
        Assert.Empty(messages);
    }

    [Fact]
    public void Unlocking_the_chart_releases_every_owned_shape_and_records_the_protection()
    {
        var shapes = new FakeShapes("G-a:bar", "G-b:bar");
        var protection = new FakeProtection();
        var arming = new ProtectionArmingState();
        var preferences = new ProtectionPreferences();
        List<string> messages = [];

        ToggleChartLockCommand.Run(shapes, protection, arming, preferences, locked: false, messages.Add);

        Assert.False(preferences.ChartLocked);
        Assert.Equal(["G-a:bar", "G-b:bar"], shapes.Unlocked);
        Assert.Equal(0, shapes.LockCount);
        Assert.Equal(1, protection.ProtectCount);
        Assert.True(arming.IsArmed);
        Assert.Empty(messages);
    }

    /// <summary>
    /// RibbonX has already changed the button when it calls the handler, so the command must SET the
    /// state it is given. A command that flipped its own flag would invert the user's intent on every
    /// press, which is exactly the "Callback signature mismatch" turn's failure mode.
    /// </summary>
    [Fact]
    public void The_chart_toggle_sets_the_requested_state_instead_of_flipping()
    {
        var shapes = new FakeShapes("G-a:bar");
        var preferences = new ProtectionPreferences();
        List<string> messages = [];

        ToggleChartLockCommand.Run(shapes, new FakeProtection(), new ProtectionArmingState(), preferences, locked: false, messages.Add);
        ToggleChartLockCommand.Run(shapes, new FakeProtection(), new ProtectionArmingState(), preferences, locked: false, messages.Add);

        Assert.False(preferences.ChartLocked);
        Assert.Equal(["G-a:bar", "G-a:bar"], shapes.Unlocked);
    }

    [Fact]
    public void The_chart_toggle_leaves_the_cell_toggle_alone()
    {
        var preferences = new ProtectionPreferences { CellsUnlocked = true };
        List<string> messages = [];

        ToggleChartLockCommand.Run(
            new FakeShapes(), new FakeProtection(), new ProtectionArmingState(), preferences, locked: true, messages.Add);

        Assert.True(preferences.ChartLocked);
        Assert.True(preferences.CellsUnlocked);
    }

    [Fact]
    public void A_refused_chart_protect_is_reported_and_is_not_treated_as_our_own()
    {
        var protection = new FakeProtection { ProtectRefusal = ProtectionWriteRefusal.HostRejected };
        var arming = new ProtectionArmingState();
        var preferences = new ProtectionPreferences();
        List<string> messages = [];

        ToggleChartLockCommand.Run(
            new FakeShapes(), protection, arming, preferences, locked: true, messages.Add);

        // The preference is set first, so the button and the state agree even when the host refuses.
        Assert.True(preferences.ChartLocked);
        Assert.False(arming.IsArmed);
        Assert.Contains("could not be re-applied", Assert.Single(messages), StringComparison.Ordinal);
    }

    [Fact]
    public void A_refused_chart_protect_while_unlocking_uses_the_release_wording()
    {
        var protection = new FakeProtection { ProtectRefusal = ProtectionWriteRefusal.HostRejected };
        List<string> messages = [];

        ToggleChartLockCommand.Run(
            new FakeShapes(), protection, new ProtectionArmingState(), new ProtectionPreferences(), locked: false, messages.Add);

        Assert.Contains("could not be updated", Assert.Single(messages), StringComparison.Ordinal);
    }

    [Fact]
    public void Locking_the_cells_sets_the_preference_and_re_protects()
    {
        var preferences = new ProtectionPreferences { CellsUnlocked = true };
        var protection = new FakeProtection();
        var arming = new ProtectionArmingState();
        List<string> messages = [];

        ToggleCellLockCommand.Run(protection, arming, preferences, locked: true, messages.Add);

        Assert.False(preferences.CellsUnlocked);
        Assert.Equal(1, protection.ProtectCount);
        Assert.True(arming.IsArmed);
        Assert.Empty(messages);
    }

    [Fact]
    public void Unlocking_the_cells_sets_the_preference_and_re_protects()
    {
        var preferences = new ProtectionPreferences();
        var protection = new FakeProtection();
        var arming = new ProtectionArmingState();
        List<string> messages = [];

        ToggleCellLockCommand.Run(protection, arming, preferences, locked: false, messages.Add);

        Assert.True(preferences.CellsUnlocked);
        Assert.Equal(1, protection.ProtectCount);
        Assert.True(arming.IsArmed);
        Assert.Empty(messages);
    }

    [Fact]
    public void The_cell_toggle_sets_the_requested_state_instead_of_flipping()
    {
        var preferences = new ProtectionPreferences();
        List<string> messages = [];

        ToggleCellLockCommand.Run(new FakeProtection(), new ProtectionArmingState(), preferences, locked: false, messages.Add);
        ToggleCellLockCommand.Run(new FakeProtection(), new ProtectionArmingState(), preferences, locked: false, messages.Add);

        Assert.True(preferences.CellsUnlocked);
    }

    [Fact]
    public void The_cell_toggle_leaves_the_chart_toggle_alone()
    {
        var preferences = new ProtectionPreferences { ChartLocked = false };
        List<string> messages = [];

        ToggleCellLockCommand.Run(new FakeProtection(), new ProtectionArmingState(), preferences, locked: true, messages.Add);

        Assert.False(preferences.ChartLocked);
        Assert.False(preferences.CellsUnlocked);
    }

    [Theory]
    [InlineData(true, "could not be locked again")]
    [InlineData(false, "could not be released")]
    public void A_refused_cell_protect_is_reported_with_the_direction(bool locked, string expected)
    {
        var protection = new FakeProtection { ProtectRefusal = ProtectionWriteRefusal.HostRejected };
        var arming = new ProtectionArmingState();
        List<string> messages = [];

        ToggleCellLockCommand.Run(protection, arming, new ProtectionPreferences(), locked, messages.Add);

        Assert.False(arming.IsArmed);
        Assert.Contains(expected, Assert.Single(messages), StringComparison.Ordinal);
    }

    private sealed class FakeShapes(params string[] owned) : IShapeWritePort
    {
        public List<string> Unlocked { get; } = [];

        public int LockCount { get; private set; }

        public ShapeWriteRefusal? LockRefusal { get; init; }

        public ShapeWriteOutcome UnlockShape(string primitiveId)
        {
            Unlocked.Add(primitiveId);
            return ShapeWriteOutcome.Ok();
        }

        public ShapeWriteOutcome LockOwnedShapes()
        {
            LockCount++;
            return LockRefusal is { } refusal
                ? ShapeWriteOutcome.Refused(refusal)
                : ShapeWriteOutcome.Ok();
        }

        public ShapeWriteOutcome Create(OfficeShapeRequest request) => ShapeWriteOutcome.Ok();

        public ShapeWriteOutcome Update(OfficeShapeRequest request) => ShapeWriteOutcome.Ok();

        public ShapeWriteOutcome UpdateGeometry(OfficeShapeRequest request) => ShapeWriteOutcome.Ok();

        public ShapeWriteOutcome Delete(string primitiveId) => ShapeWriteOutcome.Ok();

        public IReadOnlyList<string> ListOwned() => owned;

        public bool GeometryMatches(OfficeShapeRequest request) => true;

        public ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront) => ShapeWriteOutcome.Ok();

        public IDisposable BeginBatch() => new InertScope();

        private sealed class InertScope : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }

    private sealed class FakeProtection : IWorksheetProtectionPort
    {
        public ProtectionWriteRefusal? ProtectRefusal { get; init; }

        public int ProtectCount { get; private set; }

        public WorksheetProtectionState? Read() => new(ContentsProtected: false, ShapesProtected: false);

        public ProtectionWriteOutcome Protect()
        {
            ProtectCount++;
            return ProtectRefusal is { } refusal
                ? ProtectionWriteOutcome.Refused(refusal)
                : ProtectionWriteOutcome.Ok();
        }

        public ProtectionWriteOutcome Unprotect() => ProtectionWriteOutcome.Ok();
    }
}
