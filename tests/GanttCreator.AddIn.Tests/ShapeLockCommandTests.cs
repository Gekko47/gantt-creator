using GanttCreator.Office;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// ADR-0055 clause 6: Unlock acts on one owned selection and suspends the automatic refresh; Lock
/// shapes restores both the lock and the protection.
/// </summary>
public class ShapeLockCommandTests
{
    private const string Id = "G-0123456789abcdef0123456789abcdef:bar";

    [Fact]
    public void Unlocking_records_the_shape_releases_the_sheet_and_disarms_the_next_arm()
    {
        var selection = new FakeSelection(SelectedShapeOutcome.Ok(Id));
        var shapes = new FakeShapes();
        var protection = new FakeProtection();
        var unlocks = new ShapeUnlockState();
        var arming = new ProtectionArmingState();
        arming.MarkArmed();
        List<string> messages = [];

        UnlockShapeCommand.Run(selection, shapes, protection, unlocks, arming, messages.Add);

        Assert.Equal([Id], shapes.Unlocked);
        Assert.True(unlocks.IsAnyUnlocked);
        Assert.Equal(1, protection.UnprotectCount);
        Assert.False(arming.IsArmed);
        Assert.Empty(messages);
    }

    [Fact]
    public void A_cell_selection_is_refused_with_an_actionable_message()
    {
        var selection = new FakeSelection(
            SelectedShapeOutcome.Refused(SelectedShapeRefusalReason.SelectionIsNotASingleShape));
        var shapes = new FakeShapes();
        var protection = new FakeProtection();
        var unlocks = new ShapeUnlockState();
        List<string> messages = [];

        UnlockShapeCommand.Run(
            selection, shapes, protection, unlocks, new ProtectionArmingState(), messages.Add);

        Assert.Empty(shapes.Unlocked);
        Assert.False(unlocks.IsAnyUnlocked);
        Assert.Equal(0, protection.UnprotectCount);
        Assert.Contains("Select one chart shape", Assert.Single(messages), StringComparison.Ordinal);
    }

    /// <summary>
    /// A shape the add-in does not own must not be unlocked, and the message must say why rather than
    /// blaming the selection.
    /// </summary>
    [Fact]
    public void A_user_shape_is_refused_as_not_ours()
    {
        var selection = new FakeSelection(SelectedShapeOutcome.Refused(SelectedShapeRefusalReason.ShapeNotOwned));
        List<string> messages = [];

        UnlockShapeCommand.Run(
            selection, new FakeShapes(), new FakeProtection(), new ShapeUnlockState(),
            new ProtectionArmingState(), messages.Add);

        Assert.Contains("not drawn by Gantt Creator", Assert.Single(messages), StringComparison.Ordinal);
    }

    /// <summary>
    /// The unlock is recorded BEFORE the release, so a host that refuses the release still leaves the
    /// scheduler suspended rather than rebuilding the chart over a half-finished edit.
    /// </summary>
    [Fact]
    public void A_refused_release_still_leaves_the_refresh_suspended()
    {
        var selection = new FakeSelection(SelectedShapeOutcome.Ok(Id));
        var protection = new FakeProtection { UnprotectRefusal = ProtectionWriteRefusal.HostRejected };
        var unlocks = new ShapeUnlockState();
        List<string> messages = [];

        UnlockShapeCommand.Run(
            selection, new FakeShapes(), protection, unlocks, new ProtectionArmingState(), messages.Add);

        Assert.True(unlocks.IsAnyUnlocked);
        Assert.Contains("could not be released", Assert.Single(messages), StringComparison.Ordinal);
    }

    [Fact]
    public void Locking_relocks_clears_the_unlocks_and_re_applies_the_protection()
    {
        var shapes = new FakeShapes();
        var protection = new FakeProtection();
        var unlocks = new ShapeUnlockState();
        _ = unlocks.Add(Id);
        var arming = new ProtectionArmingState();
        List<string> messages = [];

        LockShapesCommand.Run(shapes, protection, unlocks, arming, messages.Add);

        Assert.Equal(1, shapes.LockCount);
        Assert.False(unlocks.IsAnyUnlocked);
        Assert.True(arming.IsArmed);
        Assert.Equal(1, protection.ProtectCount);
        Assert.Empty(messages);
    }

    /// <summary>
    /// The unlocks are cleared only after the shapes are locked: clearing first would re-enable the
    /// automatic refresh while a shape is still writable.
    /// </summary>
    [Fact]
    public void A_refused_lock_leaves_the_unlocks_and_the_suspension_in_place()
    {
        var shapes = new FakeShapes { LockRefusal = ShapeWriteRefusal.TargetProtected };
        var protection = new FakeProtection();
        var unlocks = new ShapeUnlockState();
        _ = unlocks.Add(Id);
        List<string> messages = [];

        LockShapesCommand.Run(shapes, protection, unlocks, new ProtectionArmingState(), messages.Add);

        Assert.True(unlocks.IsAnyUnlocked);
        Assert.Equal(0, protection.ProtectCount);
        Assert.Contains("could not be locked", Assert.Single(messages), StringComparison.Ordinal);
    }

    private sealed class FakeSelection(SelectedShapeOutcome outcome) : ISelectedShapePort
    {
        public SelectedShapeOutcome Read() => outcome;
    }

    private sealed class FakeShapes : IShapeWritePort
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

        public IReadOnlyList<string> ListOwned() => [];

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
        public ProtectionWriteRefusal? UnprotectRefusal { get; init; }

        public int ProtectCount { get; private set; }

        public int UnprotectCount { get; private set; }

        public WorksheetProtectionState? Read() => new(ContentsProtected: false, ShapesProtected: false);

        public ProtectionWriteOutcome Protect()
        {
            ProtectCount++;
            return ProtectionWriteOutcome.Ok();
        }

        public ProtectionWriteOutcome Unprotect()
        {
            UnprotectCount++;
            return UnprotectRefusal is { } refusal
                ? ProtectionWriteOutcome.Refused(refusal)
                : ProtectionWriteOutcome.Ok();
        }
    }
}
