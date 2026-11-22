namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R7.5 PD-10: the paste-residue cleanup. A paste that materialised is deleted once on a
/// failure path; a missing group or refused delete is recorded and named, never retried; no
/// paste means nothing to clean.
/// </summary>
public sealed class PowerPointPasteResidueCleanupTests
{
    private const int _groupId = 42;

    [Fact]
    public void A_present_group_id_deletes_the_pasted_group()
    {
        FakeDeletePort port = new();

        PowerPointGroupDeleteResult result = Assert.IsType<PowerPointGroupDeleteResult>(
            PowerPointPasteResidueCleanup.Cleanup(_groupId, port, technicalRecord: null));

        Assert.Equal(PowerPointGroupDeleteOutcome.Deleted, result.Outcome);
        Assert.Equal([_groupId], port.DeletedIds);
    }

    [Fact]
    public void A_null_group_id_cleans_nothing()
    {
        FakeDeletePort port = new();

        PowerPointGroupDeleteResult? result = PowerPointPasteResidueCleanup.Cleanup(
            groupId: null,
            port,
            technicalRecord: null);

        Assert.Null(result);
        Assert.Empty(port.DeletedIds);
    }

    [Fact]
    public void A_refused_delete_is_recorded_and_named()
    {
        FakeDeletePort port = new()
        {
            Outcome = PowerPointGroupDeleteOutcome.HostRejected,
            HResult = unchecked((int)0x80004005),
        };
        List<string> records = [];

        _ = PowerPointPasteResidueCleanup.Cleanup(_groupId, port, records.Add);

        string record = Assert.Single(records);
        Assert.Contains("residue cleanup refused", record, StringComparison.Ordinal);
        Assert.Contains("outcome=HostRejected", record, StringComparison.Ordinal);
        Assert.Contains("hresult=0x80004005", record, StringComparison.Ordinal);
    }

    [Fact]
    public void A_missing_group_delete_is_recorded()
    {
        FakeDeletePort port = new() { Outcome = PowerPointGroupDeleteOutcome.GroupMissing };
        List<string> records = [];

        _ = PowerPointPasteResidueCleanup.Cleanup(_groupId, port, records.Add);

        Assert.Contains("outcome=GroupMissing", Assert.Single(records), StringComparison.Ordinal);
    }

    [Fact]
    public void A_successful_delete_writes_no_refusal_record()
    {
        FakeDeletePort port = new();
        List<string> records = [];

        _ = PowerPointPasteResidueCleanup.Cleanup(_groupId, port, records.Add);

        Assert.Empty(records);
    }

    [Fact]
    public void A_cleanup_delete_is_attempted_once()
    {
        FakeDeletePort port = new();

        _ = PowerPointPasteResidueCleanup.Cleanup(_groupId, port, technicalRecord: null);

        // The cleanup is not a retry loop: exactly one delete attempt, whatever the outcome.
        Assert.Single(port.DeletedIds);
    }

    [Fact]
    public void A_null_port_is_refused() =>
        _ = Assert.Throws<ArgumentNullException>(
            () => PowerPointPasteResidueCleanup.Cleanup(_groupId, null!, technicalRecord: null));

    private sealed class FakeDeletePort : IPowerPointPastePort
    {
        internal List<int> DeletedIds { get; } = [];

        internal PowerPointGroupDeleteOutcome Outcome { get; init; } = PowerPointGroupDeleteOutcome.Deleted;

        internal int? HResult { get; init; }

        public PowerPointPasteObservation PasteSpecialShape() =>
            throw new InvalidOperationException("The cleanup never pastes.");

        public PowerPointGroupDeleteResult DeletePastedGroup(int groupId)
        {
            DeletedIds.Add(groupId);
            return new PowerPointGroupDeleteResult(Outcome, HResult);
        }
    }
}
