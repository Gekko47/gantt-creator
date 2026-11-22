using System.Globalization;

namespace GanttCreator.Office;

/// <summary>
/// R7.5 PD-10: deletes the pasted group on a failure path after a paste materialised, so a
/// failed transfer leaves no residue on the user's slide.
/// </summary>
/// <remarks>
/// <para>
/// <b>R7.5's recorded obligation.</b> The landed R7.3 paste-stage remarks state "R7.5's outcome
/// owns cleanup on every path, and ADR-0066 D3's group-delete mechanic (delete the group, never
/// the range) is recorded for it"; the architecture's "Error handling" section requires
/// "partial temporary export shapes are removed". This helper is that mechanism; R7.6's transfer
/// seam invokes it when a paste succeeded but a later stage failed.
/// </para>
/// <para>
/// <b>Attempted once, never retried.</b> A cleanup delete must not mutate further on failure;
/// a refused delete is written to the technical record and named, never silently accepted.
/// </para>
/// </remarks>
public static class PowerPointPasteResidueCleanup
{
    /// <summary>
    /// Deletes the pasted group when one materialised, and reports a refused delete.
    /// </summary>
    /// <param name="groupId">The pasted group's identifier, or <see langword="null"/> when no paste materialised.</param>
    /// <param name="port">The paste port that performs the delete.</param>
    /// <param name="technicalRecord">Receives the residue-refusal record, or <see langword="null"/>.</param>
    /// <returns>
    /// The delete result, or <see langword="null"/> when there was nothing to clean because no
    /// paste materialised.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="port"/> is <see langword="null"/>.</exception>
    public static PowerPointGroupDeleteResult? Cleanup(
        int? groupId,
        IPowerPointPastePort port,
        Action<string>? technicalRecord)
    {
        ArgumentNullException.ThrowIfNull(port);

        if (groupId is not int id)
        {
            return null;
        }

        PowerPointGroupDeleteResult result = port.DeletePastedGroup(id);
        if (result.Outcome != PowerPointGroupDeleteOutcome.Deleted)
        {
            string head = string.Create(
                CultureInfo.InvariantCulture,
                $"r7.5 transfer residue cleanup refused outcome={result.Outcome}");
            string code = result.HResult is int hresult
                ? string.Create(CultureInfo.InvariantCulture, $" hresult=0x{hresult:X8}")
                : string.Empty;
            technicalRecord?.Invoke(head + code);
        }

        return result;
    }
}
