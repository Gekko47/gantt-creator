using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

public partial class ExcelShapeWriter
{
    /// <inheritdoc />
    public ShapeWriteOutcome Delete(string primitiveId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);

        // ADR-0008 D4's first statement, before any workbook access. ADR-0054: inside a
        // reconcile batch the target was already checked when it was resolved and events are
        // off for the whole pass, so the repeated query can only return the same answer.
        if (!_batch.Active && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        Excel.Shape? existing = FindShapeByName(shapes, NameFor(primitiveId));
        if (existing is null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        // The ownership filter, not the name, authorises the delete. A user
        // shape that happens to share a name is never deleted (R4.8 D1).
        if (!CarriesOwnershipTagFor(existing, primitiveId))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        existing.Delete();

        // A deleted name must not be served from the batch memo afterwards: a later create
        // in the same batch legitimately reuses it. The remembered ownership answer goes
        // with it, or the new shape would be judged by the old shape's tag.
        if (_batch.Proxies is not null)
        {
            string deletedName = NameFor(primitiveId);
            _ = _batch.Proxies.Remove(deletedName);
            _ = _batch.Tags.Remove(deletedName);
        }

        return ShapeWriteOutcome.Ok();
    }

    /// <inheritdoc />
    public IReadOnlyList<string> ListOwned()
    {
        Excel.Shapes? shapes = ResolveShapes(out _);
        if (shapes is null)
        {
            return [];
        }

        List<string> owned = [];
        int count = GetShapeCount(shapes);
        int inventoryLogged = 0;
        for (int index = 1; index <= count; index++)
        {
            Excel.Shape? shape = GetShapeAt(shapes, index);
            if (shape is null)
            {
                continue;
            }

            string name = shape.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            // Read the tag only for a name this writer could have created (O4). The
            // alternative-text read is a host round trip; a name outside our own two
            // shapes cannot be one of ours, so the read buys nothing for it. The proof
            // is in CouldBeOwnedName, and the tag comparison stays exactly as it was
            // for every name that passes this test -- including the recompute that
            // makes a user's renamed copy of one of our shapes fail ownership.
            bool ownedTag = CouldBeOwnedName(name) && CarriesTagForName(shape, name);
            if (ownedTag)
            {
                owned.Add(name);
            }

            // This walk already visits every shape and reads every name, which is exactly
            // what EnsureBatchNameIndex would spend a second pass doing. Feeding it here
            // means the reconcile's first name lookup is a dictionary hit rather than
            // another full scan with a Name read per shape.
            if (_batch.Proxies is not null)
            {
                _ = _batch.Proxies.TryAdd(name, shape);
            }

            // Phase-0 inventory: a chart-shaped shape that is NOT owned, or any
            // splitter band, is logged once per shape per refresh. Names carry the
            // chart's own "<rowId>:<role>" form, so a shape the reconciler cannot see
            // (an untagged band from an older build) is identifiable here rather than
            // inferred from the pixels. Bounded to 40 lines so a heavily annotated
            // sheet cannot flood the log.
            if (inventoryLogged < 40
                && (name.Contains("plitter", StringComparison.OrdinalIgnoreCase)
                    || (!ownedTag && name.Contains(':', StringComparison.Ordinal))))
            {
                TraceInventoryShape(shape, name, ownedTag);
                inventoryLogged++;
            }
        }

        if (_batch.Proxies is not null)
        {
            // The loop above covered index 1..Count, so every name the host holds is now
            // in the memo and an absent name can be answered without a host call.
            _batch.NameIndexed = true;
        }

        // Ordinal sort so the returned order is a property of the shape names
        // rather than of the host's z-order, which R4.5 and R4.7 both mutate.
        owned.Sort(StringComparer.Ordinal);
        return owned;
    }

    /// <inheritdoc />
    public ShapeWriteOutcome UnlockShape(string primitiveId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);

        // Same guard-first rule as every other entry point, and it is load-bearing here too: an
        // unlock clears `Locked`, which the host refuses on a sheet the add-in does not own.
        if (!_batch.Active && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        Excel.Shape? existing = FindShapeByName(shapes, NameFor(primitiveId));

        // The tag authorises the unlock, not the name: handing the user a writable shape the add-in
        // does not own would turn "unlock" into a way to edit anything on the sheet (ADR-0055 D6).
        if (existing is null || !CarriesOwnershipTagFor(existing, primitiveId))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        if (ReadLocked(existing))
        {
            ApplyUnlock(existing);
        }

        return ShapeWriteOutcome.Ok();
    }

    /// <inheritdoc />
    public ShapeWriteOutcome LockOwnedShapes()
    {
        if (!_batch.Active && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        int count = GetShapeCount(shapes);
        for (int index = 1; index <= count; index++)
        {
            Excel.Shape? shape = GetShapeAt(shapes, index);
            if (shape is null)
            {
                continue;
            }

            string name = shape.Name;
            if (string.IsNullOrWhiteSpace(name) || !CouldBeOwnedName(name))
            {
                continue;
            }

            // The tag, not the name, is the ownership proof -- the same rule Delete and Update
            // follow, so a user's shape that shares a name is left unlocked.
            if (!CarriesTagForName(shape, name))
            {
                continue;
            }

            if (!ReadLocked(shape))
            {
                ApplyLock(shape);
            }
        }

        return ShapeWriteOutcome.Ok();
    }

    /// <summary>Reads a shape's lock state. Test seam over the COM property.</summary>
    /// <param name="shape">The live shape.</param>
    /// <returns>The host's lock state.</returns>
    internal virtual bool ReadLocked(Excel.Shape shape) => shape.Locked;

    /// <summary>Locks a shape. Test seam over the COM property.</summary>
    /// <param name="shape">The live shape.</param>
    internal virtual void ApplyLock(Excel.Shape shape) => shape.Locked = true;

    /// <summary>Clears a shape's lock. Test seam over the COM property.</summary>
    /// <param name="shape">The live shape.</param>
    /// <remarks>The PIA types this property as <c>Boolean</c>, not <c>MsoTriState</c>.</remarks>
    internal virtual void ApplyUnlock(Excel.Shape shape) => shape.Locked = false;

}
