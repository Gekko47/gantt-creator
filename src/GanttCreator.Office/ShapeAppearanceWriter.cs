using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// The narrow host surface <see cref="ShapeAppearanceWriter"/> drives.
/// </summary>
/// <remarks>
/// <para>
/// Implemented explicitly by <see cref="ExcelShapeWriter"/>, and only by it. The
/// collaborator takes this interface rather than the concrete writer so the
/// dependency runs one way: writer → collaborator → host surface. A collaborator
/// that held the writer would make the two types mutually dependent.
/// </para>
/// <para>
/// Every member is one of the writer's existing <c>internal virtual</c> seams or its
/// existing best-effort discard. Routing through them is what keeps the refactor
/// behaviour-preserving: a test subclass that overrides <c>ApplyStyle</c> still
/// observes the call, because the collaborator reaches the host through this
/// interface and the explicit implementations forward to the virtual members.
/// </para>
/// </remarks>
internal interface IShapeAppearanceHost
{
    /// <summary>Writes a resolved style onto one shape, verifying before it writes.</summary>
    /// <param name="shape">The shape to style.</param>
    /// <param name="request">The request whose style is applied.</param>
    void ApplyStyle(Excel.Shape shape, OfficeShapeRequest request);

    /// <summary>Writes a text box's content, typography, and alignment.</summary>
    /// <param name="shape">The text box.</param>
    /// <param name="request">The request whose text members are written.</param>
    void ApplyText(Excel.Shape shape, OfficeShapeRequest request);

    /// <summary>Writes only a text box's string, for a formatting-preserving duplicate.</summary>
    /// <param name="shape">The text box.</param>
    /// <param name="request">The request carrying the string.</param>
    void ApplyTextContent(Excel.Shape shape, OfficeShapeRequest request);

    /// <summary>Whether the live typography already matches the request.</summary>
    /// <param name="shape">The text box to compare.</param>
    /// <param name="request">The request whose typography would be written.</param>
    /// <returns><see langword="true"/> when no typography write is needed.</returns>
    bool TextFormattingMatches(Excel.Shape shape, OfficeShapeRequest request);

    /// <summary>Whether the live string already equals the request's.</summary>
    /// <param name="shape">The text box to compare.</param>
    /// <param name="expected">The scene's string, or null for an empty one.</param>
    /// <returns><see langword="true"/> when the string needs no write.</returns>
    bool TextMatches(Excel.Shape shape, string? expected);

    /// <summary>Deletes a shape that could not be styled or stamped, best effort.</summary>
    /// <param name="shape">The shape to delete.</param>
    void DiscardUnownedShape(Excel.Shape shape);
}

/// <summary>
/// Applies and verifies shape appearance — style and text — on behalf of
/// <see cref="ExcelShapeWriter"/>, containing host failures as typed results.
/// </summary>
/// <param name="host">The writer surface every write and read goes through.</param>
/// <remarks>
/// <para>
/// This collaborator owns the <em>decision and containment</em> half of appearance:
/// whether a write is needed, what a host failure means on each path, and which
/// exceptions are contract faults rather than host faults. The host keeps the
/// <em>write</em> half through its virtual seams, so no COM member is touched here.
/// </para>
/// <para>
/// The three create-path methods discard the shape on failure; the three update-path
/// methods do not, because there is nothing to discard — the shape was found, not
/// created. That asymmetry is deliberate and documented on the create port.
/// </para>
/// <para>
/// <see cref="ArgumentOutOfRangeException"/> rethrown by
/// <see cref="TryApplyContent"/> and <see cref="TryApplyStyle"/> is a contract fault
/// (an unmapped hatch pattern or alignment), never a host failure, so it is not
/// convertible to a refusal.
/// </para>
/// </remarks>
internal sealed class ShapeAppearanceWriter(IShapeAppearanceHost host)
{
    /// <summary>
    /// Applies the style, discarding the shape if the host refuses it.
    /// </summary>
    /// <param name="shape">The created shape.</param>
    /// <param name="request">The request whose style is applied.</param>
    /// <returns>
    /// <see langword="true"/> when the style landed; <see langword="false"/>
    /// when the host refused it and the shape has been discarded.
    /// </returns>
    /// <remarks>
    /// Best-effort on the same reasoning as the ownership stamp's discard: a shape
    /// that could not be styled must not be left on the sheet, because the caller is
    /// about to be told the create failed and would never learn about the orphan.
    /// <para>
    /// The outcome is returned rather than swallowed, because the caller must not
    /// continue past it. Continuing would stamp ownership onto a shape this method has
    /// just deleted, and that write is either an escaping error or a false success for
    /// a shape that no longer exists. The caller turns <see langword="false"/> into
    /// <see cref="ShapeWriteRefusal.HostRejected"/> and skips the ownership stamp
    /// entirely.
    /// </para>
    /// </remarks>
    internal bool ApplyStyleOrDiscard(Excel.Shape shape, OfficeShapeRequest request)
    {
        // CA1031: see the remarks. A throwing style write must not escape create as an
        // unhandled COM exception from inside a render command, and must not replace
        // the typed refusal the caller already has.
#pragma warning disable CA1031
        try
        {
            host.ApplyStyle(shape, request);
        }
        catch (Exception)
        {
            host.DiscardUnownedShape(shape);
            return false;
        }
#pragma warning restore CA1031

        return true;
    }

    /// <summary>
    /// Writes a new text box's content, discarding the shape if the host refuses.
    /// </summary>
    /// <param name="shape">The created text box.</param>
    /// <param name="request">The request whose text members are written.</param>
    /// <returns>
    /// <see langword="true"/> when the content landed; <see langword="false"/>
    /// when the host refused it and the shape has been discarded.
    /// </returns>
    /// <remarks>
    /// <para>
    /// The create-path twin of <see cref="ApplyStyleOrDiscard"/>, and it exists
    /// because a text box is the one kind whose content is written inside the host's
    /// shape creation, before the shape is stamped and before the caller has any
    /// handle on it. A throwing <c>ApplyText</c> there had two bad outcomes: the
    /// exception escaped the create as an unhandled COM error from inside a render
    /// command, and the text box stayed on the sheet with no ownership tag, so
    /// nothing could ever update or remove it.
    /// </para>
    /// <para>
    /// Returning <see langword="false"/> makes the host's creation step return
    /// <see langword="null"/>, which the create path already reports as
    /// <see cref="ShapeWriteRefusal.HostRejected"/> — the same typed refusal every
    /// other create failure uses, so the caller needs no new case.
    /// </para>
    /// </remarks>
    internal bool ApplyTextOrDiscard(Excel.Shape shape, OfficeShapeRequest request)
    {
        // CA1031: see the remarks. A throwing content write must not escape the create
        // as an unhandled COM exception, and must not leave an untagged text box behind
        // for a caller that was told the create failed.
#pragma warning disable CA1031
        try
        {
            host.ApplyText(shape, request);
        }
        catch (Exception)
        {
            host.DiscardUnownedShape(shape);
            return false;
        }
#pragma warning restore CA1031

        return true;
    }

    /// <summary>
    /// Writes only a duplicated text box's string, discarding it if the host refuses.
    /// </summary>
    /// <param name="shape">The duplicate, whose formatting came from the template.</param>
    /// <param name="request">The request carrying the string to write.</param>
    /// <returns>
    /// <see langword="true"/> when the string landed; <see langword="false"/> when the
    /// host refused it and the duplicate was discarded.
    /// </returns>
    /// <remarks>
    /// A template duplicate already carries every formatting member the full text write
    /// applies: the template key includes the font family, size, weight, alignment,
    /// text colour, and the left inset that expresses a grouped child's indentation.
    /// Re-writing them cost 25.7 ms per text box on the measured 20-activity copy and
    /// was 71% of the whole create block; the string is the only member that differs.
    /// The first shape of each key still takes the full path, so its formatting is never
    /// inherited from an empty host default.
    /// </remarks>
    internal bool ApplyTextContentOrDiscard(Excel.Shape shape, OfficeShapeRequest request)
    {
        // CA1031: same rule as ApplyTextOrDiscard - a host failure is a typed refusal,
        // never an escaping COM exception, and the duplicate must not survive untagged.
#pragma warning disable CA1031
        try
        {
            host.ApplyTextContent(shape, request);
        }
        catch (Exception)
        {
            host.DiscardUnownedShape(shape);
            return false;
        }
#pragma warning restore CA1031

        return true;
    }

    /// <summary>
    /// Applies a text box's content on the update path, containing a host failure.
    /// </summary>
    /// <param name="shape">The owned shape whose content is rewritten.</param>
    /// <param name="request">The request whose text members are written.</param>
    /// <returns>
    /// <see langword="true"/> when the content landed; <see langword="false"/>
    /// when the host refused the write.
    /// </returns>
    /// <remarks>
    /// The update-path twin of <see cref="ApplyTextOrDiscard"/>, and it differs in
    /// exactly one respect: there is nothing to discard. The shape was found, not
    /// created, so a failed content write leaves an existing owned shape in place and
    /// the caller is told the update failed.
    /// <para>
    /// <see cref="ArgumentOutOfRangeException"/> is deliberately not caught. It is what
    /// the host's alignment mapping raises for a scene alignment with no host mapping -
    /// a contract fault the caller must see, and the same rule
    /// <see cref="ApplyTextOrDiscard"/> follows.
    /// </para>
    /// </remarks>
    internal bool TryApplyContent(Excel.Shape shape, OfficeShapeRequest request)
    {
        // CA1031: see the remarks. A host refusal must not escape the update as an
        // unhandled COM exception, and must not replace the typed refusal the
        // reconciler already knows how to report.
#pragma warning disable CA1031
        try
        {
            // Read-back verification: a shape whose live typography and string already
            // match the request pays reads and no writes (ADR-0054 amendment 2026-11-20).
            if (host.TextFormattingMatches(shape, request) && host.TextMatches(shape, request.Text))
            {
                return true;
            }

            host.ApplyText(shape, request);
        }
        catch (ArgumentOutOfRangeException)
        {
            throw;
        }
        catch (Exception)
        {
            return false;
        }
#pragma warning restore CA1031

        return true;
    }

    /// <summary>
    /// Writes only the string of a text box whose formatting the grouped pass already applied.
    /// </summary>
    /// <param name="shape">The owned text box.</param>
    /// <param name="request">The request carrying the string.</param>
    /// <returns><see langword="true"/> when the string landed; otherwise a host refusal.</returns>
    internal bool TryApplyTextContent(Excel.Shape shape, OfficeShapeRequest request)
    {
        // CA1031: the same rule as TryApplyContent - a host refusal is a typed result,
        // never an escaping COM exception from inside a render command.
#pragma warning disable CA1031
        try
        {
            if (host.TextMatches(shape, request.Text))
            {
                return true;
            }

            host.ApplyTextContent(shape, request);
        }
        catch (Exception)
        {
            return false;
        }
#pragma warning restore CA1031

        return true;
    }

    /// <summary>
    /// Applies the style on the update path, containing a host failure.
    /// </summary>
    /// <param name="shape">The owned shape whose style is rewritten.</param>
    /// <param name="request">The request whose style members are written.</param>
    /// <returns>
    /// <see langword="true"/> when the style landed; <see langword="false"/> when
    /// the host refused the write.
    /// </returns>
    /// <remarks>
    /// The update-path twin of <see cref="ApplyStyleOrDiscard"/>, and it differs in
    /// exactly one respect: there is nothing to discard, for the same reason
    /// <see cref="TryApplyContent"/> does not discard. The geometry write has already
    /// run by this point and is not undone, so a refused style leaves the shape moved
    /// and un-restyled; that is stated on
    /// <see cref="IShapeWritePort.Update(OfficeShapeRequest)"/> rather than left for a
    /// reader to infer.
    /// <para>
    /// <see cref="ArgumentOutOfRangeException"/> is deliberately not caught, for the
    /// reason the host's style write documents: it is its hatch mapping's contract
    /// fault, not a host failure.
    /// </para>
    /// </remarks>
    internal bool TryApplyStyle(Excel.Shape shape, OfficeShapeRequest request)
    {
        // CA1031: see the remarks. A host refusal must not escape the update as an
        // unhandled COM exception, and must not replace the typed refusal the
        // reconciler already knows how to report.
#pragma warning disable CA1031
        try
        {
            host.ApplyStyle(shape, request);
        }
        catch (ArgumentOutOfRangeException)
        {
            throw;
        }
        catch (Exception)
        {
            return false;
        }
#pragma warning restore CA1031

        return true;
    }
}
