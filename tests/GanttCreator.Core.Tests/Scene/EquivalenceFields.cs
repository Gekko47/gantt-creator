using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// One required field of one entity-guide field-contract row, as transcribed at R3.16
/// from <c>docs/07-GANTT-ENTITY-GUIDE.md</c> revision 6 "Per-entity field contract".
/// </summary>
/// <remarks>
/// <para>
/// The field list is the contract under test, so it is transcribed exactly and each
/// entry carries the guide section it comes from. A field that the scene does not
/// carry with a non-default, correctly typed value is a defect in the model, never a
/// renderer-side workaround (R3.14 D1).
/// </para>
/// <para>
/// <b>Why a field names a member.</b> A field-contract row names <em>every</em>
/// primitive the row spans - a header is a rectangle plus a label, the data panel is
/// five families of primitive - and its fields are per primitive, not per row. A field
/// therefore states which member it is read from, so "the header text is centred" and
/// "the body cell fill sits at Background" are both stated without either being
/// checked against the wrong primitive. The panel row is the concrete case: R3.16
/// corrected it to name three z-layers, and a single-primitive row could not have
/// carried that correction at all.
/// </para>
/// <para>
/// <b>Ownership.</b> "Shared entity contract" requires ownership metadata beginning
/// <c>GanttCreator.</c> <em>in Excel shapes</em>. The scene's ownership metadata is
/// <see cref="ScenePrimitive.OwnerId"/> plus the role-derived
/// <see cref="ScenePrimitive.PrimitiveId"/>, which is what a renderer composes that
/// tag from; the literal prefixed tag is a Phase-4 port member owned by R4.1 D2 /
/// R4.3. Asserting a <c>GanttCreator.</c> string in Core would be a model extension,
/// which this row forbids.
/// </para>
/// </remarks>
/// <param name="Name">The stable field name reported when the scene fails the field.</param>
/// <param name="GuideCitation">The entity-guide section this field is transcribed from.</param>
/// <param name="Member">The key of the row member this field is read from.</param>
/// <param name="IsSatisfied">
/// Whether the scene carries this field with a non-default, correctly typed value.
/// The whole scene is passed because one field - an external label's placement - is
/// only decidable against the entity it belongs to.
/// </param>
internal sealed record EquivalenceField(
    string Name,
    string GuideCitation,
    string Member,
    Func<GanttScene, ScenePrimitive, bool> IsSatisfied);

/// <summary>One primitive of a field-contract row, keyed by the role its fields name it by.</summary>
/// <param name="Key">The member key that fields refer to this primitive by.</param>
/// <param name="PrimitiveId">The stable role-derived identifier of that primitive.</param>
internal sealed record EquivalenceMember(string Key, string PrimitiveId);

/// <summary>
/// One entity-guide field-contract row, with every field a renderer must read.
/// </summary>
/// <param name="Name">The row's name, used in assertion messages.</param>
/// <param name="GuideRowName">
/// The row's name exactly as the guide's table spells it. The drift guard compares this
/// set against the parsed guide, so a renamed guide row fails the suite.
/// </param>
/// <param name="GuideCitation">The guide row this transcription covers.</param>
/// <param name="Members">The primitives this row spans, in transcription order.</param>
/// <param name="Fields">The required fields, in transcription order.</param>
/// <param name="Exclusion">
/// For a row the model deliberately does not carry: the token that must appear in no
/// primitive identifier. A row is either asserted against real members or excluded,
/// never silently empty.
/// </param>
internal sealed record EquivalenceRow(
    string Name,
    string GuideRowName,
    string GuideCitation,
    IReadOnlyList<EquivalenceMember> Members,
    IReadOnlyList<EquivalenceField> Fields,
    string? Exclusion = null)
{
    /// <summary>Gets whether this row asserts that the model carries nothing at all.</summary>
    public bool IsExclusion => Exclusion is not null;

    /// <summary>
    /// Collects the names of the fields this row does not find satisfied in the scene.
    /// </summary>
    /// <param name="scene">The scene under test.</param>
    /// <returns>
    /// The unsatisfied field names, with one entry naming any absent primitive rather
    /// than passing silently. An empty list means the scene carries every field.
    /// </returns>
    /// <remarks>
    /// This collects rather than asserts so a test can prove which field a
    /// deliberately field-stripped variant loses, which is the row's positive test.
    /// </remarks>
    public IReadOnlyList<string> UnsatisfiedFields(GanttScene scene)
    {
        ArgumentNullException.ThrowIfNull(scene);

        List<string> unsatisfied = [];
        foreach (EquivalenceField field in Fields)
        {
            EquivalenceMember? member = Members.FirstOrDefault(candidate =>
                string.Equals(candidate.Key, field.Member, StringComparison.Ordinal));
            if (member is null)
            {
                unsatisfied.Add($"{field.Name} (no member '{field.Member}' is declared)");
                continue;
            }

            ScenePrimitive? primitive = scene.Primitives.FirstOrDefault(candidate =>
                string.Equals(candidate.PrimitiveId, member.PrimitiveId, StringComparison.Ordinal));
            if (primitive is null)
            {
                unsatisfied.Add($"the primitive '{member.PrimitiveId}' is absent from the scene");
                continue;
            }

            if (!field.IsSatisfied(scene, primitive))
            {
                unsatisfied.Add(field.Name);
            }
        }

        return unsatisfied;
    }

    /// <summary>
    /// Finds the first primitive identifier in the scene containing the given token.
    /// </summary>
    /// <param name="scene">The scene under test.</param>
    /// <param name="token">The identifier fragment to look for.</param>
    /// <returns>The offending identifier, or <see langword="null"/> when there is none.</returns>
    /// <remarks>
    /// Used by an exclusion row, whose contract is that the model carries no such
    /// primitive, so the test proves the absence is real rather than untested.
    /// </remarks>
    public static string? FirstIdentifierContaining(GanttScene scene, string token)
    {
        ArgumentNullException.ThrowIfNull(scene);
        ArgumentException.ThrowIfNullOrWhiteSpace(token);

        return scene.Primitives
            .Select(primitive => primitive.PrimitiveId)
            .FirstOrDefault(identifier => identifier.Contains(token, StringComparison.OrdinalIgnoreCase));
    }
}
