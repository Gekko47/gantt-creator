namespace GanttCreator.Office;

/// <summary>
/// Selects which ownership scheme a shape-writing adapter stamps and proves: the live
/// chart's <see cref="ShapeOwnershipTag"/> scheme, or the editable-export staging scheme
/// (R6.2, work-item D4).
/// </summary>
/// <remarks>
/// <para>
/// The two schemes differ in exactly two places - the name written on create and the
/// tag written and proved - so they are stated once here rather than as branches inside
/// each adapter operation. <see cref="ExcelShapeWriter"/> takes one scheme for its whole
/// life; a staging adapter is the same writer constructed with
/// <see cref="Staging"/>, which is why the staging composition reuses the live create,
/// update, delete, and list machinery instead of duplicating COM code.
/// </para>
/// <para>
/// <b>The invariant both schemes preserve is <c>tag = f(name)</c>.</b> Ownership is
/// proved by recomputing the expected tag from the live shape name, never by trusting
/// the stored alternative text. Because the two tag prefixes differ, each scheme's
/// recomputation can only ever match its own family's shapes - a live scan is blind to
/// staged shapes and a staging scan is blind to live owned shapes, structurally rather
/// than by convention.
/// </para>
/// </remarks>
public sealed class ShapeOwnershipScheme
{
    private readonly string _namePrefix;
    private readonly Func<string, string> _tagForName;

    private ShapeOwnershipScheme(bool isStaging, string namePrefix, Func<string, string> tagForName)
    {
        ArgumentNullException.ThrowIfNull(namePrefix);
        ArgumentNullException.ThrowIfNull(tagForName);
        IsStaging = isStaging;
        _namePrefix = namePrefix;
        _tagForName = tagForName;
    }

    /// <summary>
    /// Gets the live chart's scheme: the shape name is the primitive identifier and
    /// the tag is <see cref="ShapeOwnershipTag.ForPrimitiveId"/> of it.
    /// </summary>
    public static ShapeOwnershipScheme Live { get; } = new(
        isStaging: false,
        string.Empty,
        ShapeOwnershipTag.ForPrimitiveId);

    /// <summary>
    /// Gets the editable-export staging scheme: the shape name is
    /// <see cref="ShapeStagingTag.NamePrefix"/> followed by the primitive identifier,
    /// and the tag is <see cref="ShapeStagingTag.ForStagedName"/> of that name.
    /// </summary>
    public static ShapeOwnershipScheme Staging { get; } = new(
        isStaging: true,
        ShapeStagingTag.NamePrefix,
        ShapeStagingTag.ForStagedName);

    /// <summary>Gets whether this is the staging scheme.</summary>
    public bool IsStaging { get; }

    /// <summary>Builds the shape name this scheme writes for one primitive identifier.</summary>
    /// <param name="primitiveId">The stable scene primitive identifier.</param>
    /// <returns>The exact shape name to write and look up.</returns>
    /// <exception cref="ArgumentException">
    /// <paramref name="primitiveId"/> is <see langword="null"/>, empty, or whitespace.
    /// </exception>
    public string NameFor(string primitiveId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);
        return _namePrefix + primitiveId;
    }

    /// <summary>Builds the exact ownership tag this scheme expects for one shape name.</summary>
    /// <param name="shapeName">The full shape name, already composed by <see cref="NameFor"/>.</param>
    /// <returns>The tag the scheme's ownership proof compares against.</returns>
    /// <exception cref="ArgumentException">
    /// <paramref name="shapeName"/> is <see langword="null"/>, empty, or whitespace.
    /// </exception>
    public string TagForName(string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
        return _tagForName(shapeName);
    }
}
