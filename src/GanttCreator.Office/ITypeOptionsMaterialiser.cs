namespace GanttCreator.Office;

/// <summary>Why TypeOptions materialisation refused to mutate the workbook.</summary>
public enum TypeOptionsRefusalReason
{
    /// <summary>No active workbook exists.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The configuration worksheet is missing.</summary>
    ConfigSheetMissing = 1,

    /// <summary>The visible data table or the Type catalogue table is missing.</summary>
    TableMissing = 2,

    /// <summary>The configuration catalogue or stored hash is invalid.</summary>
    CatalogueHashMismatch = 3,

    /// <summary>The target worksheet or workbook structure is protected.</summary>
    TargetProtected = 4,
}

/// <summary>The typed result of materialising the Type dropdown.</summary>
/// <param name="Succeeded">Whether the name and validation were applied.</param>
/// <param name="Refusal">The refusal reason when unsuccessful.</param>
public sealed record TypeOptionsMaterialiseOutcome(
    bool Succeeded,
    TypeOptionsRefusalReason? Refusal)
{
    /// <summary>Creates a successful outcome.</summary>
    public static TypeOptionsMaterialiseOutcome Ok() => new(true, null);

    /// <summary>Creates a refusal outcome.</summary>
    public static TypeOptionsMaterialiseOutcome Refused(TypeOptionsRefusalReason reason) => new(false, reason);
}

/// <summary>Materialises the TypeOptions name and Type-column validation.</summary>
public interface ITypeOptionsMaterialiser
{
    /// <summary>Applies the named-range dropdown to the current Type body.</summary>
    TypeOptionsMaterialiseOutcome Materialise();
}
