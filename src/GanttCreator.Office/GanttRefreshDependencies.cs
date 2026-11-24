using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// The required collaborators of one refresh pipeline.
/// </summary>
/// <remarks>
/// A parameter object rather than a twenty-two-argument constructor: the nine
/// required ports are validated once here, and the pipeline's optional values
/// live in <see cref="GanttRefreshOptions"/>.
/// </remarks>
public sealed record GanttRefreshDependencies
{
    /// <summary>
    /// Initialises the dependency set, refusing a null required port at
    /// construction rather than at first use.
    /// </summary>
    public GanttRefreshDependencies(
        IGanttTableReader tableReader,
        IConfigCatalogueReader configReader,
        IWorksheetProtectionGuard protectionGuard,
        IPanelGridMeasurementPort panelMeasurement,
        IDurationWritePort durationWriter,
        IRowHeightNormalisationPort rowHeightNormaliser,
        IOutlineGroupPort outlineWriter,
        ISceneBuildRequestFactory requestFactory,
        IShapeWritePort shapeWriter)
    {
        TableReader = tableReader ?? throw new ArgumentNullException(nameof(tableReader));
        ConfigReader = configReader ?? throw new ArgumentNullException(nameof(configReader));
        ProtectionGuard = protectionGuard ?? throw new ArgumentNullException(nameof(protectionGuard));
        PanelMeasurement = panelMeasurement ?? throw new ArgumentNullException(nameof(panelMeasurement));
        DurationWriter = durationWriter ?? throw new ArgumentNullException(nameof(durationWriter));
        RowHeightNormaliser = rowHeightNormaliser ?? throw new ArgumentNullException(nameof(rowHeightNormaliser));
        OutlineWriter = outlineWriter ?? throw new ArgumentNullException(nameof(outlineWriter));
        RequestFactory = requestFactory ?? throw new ArgumentNullException(nameof(requestFactory));
        ShapeWriter = shapeWriter ?? throw new ArgumentNullException(nameof(shapeWriter));
    }

    /// <summary>Reads and validates the visible table.</summary>
    public IGanttTableReader TableReader { get; }

    /// <summary>Reads the configuration catalogues.</summary>
    public IConfigCatalogueReader ConfigReader { get; }

    /// <summary>The protection guard consulted before any write.</summary>
    public IWorksheetProtectionGuard ProtectionGuard { get; }

    /// <summary>Measures the data panel.</summary>
    public IPanelGridMeasurementPort PanelMeasurement { get; }

    /// <summary>Writes the derived <c>Duration</c> column.</summary>
    public IDurationWritePort DurationWriter { get; }

    /// <summary>Normalises body-row heights.</summary>
    public IRowHeightNormalisationPort RowHeightNormaliser { get; }

    /// <summary>Rebuilds the native outline.</summary>
    public IOutlineGroupPort OutlineWriter { get; }

    /// <summary>Builds the scene request from the read table.</summary>
    public ISceneBuildRequestFactory RequestFactory { get; }

    /// <summary>Reconciles the scene's shapes.</summary>
    public IShapeWritePort ShapeWriter { get; }
}

/// <summary>
/// The optional collaborators and flags of one refresh pipeline.
/// </summary>
/// <param name="ColumnPresentation">Repairs the managed columns' presentation, or null to skip.</param>
/// <param name="IdentityRepairer">Seeds or repairs row identifiers, or null to skip.</param>
/// <param name="StateScope">The application-state scope, or null for the no-op scope.</param>
/// <param name="StyleKeyWriter">Writes the engine-owned <c>StyleKey</c> column, or null to skip.</param>
/// <param name="RowWarningHighlighter">Paints validation fills, or null to skip.</param>
/// <param name="TypeChangeNormaliser">Heals type-dependent engine-hidden columns, or null to skip.</param>
/// <param name="ChangeTracker">Tracks pending automatic changes, or null.</param>
/// <param name="ShapeDeltaCache">The payload-signature cache, or null.</param>
/// <param name="UseDeltaCache">Whether this pass may skip unchanged payload writes.</param>
/// <param name="Trace">The diagnostic sink, or null for no trace.</param>
/// <param name="ProtectionOwnership">The protection-ownership source, or null.</param>
/// <param name="ProtectionArming">The protection policy, or null when this caller does not arm.</param>
/// <param name="RepairColumnClassification">Whether this pass re-asserts the column classification.</param>
public sealed record GanttRefreshOptions(
    IColumnPresentationPort? ColumnPresentation = null,
    IGanttRowIdentityRepairer? IdentityRepairer = null,
    IApplicationStateScope? StateScope = null,
    IStyleKeyWritePort? StyleKeyWriter = null,
    IRowWarningHighlighter? RowWarningHighlighter = null,
    ITypeChangeNormaliser? TypeChangeNormaliser = null,
    ChangeTracker? ChangeTracker = null,
    ShapeDeltaCache? ShapeDeltaCache = null,
    bool UseDeltaCache = true,
    Action<string>? Trace = null,
    ProtectionOwnershipSource? ProtectionOwnership = null,
    WorksheetProtectionArming? ProtectionArming = null,
    bool RepairColumnClassification = true);
