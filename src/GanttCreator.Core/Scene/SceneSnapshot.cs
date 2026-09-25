using System.Text.Json;
using System.Text.Json.Serialization;

namespace GanttCreator.Core.Scene;

/// <summary>Canonical, deterministic JSON serialization for <see cref="GanttScene"/>.</summary>
public static class SceneSnapshot
{
    private static readonly JsonSerializerOptions _serializerOptions = new()
    {
        PropertyNamingPolicy = null,
        PropertyNameCaseInsensitive = false,
        WriteIndented = false,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
    };

    /// <summary>Serializes a scene to canonical compact JSON.</summary>
    /// <param name="scene">The scene to serialize.</param>
    /// <returns>Canonical invariant JSON.</returns>
    public static string Serialize(GanttScene scene)
    {
        ArgumentNullException.ThrowIfNull(scene);
        return JsonSerializer.Serialize(ToDocument(scene), _serializerOptions);
    }

    /// <summary>Deserializes and validates a scene from strict canonical JSON.</summary>
    /// <param name="json">The scene JSON.</param>
    /// <returns>The validated scene.</returns>
    /// <exception cref="JsonException">Thrown when the JSON is malformed or contains an unknown field.</exception>
    /// <exception cref="InvalidDataException">Thrown when the JSON contains invalid scene data.</exception>
    public static GanttScene Deserialize(string json)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(json);
        SceneDocument document = JsonSerializer.Deserialize<SceneDocument>(json, _serializerOptions)
            ?? throw new InvalidDataException("Scene JSON is empty.");
        try
        {
            return FromDocument(document);
        }
        catch (JsonException)
        {
            throw;
        }
        catch (Exception exception) when (exception is ArgumentException or FormatException or OverflowException)
        {
            throw new InvalidDataException("Scene JSON contains invalid scene data.", exception);
        }
    }

    private static SceneDocument ToDocument(GanttScene scene) => new()
    {
        Version = 1,
        ChartBounds = ToRect(scene.ChartBounds),
        PlotBounds = ToRect(scene.PlotBounds),
        Primitives = [.. scene.Primitives.Select(ToPrimitive)],
        Warnings = [.. scene.Warnings.Select(ToWarning)],
    };

    private static PrimitiveDocument ToPrimitive(ScenePrimitive primitive) => primitive switch
    {
        SceneRect rect => new PrimitiveDocument
        {
            Kind = "rect",
            PrimitiveId = rect.PrimitiveId,
            OwnerId = rect.OwnerId.Value,
            ZLayer = (int)rect.ZLayer,
            EntityType = ToNullableInt(rect.EntityType),
            LaneOrder = rect.LaneOrder,
            StackIndex = rect.StackIndex,
            SortOrder = rect.SortOrder,
            Bounds = ToRect(rect.Bounds),
            Style = ToStyle(rect.Style),
        },
        SceneLine line => new PrimitiveDocument
        {
            Kind = "line",
            PrimitiveId = line.PrimitiveId,
            OwnerId = line.OwnerId.Value,
            ZLayer = (int)line.ZLayer,
            EntityType = ToNullableInt(line.EntityType),
            LaneOrder = line.LaneOrder,
            StackIndex = line.StackIndex,
            SortOrder = line.SortOrder,
            From = ToPoint(line.From),
            To = ToPoint(line.To),
            Style = ToStyle(line.Style),
        },
        ScenePolygon polygon => new PrimitiveDocument
        {
            Kind = "polygon",
            PrimitiveId = polygon.PrimitiveId,
            OwnerId = polygon.OwnerId.Value,
            ZLayer = (int)polygon.ZLayer,
            EntityType = ToNullableInt(polygon.EntityType),
            LaneOrder = polygon.LaneOrder,
            StackIndex = polygon.StackIndex,
            SortOrder = polygon.SortOrder,
            Points = [.. polygon.Points.Select(ToPoint)],
            Style = ToStyle(polygon.Style),
        },
        SceneText text => new PrimitiveDocument
        {
            Kind = "text",
            PrimitiveId = text.PrimitiveId,
            OwnerId = text.OwnerId.Value,
            ZLayer = (int)text.ZLayer,
            EntityType = ToNullableInt(text.EntityType),
            LaneOrder = text.LaneOrder,
            StackIndex = text.StackIndex,
            SortOrder = text.SortOrder,
            Text = text.Text,
            TextBounds = ToRect(text.TextBounds),
            Style = ToStyle(text.Style),
            Alignment = (int)text.Alignment,
        },
        SceneGroup group => new PrimitiveDocument
        {
            Kind = "group",
            PrimitiveId = group.PrimitiveId,
            OwnerId = group.OwnerId.Value,
            ZLayer = (int)group.ZLayer,
            EntityType = ToNullableInt(group.EntityType),
            LaneOrder = group.LaneOrder,
            StackIndex = group.StackIndex,
            SortOrder = group.SortOrder,
            ChildPrimitiveIds = [.. group.ChildPrimitiveIds],
        },
        _ => throw new InvalidDataException($"Unsupported scene primitive '{primitive.GetType().Name}'."),
    };

    private static WarningDocument ToWarning(SceneWarning warning) => new()
    {
        OwnerId = warning.OwnerId.Value,
        Code = warning.Code,
        Message = warning.Message,
    };

    private static GanttScene FromDocument(SceneDocument document)
    {
        if (document.Version != 1)
        {
            throw new InvalidDataException($"Unsupported scene snapshot version {document.Version}.");
        }

        ScenePrimitive[] primitives = [.. (document.Primitives ?? []).Select(FromPrimitive)];
        SceneWarning[] warnings = [.. (document.Warnings ?? []).Select(FromWarning)];
        SceneCreationOutcome outcome = GanttScene.TryCreate(
            FromRect(Required(document.ChartBounds, "chartBounds"), "chartBounds"),
            FromRect(Required(document.PlotBounds, "plotBounds"), "plotBounds"),
            primitives,
            warnings);
        return outcome.Scene ?? throw new InvalidDataException($"Scene creation refused: {outcome.Refusal}.");
    }

    private static ScenePrimitive FromPrimitive(PrimitiveDocument document)
    {
        GanttRowId owner = ParseOwner(document.OwnerId);
        ZLayer layer = ParseLayer(document.ZLayer);
        GanttEntityType? entityType = FromNullableInt(document.EntityType);
        return document.Kind switch
        {
            "rect" => new SceneRect(
                Required(document.PrimitiveId, "primitiveId"),
                owner,
                layer,
                FromRect(Required(document.Bounds, "bounds"), "bounds"),
                FromStyle(Required(document.Style, "style")),
                entityType,
                document.LaneOrder,
                document.StackIndex,
                document.SortOrder),
            "line" => new SceneLine(
                Required(document.PrimitiveId, "primitiveId"),
                owner,
                layer,
                FromPoint(Required(document.From, "from"), "from"),
                FromPoint(Required(document.To, "to"), "to"),
                FromStyle(Required(document.Style, "style")),
                entityType,
                document.LaneOrder,
                document.StackIndex,
                document.SortOrder),
            "polygon" => new ScenePolygon(
                Required(document.PrimitiveId, "primitiveId"),
                owner,
                layer,
                [.. (document.Points ?? []).Select(point => FromPoint(point, "point"))],
                FromStyle(Required(document.Style, "style")),
                entityType,
                document.LaneOrder,
                document.StackIndex,
                document.SortOrder),
            "text" => new SceneText(
                Required(document.PrimitiveId, "primitiveId"),
                owner,
                layer,
                Required(document.Text, "text"),
                FromRect(Required(document.TextBounds, "textBounds"), "textBounds"),
                FromStyle(Required(document.Style, "style")),
                (GanttLabelPosition)RequiredInt(document.Alignment, "alignment"),
                entityType,
                document.LaneOrder,
                document.StackIndex,
                document.SortOrder),
            "group" => new SceneGroup(
                Required(document.PrimitiveId, "primitiveId"),
                owner,
                layer,
                document.ChildPrimitiveIds ?? [],
                entityType,
                document.LaneOrder,
                document.StackIndex,
                document.SortOrder),
            _ => throw new InvalidDataException($"Unknown scene primitive kind '{document.Kind}'."),
        };
    }

    private static SceneWarning FromWarning(WarningDocument document) =>
        new(ParseOwner(document.OwnerId), Required(document.Code, "warning.code"), Required(document.Message, "warning.message"));

    private static RectDocument ToRect(RectD rect) => new()
    {
        X = rect.X,
        Y = rect.Y,
        Width = rect.Width,
        Height = rect.Height,
    };

    private static RectD FromRect(RectDocument document, string field) =>
        new(
            RequiredFinite(document.X, field + ".x"),
            RequiredFinite(document.Y, field + ".y"),
            RequiredFinite(document.Width, field + ".width"),
            RequiredFinite(document.Height, field + ".height"));

    private static PointDocument ToPoint(PointD point) => new() { X = point.X, Y = point.Y };

    private static PointD FromPoint(PointDocument document, string field) =>
        new(RequiredFinite(document.X, field + ".x"), RequiredFinite(document.Y, field + ".y"));

    private static StyleDocument ToStyle(SceneStyle style) => new()
    {
        StyleKey = style.StyleKey,
        FillColour = style.FillColour?.ToString(),
        StrokeColour = style.StrokeColour?.ToString(),
        OutlineWidthPt = style.OutlineWidthPt,
        HatchPattern = (int)style.HatchPattern,
        FontFamily = style.FontFamily,
        FontSizePt = style.FontSizePt,
        Bold = style.Bold,
        Alignment = style.Alignment is { } alignment ? (int)alignment : null,
    };

    private static SceneStyle FromStyle(StyleDocument document)
    {
        GanttHatchPattern hatchPattern = Enum.IsDefined(typeof(GanttHatchPattern), document.HatchPattern)
            ? (GanttHatchPattern)document.HatchPattern
            : throw new InvalidDataException($"Invalid hatch pattern {document.HatchPattern}.");
        GanttLabelPosition? alignment = document.Alignment is { } value
            ? Enum.IsDefined(typeof(GanttLabelPosition), value)
                ? (GanttLabelPosition)value
                : throw new InvalidDataException($"Invalid label alignment {value}.")
            : null;
        return new SceneStyle(
            Required(document.StyleKey, "style.styleKey"),
            ParseColour(document.FillColour, "style.fillColour"),
            ParseColour(document.StrokeColour, "style.strokeColour"),
            document.OutlineWidthPt,
            hatchPattern,
            document.FontFamily,
            document.FontSizePt,
            document.Bold,
            alignment);
    }

    private static ColourHex? ParseColour(string? value, string field) =>
        value is null
            ? null
            : ColourHex.TryParse(value, out ColourHex? colour) && colour is not null
                ? colour
                : throw new InvalidDataException($"Invalid colour in {field}.");

    private static GanttRowId ParseOwner(string? value) =>
        GanttRowId.TryParse(value, out GanttRowId? owner) && owner is not null
            ? owner
            : throw new InvalidDataException("Invalid scene owner ID.");

    private static ZLayer ParseLayer(int value) =>
        Enum.IsDefined(typeof(ZLayer), value) ? (ZLayer)value : throw new InvalidDataException($"Invalid z-layer {value}.");

    private static int? ToNullableInt(GanttEntityType? value) => value is { } type ? (int)type : null;

    private static GanttEntityType? FromNullableInt(int? value) =>
        value is null ? null : Enum.IsDefined(typeof(GanttEntityType), value.Value) ? (GanttEntityType)value.Value : throw new InvalidDataException($"Invalid entity type {value}.");

    private static T Required<T>(T? value, string field) where T : class =>
        value ?? throw new InvalidDataException($"Missing scene field '{field}'.");

    private static double RequiredFinite(double value, string field) =>
        double.IsFinite(value) ? value : throw new InvalidDataException($"Non-finite scene field '{field}'.");

    private static int RequiredInt(int? value, string field) =>
        value ?? throw new InvalidDataException($"Missing scene field '{field}'.");

    private sealed class SceneDocument
    {
        public int Version { get; init; }
        public RectDocument? ChartBounds { get; init; }
        public RectDocument? PlotBounds { get; init; }
        public List<PrimitiveDocument>? Primitives { get; init; }
        public List<WarningDocument>? Warnings { get; init; }
    }

    private sealed class PrimitiveDocument
    {
        public string? Kind { get; init; }
        public string? PrimitiveId { get; init; }
        public string? OwnerId { get; init; }
        public int ZLayer { get; init; }
        public int? EntityType { get; init; }
        public int? LaneOrder { get; init; }
        public int? StackIndex { get; init; }
        public int? SortOrder { get; init; }
        public RectDocument? Bounds { get; init; }
        public PointDocument? From { get; init; }
        public PointDocument? To { get; init; }
        public List<PointDocument>? Points { get; init; }
        public string? Text { get; init; }
        public RectDocument? TextBounds { get; init; }
        public StyleDocument? Style { get; init; }
        public int? Alignment { get; init; }
        public List<string>? ChildPrimitiveIds { get; init; }
    }

    private sealed class RectDocument
    {
        public double X { get; init; }
        public double Y { get; init; }
        public double Width { get; init; }
        public double Height { get; init; }
    }

    private sealed class PointDocument
    {
        public double X { get; init; }
        public double Y { get; init; }
    }

    private sealed class StyleDocument
    {
        public string? StyleKey { get; init; }
        public string? FillColour { get; init; }
        public string? StrokeColour { get; init; }
        public double? OutlineWidthPt { get; init; }
        public int HatchPattern { get; init; }
        public string? FontFamily { get; init; }
        public double? FontSizePt { get; init; }
        public bool? Bold { get; init; }
        public int? Alignment { get; init; }
    }

    private sealed class WarningDocument
    {
        public string? OwnerId { get; init; }
        public string? Code { get; init; }
        public string? Message { get; init; }
    }
}
