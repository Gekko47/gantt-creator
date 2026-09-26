namespace GanttCreator.Core.Scene;

/// <summary>A single scene invariant finding.</summary>
/// <param name="Code">The stable finding code, used by tests and diagnostics.</param>
/// <param name="PrimitiveId">The offending primitive identifier, or <see langword="null"/> for a scene-level finding.</param>
/// <param name="Message">The deterministic human-readable explanation.</param>
public sealed record SceneValidationFinding(string Code, string? PrimitiveId, string Message);

/// <summary>The zero-throwing result of inspecting a candidate scene.</summary>
/// <param name="Findings">The findings in deterministic discovery order.</param>
public sealed record SceneValidationReport(IReadOnlyList<SceneValidationFinding> Findings)
{
    /// <summary>The finding code for a non-finite line endpoint.</summary>
    public const string NonFiniteLineEndpointCode = "NonFiniteLineEndpoint";

    /// <summary>The finding code for a label outside the chart bounds.</summary>
    public const string LabelOutsideChartCode = "LabelOutsideChart";

    /// <summary>The finding code for a repeated warning for one owner and code.</summary>
    public const string DuplicateWarningCode = "DuplicateWarning";

    /// <summary>The finding code for a primitive ID that is not <c>owner:role</c>.</summary>
    public const string MalformedPrimitiveIdCode = "MalformedPrimitiveId";

    /// <summary>Gets whether the candidate scene is invariant-clean.</summary>
    public bool IsClean => Findings.Count == 0;
}

/// <summary>
/// Reports scene-invariant findings over a *candidate* primitive list without
/// throwing, so a broken scene is diagnosable rather than merely refused.
/// </summary>
/// <remarks>
/// <para>
/// This is deliberately a pre-construction inspector. Most geometry invariants are
/// already enforced at construction: <see cref="RectD"/> throws on a non-finite or
/// negative extent, <see cref="ScenePolygon"/> throws on fewer than three points or a
/// non-finite point, and <see cref="GanttScene.TryCreate"/> refuses duplicate
/// primitive IDs, unresolved group children, and cycles. Those are *unreachable*
/// here, so re-implementing them would be dead code that could never be positively
/// tested, and AGENTS.md requires every new <c>if (bad) { error }</c> to ship with a
/// positive test. This type therefore owns only the invariants nothing else can
/// express.
/// </para>
/// <para>
/// Every finding is emitted in deterministic discovery order: primitive-ID shape,
/// then per-primitive geometry in input order, then warnings in input order.
/// </para>
/// </remarks>
public static class SceneValidator
{
    /// <summary>Inspects a candidate primitive list and reports invariant findings.</summary>
    /// <param name="primitives">The candidate primitives, in any order.</param>
    /// <param name="chartBounds">The chart bounds labels must stay inside.</param>
    /// <param name="warnings">The candidate warnings, or <see langword="null"/> when none are supplied.</param>
    /// <returns>A zero-throwing validation report.</returns>
    public static SceneValidationReport Validate(
        IReadOnlyList<ScenePrimitive>? primitives,
        RectD chartBounds,
        IReadOnlyList<SceneWarning>? warnings = null)
    {
        List<SceneValidationFinding> findings = [];
        if (primitives is null)
        {
            return new SceneValidationReport(findings);
        }

        foreach (ScenePrimitive primitive in primitives)
        {
            if (primitive is not null && !IsWellFormedId(primitive.PrimitiveId))
            {
                findings.Add(new SceneValidationFinding(
                    SceneValidationReport.MalformedPrimitiveIdCode,
                    primitive.PrimitiveId,
                    "A primitive identifier must be an owner identifier followed by ':' and a nonblank role."));
            }
        }

        foreach (ScenePrimitive primitive in primitives)
        {
            if (primitive is null)
            {
                continue;
            }

            // PointD performs no validation of its own, so this is the one geometry
            // type that can carry a NaN or Infinity into a constructed primitive.
            if (primitive is SceneLine line && (!IsFinite(line.From) || !IsFinite(line.To)))
            {
                findings.Add(new SceneValidationFinding(
                    SceneValidationReport.NonFiniteLineEndpointCode,
                    line.PrimitiveId,
                    "A line endpoint must have finite coordinates."));
            }

            if (primitive is SceneText text && !Contains(chartBounds, text.TextBounds))
            {
                findings.Add(new SceneValidationFinding(
                    SceneValidationReport.LabelOutsideChartCode,
                    text.PrimitiveId,
                    "A label must be fully contained by the chart bounds."));
            }
        }

        AddDuplicateWarningFindings(warnings, findings);
        return new SceneValidationReport(findings);
    }

    /// <summary>Inspects a built scene and reports invariant findings.</summary>
    /// <param name="scene">The built scene.</param>
    /// <returns>A zero-throwing validation report.</returns>
    public static SceneValidationReport Validate(GanttScene? scene)
    {
        ArgumentNullException.ThrowIfNull(scene);
        return Validate(scene.Primitives, scene.ChartBounds, scene.Warnings);
    }

    private static void AddDuplicateWarningFindings(
        IReadOnlyList<SceneWarning>? warnings,
        List<SceneValidationFinding> findings)
    {
        if (warnings is null)
        {
            return;
        }

        // Warnings are keyed by owner and code, not by message: the same condition
        // reaching the same row twice is one finding however it was worded, and two
        // genuinely different messages for one code would still be a duplicate the
        // user must not see twice.
        var seen = new HashSet<(SceneOwnerId Owner, string Code)>();
        foreach (SceneWarning warning in warnings)
        {
            if (warning is null)
            {
                continue;
            }

            if (!seen.Add((warning.OwnerId, warning.Code)))
            {
                findings.Add(new SceneValidationFinding(
                    SceneValidationReport.DuplicateWarningCode,
                    null,
                    $"The warning '{warning.Code}' was raised more than once for the same owner."));
            }
        }
    }

    private static bool IsWellFormedId(string primitiveId)
    {
        var separator = primitiveId.LastIndexOf(':');
        return separator > 0
            && separator < primitiveId.Length - 1
            && primitiveId.Length > 0;
    }

    private static bool Contains(RectD bounds, RectD candidate) =>
        bounds.Contains(new PointD(candidate.Left, candidate.Top))
        && bounds.Contains(new PointD(candidate.Right, candidate.Bottom));

    private static bool IsFinite(PointD point) => double.IsFinite(point.X) && double.IsFinite(point.Y);
}
