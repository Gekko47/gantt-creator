using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R6.2 D4: the staging ownership scheme is distinct from the live chart's scheme, and
/// each scheme's scans are structurally blind to the other's shapes.
/// </summary>
/// <remarks>
/// The writer facts share one modelled shape collection between a live-scheme writer and
/// a staging-scheme writer, which is the production situation the work item's D4 exists
/// for: staged shapes and live shapes on the same worksheet. The load-bearing assertions
/// are the two <c>ListOwned</c> blindness facts and the mutual delete refusals - those
/// are what stop a coalesced refresh from deleting a staged shape mid-composition and a
/// staging cleanup from touching the user's chart.
/// </remarks>
public sealed class ShapeStagingOwnershipTests
{
    private const string _stagedPrimitiveId = "chart:panel-cell:Type";

    [Fact]
    public void The_staging_tag_is_a_distinct_bounded_scheme()
    {
        string stagedName = ShapeStagingTag.NamePrefix + _stagedPrimitiveId;
        string tag = ShapeStagingTag.ForStagedName(stagedName);

        Assert.StartsWith(ShapeStagingTag.Prefix, tag, StringComparison.Ordinal);
        Assert.Equal(ShapeStagingTag.Prefix.Length + ShapeStagingTag.HashLength, tag.Length);
        Assert.True(ShapeStagingTag.IsStagingTag(tag));

        // The two families can never classify each other's tags.
        Assert.False(ShapeOwnershipTag.IsOwnedTag(tag));
        Assert.False(ShapeStagingTag.IsStagingTag(ShapeOwnershipTag.ForPrimitiveId(_stagedPrimitiveId)));

        // Ordinal, culture-invariant, deterministic.
        Assert.Equal(tag, ShapeStagingTag.ForStagedName(stagedName));
    }

    [Fact]
    public void The_group_identity_is_a_staging_shape_that_cannot_collide_with_a_child()
    {
        // R6.4 D1: the group is a temporary staging shape with its own deterministic name.
        Assert.Equal(ShapeStagingTag.NamePrefix + "export-group", ShapeStagingTag.GroupName);

        // The colon-free identifier is structurally distinct from every scene primitive
        // id (which always carries the {ownerId}:{role} separator), so the group's name
        // can never collide with a staged child's name.
        Assert.False(ShapeStagingTag.GroupPrimitiveId.Contains(':', StringComparison.Ordinal));
        Assert.True(ShapeStagingTag.GroupName.Contains(':', StringComparison.Ordinal));

        // The full name is in the form the ownership scan reaches, and its tag is the
        // scheme's own recomputation - so the group is provably staging-owned and can
        // never be mistaken for the live chart's.
        Assert.True(ExcelShapeWriter.CouldBeOwnedName(ShapeStagingTag.GroupName));
        string tag = ShapeStagingTag.ForStagedName(ShapeStagingTag.GroupName);
        Assert.True(ShapeStagingTag.IsStagingTag(tag));
        Assert.False(ShapeOwnershipTag.IsOwnedTag(tag));
        Assert.Equal(tag, ShapeStagingTag.ForStagedName(ShapeStagingTag.GroupName));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("GanttCreator.Staging.v1:")]
    [InlineData("GanttCreator.Staging.v1:0123")]
    [InlineData("GanttCreator.Staging.v1:0123456789ABCDEZ")]
    public void A_malformed_staging_tag_is_not_staging_owned(string? alternativeText) =>
        Assert.False(ShapeStagingTag.IsStagingTag(alternativeText));

    [Fact]
    public void The_schemes_compose_the_expected_names_and_tags()
    {
        // Live: identity name, the existing owned tag.
        Assert.Equal(_stagedPrimitiveId, ShapeOwnershipScheme.Live.NameFor(_stagedPrimitiveId));
        Assert.Equal(
            ShapeOwnershipTag.ForPrimitiveId(_stagedPrimitiveId),
            ShapeOwnershipScheme.Live.TagForName(_stagedPrimitiveId));
        Assert.False(ShapeOwnershipScheme.Live.IsStaging);

        // Staging: prefixed name, staging tag derived from the full name.
        string stagedName = ShapeOwnershipScheme.Staging.NameFor(_stagedPrimitiveId);
        Assert.Equal(ShapeStagingTag.NamePrefix + _stagedPrimitiveId, stagedName);
        Assert.Equal(ShapeStagingTag.ForStagedName(stagedName), ShapeOwnershipScheme.Staging.TagForName(stagedName));
        Assert.True(ShapeOwnershipScheme.Staging.IsStaging);
    }

    [Fact]
    public void A_scheme_refuses_a_blank_name_input()
    {
        _ = Assert.Throws<ArgumentNullException>(() => ShapeOwnershipScheme.Live.NameFor(null!));
        _ = Assert.Throws<ArgumentException>(() => ShapeOwnershipScheme.Live.NameFor(" "));
        _ = Assert.Throws<ArgumentException>(() => ShapeOwnershipScheme.Staging.TagForName(string.Empty));
        _ = Assert.Throws<ArgumentNullException>(() => ShapeStagingTag.ForStagedName(null!));
    }

    [Fact]
    public void A_staging_create_stamps_the_staging_prefix_and_tag()
    {
        SharedShapes sheet = new();
        SchemeWriter staging = new(sheet, ShapeOwnershipScheme.Staging);

        ShapeWriteOutcome outcome = staging.Create(RectangleRequest(_stagedPrimitiveId));

        Assert.True(outcome.Succeeded, "Create refused: " + outcome.Refusal);
        Excel.Shape created = Assert.Single(sheet.Shapes);
        Assert.Equal(ShapeStagingTag.NamePrefix + _stagedPrimitiveId, created.Name);
        Assert.Equal(ShapeStagingTag.ForStagedName(created.Name), created.AlternativeText);
        Assert.True(ShapeStagingTag.IsStagingTag(created.AlternativeText));
    }

    [Fact]
    public void A_live_scan_is_blind_to_staged_shapes()
    {
        SharedShapes sheet = new();
        SchemeWriter live = new(sheet, ShapeOwnershipScheme.Live);
        SchemeWriter staging = new(sheet, ShapeOwnershipScheme.Staging);

        Assert.True(live.Create(RectangleRequest("row-1:bar")).Succeeded);
        Assert.True(staging.Create(RectangleRequest(_stagedPrimitiveId)).Succeeded);

        // The reconciler consumes ListOwned, so this is the proof that a coalesced
        // refresh during staging never sees - and therefore never deletes - a staged
        // shape, and never mistakes one for a stale live shape.
        Assert.Equal(["row-1:bar"], live.ListOwned());
        Assert.Equal([ShapeStagingTag.NamePrefix + _stagedPrimitiveId], staging.ListOwned());
    }

    [Fact]
    public void A_staging_delete_refuses_a_live_owned_shape()
    {
        SharedShapes sheet = new();
        SchemeWriter live = new(sheet, ShapeOwnershipScheme.Live);
        SchemeWriter staging = new(sheet, ShapeOwnershipScheme.Staging);
        Assert.True(live.Create(RectangleRequest("row-1:bar")).Succeeded);

        ShapeWriteOutcome outcome = staging.Delete("row-1:bar");

        Assert.False(outcome.Succeeded);
        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
        Assert.Equal(["row-1:bar"], live.ListOwned());
    }

    [Fact]
    public void A_live_delete_refuses_a_staged_shape()
    {
        SharedShapes sheet = new();
        SchemeWriter live = new(sheet, ShapeOwnershipScheme.Live);
        SchemeWriter staging = new(sheet, ShapeOwnershipScheme.Staging);
        Assert.True(staging.Create(RectangleRequest(_stagedPrimitiveId)).Succeeded);

        ShapeWriteOutcome outcome = live.Delete(_stagedPrimitiveId);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
        _ = Assert.Single(sheet.Shapes);
    }

    [Fact]
    public void Staging_cleanup_removes_only_staged_shapes()
    {
        SharedShapes sheet = new();
        SchemeWriter live = new(sheet, ShapeOwnershipScheme.Live);
        SchemeWriter staging = new(sheet, ShapeOwnershipScheme.Staging);
        Assert.True(live.Create(RectangleRequest("row-1:bar")).Succeeded);
        Assert.True(staging.Create(RectangleRequest(_stagedPrimitiveId)).Succeeded);

        Assert.True(staging.Delete(_stagedPrimitiveId).Succeeded);

        // The live chart is untouched, still owned, and still the only shape left.
        Assert.Equal(["row-1:bar"], live.ListOwned());
        _ = Assert.Single(sheet.Shapes);
    }

    [Fact]
    public void A_second_staging_create_of_the_same_primitive_is_refused_as_already_exists()
    {
        // The composer pre-flights an empty staging collection; this is the writer
        // behaviour that makes the refusal safe if a caller skips that pre-flight.
        SharedShapes sheet = new();
        SchemeWriter staging = new(sheet, ShapeOwnershipScheme.Staging);
        Assert.True(staging.Create(RectangleRequest(_stagedPrimitiveId)).Succeeded);

        ShapeWriteOutcome outcome = staging.Create(RectangleRequest(_stagedPrimitiveId));

        Assert.Equal(ShapeWriteRefusal.AlreadyExists, outcome.Refusal);
    }

    [Fact]
    public void A_live_shape_and_a_staged_shape_can_share_the_primitive_identifier()
    {
        // The staged name carries the prefix, so the live name is not occupied: a
        // composition staged beside a chart whose panel primitives are absent live
        // (the live profile draws no panel) still cannot collide even if a future
        // live scene emitted the same identifier.
        SharedShapes sheet = new();
        SchemeWriter live = new(sheet, ShapeOwnershipScheme.Live);
        SchemeWriter staging = new(sheet, ShapeOwnershipScheme.Staging);

        Assert.True(staging.Create(RectangleRequest(_stagedPrimitiveId)).Succeeded);
        ShapeWriteOutcome liveCreate = live.Create(RectangleRequest(_stagedPrimitiveId));

        Assert.True(liveCreate.Succeeded, "Live create refused: " + liveCreate.Refusal);
        Assert.Equal(2, sheet.Shapes.Count);
        Assert.Equal([_stagedPrimitiveId], live.ListOwned());
        Assert.Equal([ShapeStagingTag.NamePrefix + _stagedPrimitiveId], staging.ListOwned());
    }

    private static OfficeShapeRequest RectangleRequest(string id) =>
        StagingSchemeFixtures.RectangleRequest(id);
}
