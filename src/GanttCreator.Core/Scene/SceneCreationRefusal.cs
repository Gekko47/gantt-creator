namespace GanttCreator.Core.Scene;

/// <summary>The reason a scene could not be created.</summary>
public enum SceneCreationRefusal
{
    /// <summary>The chart or plot bounds contain invalid geometry.</summary>
    InvalidBounds = 0,

    /// <summary>A primitive contains invalid geometry.</summary>
    InvalidPrimitive = 1,

    /// <summary>Two primitives share one identifier.</summary>
    DuplicatePrimitiveId = 2,

    /// <summary>A group references a primitive that is not in the scene.</summary>
    UnresolvedGroupChild = 3,
}
