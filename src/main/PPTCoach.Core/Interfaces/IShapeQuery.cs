using System.Runtime.Versioning;

namespace PPTCoach.Core.Interfaces;

/// <summary>
/// Interface for querying shape properties
/// </summary>
public interface IShapeQuery
{
    /// <summary>
    /// Gets the type of a shape
    /// </summary>
    [SupportedOSPlatform("windows")]
    string GetShapeType(object shape);

    /// <summary>
    /// Gets the fill color of a shape as RGB values
    /// </summary>
    [SupportedOSPlatform("windows")]
    (int red, int green, int blue) GetShapeFillColor(object shape);

    /// <summary>
    /// Gets the position of a shape
    /// </summary>
    [SupportedOSPlatform("windows")]
    (float left, float top) GetShapePosition(object shape);

    /// <summary>
    /// Gets the size of a shape
    /// </summary>
    [SupportedOSPlatform("windows")]
    (float width, float height) GetShapeSize(object shape);

    /// <summary>
    /// Gets the content of a table cell
    /// </summary>
    [SupportedOSPlatform("windows")]
    string GetTableCellContent(object tableShape, int row, int column);
}
