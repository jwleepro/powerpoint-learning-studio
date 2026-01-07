using System.Runtime.Versioning;

namespace PPTCoach.Core.Interfaces;

/// <summary>
/// Interface for querying font properties
/// </summary>
public interface IFontQuery
{
    /// <summary>
    /// Gets the font name from a shape
    /// </summary>
    [SupportedOSPlatform("windows")]
    string? GetFontName(object shape);

    /// <summary>
    /// Gets the font size from a shape
    /// </summary>
    [SupportedOSPlatform("windows")]
    float GetFontSize(object shape);

    /// <summary>
    /// Gets the font color from a shape as RGB values
    /// </summary>
    [SupportedOSPlatform("windows")]
    (int red, int green, int blue) GetFontColor(object shape);
}
