using System.Runtime.Versioning;

namespace PPTCoach.Core.Interfaces;

/// <summary>
/// Interface for slide manipulation operations
/// </summary>
public interface ISlideOperations
{
    /// <summary>
    /// Adds a blank slide to the presentation
    /// </summary>
    [SupportedOSPlatform("windows")]
    object AddBlankSlide(object presentation);

    /// <summary>
    /// Adds a slide with a specific layout type
    /// </summary>
    [SupportedOSPlatform("windows")]
    object AddSlideWithLayout(object presentation, int layoutType);

    /// <summary>
    /// Deletes a slide at the specified index
    /// </summary>
    [SupportedOSPlatform("windows")]
    void DeleteSlideByIndex(object presentation, int index);

    /// <summary>
    /// Moves a slide from one index to another
    /// </summary>
    [SupportedOSPlatform("windows")]
    void MoveSlide(object presentation, int fromIndex, int toIndex);
}
