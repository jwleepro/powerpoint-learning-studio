using System.Runtime.Versioning;

namespace PPTCoach.Core.Interfaces;

/// <summary>
/// Interface for querying slide information
/// </summary>
public interface ISlideQuery
{
    /// <summary>
    /// Gets the current slide number
    /// </summary>
    [SupportedOSPlatform("windows")]
    int GetCurrentSlideNumber(object presentation);

    /// <summary>
    /// Gets the total number of slides in the presentation
    /// </summary>
    [SupportedOSPlatform("windows")]
    int GetTotalSlideCount(object presentation);

    /// <summary>
    /// Gets the title of a slide
    /// </summary>
    [SupportedOSPlatform("windows")]
    string GetSlideTitle(object slide);

    /// <summary>
    /// Gets all text content from a slide
    /// </summary>
    [SupportedOSPlatform("windows")]
    string GetAllTextFromSlide(object slide);

    /// <summary>
    /// Gets all shapes on a slide
    /// </summary>
    [SupportedOSPlatform("windows")]
    List<object> GetShapesOnSlide(object slide);
}
