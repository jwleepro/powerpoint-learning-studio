using System.Runtime.Versioning;

namespace PPTCoach.Core.Interfaces;

/// <summary>
/// Interface for creating and managing PowerPoint presentations
/// </summary>
public interface IPresentationManager
{
    /// <summary>
    /// Creates a new presentation with a default blank slide
    /// </summary>
    [SupportedOSPlatform("windows")]
    object CreateNewPresentation(object powerPointInstance);
}
