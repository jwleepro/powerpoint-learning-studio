using System.Runtime.Versioning;

namespace PPTCoach.Core.Interfaces;

/// <summary>
/// Interface for monitoring presentation events
/// </summary>
public interface IPresentationMonitor
{
    /// <summary>
    /// Event fired when the current slide changes
    /// </summary>
    event EventHandler<int>? OnSlideChanged;

    /// <summary>
    /// Event fired when the selected shape changes
    /// </summary>
    event EventHandler<string>? OnShapeSelectionChanged;

    /// <summary>
    /// Event fired when the presentation is saved
    /// </summary>
    event EventHandler<string>? OnPresentationSaved;

    /// <summary>
    /// Starts monitoring for slide changes
    /// </summary>
    [SupportedOSPlatform("windows")]
    void StartMonitoringSlideChanges(object presentation);

    /// <summary>
    /// Stops monitoring for slide changes
    /// </summary>
    [SupportedOSPlatform("windows")]
    void StopMonitoringSlideChanges();

    /// <summary>
    /// Checks for slide changes and fires event if detected
    /// </summary>
    [SupportedOSPlatform("windows")]
    void CheckForSlideChange();

    /// <summary>
    /// Starts monitoring for shape selection changes
    /// </summary>
    [SupportedOSPlatform("windows")]
    void StartMonitoringShapeSelection(object presentation);

    /// <summary>
    /// Stops monitoring for shape selection changes
    /// </summary>
    [SupportedOSPlatform("windows")]
    void StopMonitoringShapeSelection();

    /// <summary>
    /// Checks for shape selection changes and fires event if detected
    /// </summary>
    [SupportedOSPlatform("windows")]
    void CheckForShapeSelectionChange();

    /// <summary>
    /// Starts monitoring for presentation save events
    /// </summary>
    [SupportedOSPlatform("windows")]
    void StartMonitoringPresentationSave(object presentation);

    /// <summary>
    /// Stops monitoring for presentation save events
    /// </summary>
    [SupportedOSPlatform("windows")]
    void StopMonitoringPresentationSave();

    /// <summary>
    /// Checks for presentation save events and fires event if detected
    /// </summary>
    [SupportedOSPlatform("windows")]
    void CheckForPresentationSave();
}
