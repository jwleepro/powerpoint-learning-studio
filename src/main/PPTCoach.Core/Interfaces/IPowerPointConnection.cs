using System.Runtime.Versioning;

namespace PPTCoach.Core.Interfaces;

/// <summary>
/// Interface for PowerPoint connection and instance management
/// </summary>
public interface IPowerPointConnection
{
    /// <summary>
    /// Checks if PowerPoint is installed on the system
    /// </summary>
    [SupportedOSPlatform("windows")]
    bool IsPowerPointInstalled();

    /// <summary>
    /// Gets a running PowerPoint instance if one exists
    /// </summary>
    [SupportedOSPlatform("windows")]
    object? GetRunningPowerPointInstance();

    /// <summary>
    /// Connects to a running PowerPoint instance or throws an exception
    /// </summary>
    [SupportedOSPlatform("windows")]
    object ConnectToPowerPointOrThrow();
}
