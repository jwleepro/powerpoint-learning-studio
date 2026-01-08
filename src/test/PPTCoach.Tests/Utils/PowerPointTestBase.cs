using PPTCoach.Core;
using System.Diagnostics;

namespace PPTCoach.Tests.Utils;

/// <summary>
/// Base class for PowerPoint tests that provides common setup and cleanup functionality
/// </summary>
public abstract class PowerPointTestBase : IDisposable
{
    protected PowerPointService PowerPointService { get; }
    protected Process? PptProcess { get; private set; }
    protected object? Instance { get; private set; }
    protected object? Presentation { get; private set; }
    protected dynamic? FirstSlide { get; private set; }

    protected PowerPointTestBase()
    {
        PowerPointService = new PowerPointService();
        PowerPointTestHelpers.EnsureNoPowerPointRunning(PowerPointService);
    }

    /// <summary>
    /// Sets up PowerPoint with a new presentation and gets the first slide
    /// </summary>
    /// <param name="windowStyle">Window style for PowerPoint (default: Minimized)</param>
    protected void SetupPowerPointWithPresentation(
        ProcessWindowStyle windowStyle = ProcessWindowStyle.Minimized)
    {
        (PptProcess, Instance) = PowerPointTestHelpers.StartPowerPointAndWait(
            PowerPointService, windowStyle);

        Presentation = PowerPointService.CreateNewPresentation(Instance);
        dynamic pres = Presentation;
        FirstSlide = pres.Slides[1];
    }

    /// <summary>
    /// Gets the shapes collection from the first slide
    /// </summary>
    protected dynamic GetFirstSlideShapes()
    {
        if (FirstSlide == null)
        {
            throw new InvalidOperationException(
                "First slide is not available. Call SetupPowerPointWithPresentation first.");
        }
        return FirstSlide.Shapes;
    }

    public void Dispose()
    {
        PowerPointTestHelpers.CleanupPowerPoint(PowerPointService, PptProcess);
        GC.SuppressFinalize(this);
    }
}
