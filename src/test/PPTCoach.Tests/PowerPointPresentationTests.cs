using PPTCoach.Core;
using System.Runtime.Versioning;

namespace PPTCoach.Tests;

/// <summary>
/// Tests for Phase 1.2: New Document Creation
/// </summary>
public class PowerPointPresentationTests
{
    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldCreateNewPowerPointPresentation()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        PowerPointTestHelpers.EnsureNoPowerPointRunning(powerPointService);

        try
        {
            // Start PowerPoint and wait for it to be ready
            object instance;
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(powerPointService);

            Assert.NotNull(instance);

            // Act
            var presentation = powerPointService.CreateNewPresentation(instance);

            // Assert
            Assert.NotNull(presentation);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReturnPresentationObjectAfterCreation()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        PowerPointTestHelpers.EnsureNoPowerPointRunning(powerPointService);

        try
        {
            // Start PowerPoint and wait for it to be ready
            object instance;
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(powerPointService);

            Assert.NotNull(instance);

            // Act
            var presentation = powerPointService.CreateNewPresentation(instance);

            // Assert
            Assert.NotNull(presentation);

            // Verify it's a valid presentation object by checking its properties
            dynamic pres = presentation;

            // Should have Slides collection
            dynamic slides = pres.Slides;
            Assert.NotNull(slides);

            // Should have a Name property
            string name = pres.Name;
            Assert.NotNull(name);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldCreatePresentationWithDefaultBlankSlide()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        PowerPointTestHelpers.EnsureNoPowerPointRunning(powerPointService);

        try
        {
            // Start PowerPoint and wait for it to be ready
            object instance;
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(powerPointService);

            Assert.NotNull(instance);

            // Act
            var presentation = powerPointService.CreateNewPresentation(instance);

            // Assert
            Assert.NotNull(presentation);

            // Verify it has a default blank slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            int slideCount = slides.Count;

            Assert.Equal(1, slideCount);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldHandleCreationFailureGracefully()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        var invalidInstance = new object(); // Invalid PowerPoint instance

        // Act & Assert
        var exception = Assert.Throws<InvalidOperationException>(() =>
            powerPointService.CreateNewPresentation(invalidInstance));

        // Verify the exception contains meaningful information about the failure
        Assert.NotNull(exception);
        Assert.Contains("presentation", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
