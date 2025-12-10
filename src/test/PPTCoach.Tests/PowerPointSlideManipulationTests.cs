using PPTCoach.Core;
using System.Runtime.Versioning;

namespace PPTCoach.Tests;

/// <summary>
/// Tests for Phase 1.4: Slide Manipulation
/// </summary>
public class PowerPointSlideManipulationTests
{
    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldAddNewBlankSlide()
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

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get initial slide count
            int initialSlideCount = powerPointService.GetTotalSlideCount(presentation);

            // Act
            var newSlide = powerPointService.AddBlankSlide(presentation);

            // Assert
            Assert.NotNull(newSlide);
            int finalSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(initialSlideCount + 1, finalSlideCount);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldAddSlideWithSpecificLayout()
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

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get initial slide count
            int initialSlideCount = powerPointService.GetTotalSlideCount(presentation);

            // Act
            // 1 = ppLayoutTitle (Title slide layout)
            var newSlide = powerPointService.AddSlideWithLayout(presentation, 1);

            // Assert
            Assert.NotNull(newSlide);
            int finalSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(initialSlideCount + 1, finalSlideCount);

            // Verify the slide has the correct layout by checking for title placeholder
            dynamic slide = newSlide;
            dynamic shapes = slide.Shapes;
            Assert.True(shapes.Count > 0, "Title slide layout should have at least one placeholder");
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldDeleteSlideByIndex()
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

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Add a second slide so we have multiple slides
            powerPointService.AddBlankSlide(presentation);

            // Get initial slide count (should be 2)
            int initialSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(2, initialSlideCount);

            // Act - Delete the second slide (index 2)
            powerPointService.DeleteSlideByIndex(presentation, 2);

            // Assert
            int finalSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(1, finalSlideCount);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldMoveSlideFromOnePositionToAnother()
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

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Add slides with specific layouts to identify them
            // Slide 1: Default blank (already exists)
            // Slide 2: Title layout
            var slide2 = powerPointService.AddSlideWithLayout(presentation, 1); // Title layout
            // Slide 3: Title and Content layout
            var slide3 = powerPointService.AddSlideWithLayout(presentation, 2); // Title and Content layout

            // Verify we have 3 slides
            int totalSlides = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(3, totalSlides);

            // Get the layout of slide at position 1 before moving
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slideAtPos1Before = slides[1];
            int layoutBeforeMove = slideAtPos1Before.Layout;

            // Act - Move slide from position 2 to position 1
            powerPointService.MoveSlide(presentation, 2, 1);

            // Assert
            // Total count should remain the same
            int finalSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(3, finalSlideCount);

            // The slide that was at position 2 should now be at position 1
            dynamic slideAtPos1After = slides[1];
            int layoutAfterMove = slideAtPos1After.Layout;
            Assert.Equal(1, layoutAfterMove); // Title layout (which was at position 2)
            Assert.NotEqual(layoutBeforeMove, layoutAfterMove);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldPreventDeletingWhenOnlyOneSlideExists()
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

            // Create a new presentation with only one slide
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Verify we have exactly 1 slide
            int initialSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(1, initialSlideCount);

            // Act & Assert - Attempting to delete the only slide should throw an exception
            var exception = Assert.Throws<InvalidOperationException>(() =>
                powerPointService.DeleteSlideByIndex(presentation, 1));

            // Verify the exception contains meaningful information
            Assert.NotNull(exception);
            Assert.Contains("last slide", exception.Message, StringComparison.OrdinalIgnoreCase);

            // Verify the slide was not deleted
            int finalSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(1, finalSlideCount);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }
}
