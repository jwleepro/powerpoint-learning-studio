using PPTCoach.Core;
using PPTCoach.Tests.Utils;
using System.Runtime.Versioning;

namespace PPTCoach.Tests.Phase01;

/// <summary>
/// Tests for Phase 1.6: Event Detection
/// </summary>
public class PowerPointEventDetectionTests
{
    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldDetectSlideChangeEvent()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;
        dynamic? instance = null;
        bool slideChangeDetected = false;
        int initialSlideNumber = 0;
        int newSlideNumber = 0;

        try
        {
            // Ensure no PowerPoint is running initially
            PowerPointTestHelpers.EnsureNoPowerPointRunning(powerPointService);

            // Start PowerPoint and wait for it to be ready
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(powerPointService);
            Assert.NotNull(instance);

            // Create a new presentation (it will have 1 default slide from CreateNewPresentation)
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Add a second slide so we can test slide change
            powerPointService.AddBlankSlide(presentation);

            // Get initial slide number
            initialSlideNumber = powerPointService.GetCurrentSlideNumber(presentation);

            // Subscribe to slide change event
            powerPointService.OnSlideChanged += (sender, slideNumber) =>
            {
                slideChangeDetected = true;
                newSlideNumber = slideNumber;
            };

            // Start monitoring slide changes
            powerPointService.StartMonitoringSlideChanges(presentation);

            // Act - Navigate to the second slide using ActiveWindow.View
            dynamic pres = presentation;
            dynamic app = pres.Application;
            dynamic activeWindow = app.ActiveWindow;

            // Set window to Normal view and navigate to slide 2
            activeWindow.ViewType = 1; // ppViewNormal
            dynamic view = activeWindow.View;
            dynamic slides = pres.Slides;
            dynamic slide2 = slides[2];
            view.GotoSlide(slide2.SlideIndex);

            // Give PowerPoint time to complete the navigation
            System.Threading.Thread.Sleep(500);

            // Check for slide change (simulates polling mechanism)
            powerPointService.CheckForSlideChange();

            // Assert
            Assert.True(slideChangeDetected, "Slide change event should have been detected");
            Assert.Equal(1, initialSlideNumber);
            Assert.Equal(2, newSlideNumber);
        }
        finally
        {
            powerPointService.StopMonitoringSlideChanges();
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldDetectShapeSelectionChangeEvent()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;
        dynamic? instance = null;
        bool shapeSelectionChangeDetected = false;
        string? selectedShapeName = null;

        try
        {
            // Ensure no PowerPoint is running initially
            PowerPointTestHelpers.EnsureNoPowerPointRunning(powerPointService);

            // Start PowerPoint and wait for it to be ready
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(powerPointService);
            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Add some shapes to the slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];
            dynamic shapes = slide.Shapes;

            // Add two text boxes
            dynamic shape1 = shapes.AddTextbox(1, 100, 100, 200, 50); // msoTextOrientationHorizontal = 1
            shape1.Name = "TestShape1";
            shape1.TextFrame.TextRange.Text = "Shape 1";

            dynamic shape2 = shapes.AddTextbox(1, 100, 200, 200, 50);
            shape2.Name = "TestShape2";
            shape2.TextFrame.TextRange.Text = "Shape 2";

            // Subscribe to shape selection change event
            powerPointService.OnShapeSelectionChanged += (sender, shapeName) =>
            {
                shapeSelectionChangeDetected = true;
                selectedShapeName = shapeName;
            };

            // Start monitoring shape selection changes
            powerPointService.StartMonitoringShapeSelection(presentation);

            // Act - Select a shape
            shape1.Select();
            System.Threading.Thread.Sleep(500);

            // Check for shape selection change (simulates polling mechanism)
            powerPointService.CheckForShapeSelectionChange();

            // Assert
            Assert.True(shapeSelectionChangeDetected, "Shape selection change event should have been detected");
            Assert.Equal("TestShape1", selectedShapeName);
        }
        finally
        {
            powerPointService.StopMonitoringShapeSelection();
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }
}
