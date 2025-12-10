using PPTCoach.Core;
using PPTCoach.Tests.Utils;
using System.Runtime.Versioning;

namespace PPTCoach.Tests.Phase01;

/// <summary>
/// Tests for Phase 1.3: Slide Information Reading
/// </summary>
public class PowerPointSlideInfoTests
{
    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetCurrentSlideNumber()
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

            // Act
            int currentSlideNumber = powerPointService.GetCurrentSlideNumber(presentation);

            // Assert
            Assert.Equal(1, currentSlideNumber);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReturn0WhenNoSlideIsSelected()
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

            // Create a new presentation (without selecting any slide in slide sorter view)
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Deselect any slide by going to slide sorter view without selection
            dynamic pptApp = instance;
            dynamic activeWindow = pptApp.ActiveWindow;

            // Switch to slide sorter view (where no slide might be selected)
            // ViewType: 1 = Normal, 2 = Outline, 3 = SlideSorter, 7 = NotesPage, 9 = ReadingView
            activeWindow.ViewType = 3; // Slide Sorter View

            // Clear selection in slide sorter view
            dynamic selection = activeWindow.Selection;
            selection.Unselect();

            // Act
            int currentSlideNumber = powerPointService.GetCurrentSlideNumber(presentation);

            // Assert
            Assert.Equal(0, currentSlideNumber);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetTotalSlideCount()
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

            // Act
            int totalSlideCount = powerPointService.GetTotalSlideCount(presentation);

            // Assert
            Assert.Equal(1, totalSlideCount);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReadSlideTitleText()
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

            // Get the first slide and replace it with a title slide layout
            dynamic pres = presentation;
            dynamic slides = pres.Slides;

            // Delete the blank slide and add a slide with title layout
            slides[1].Delete();
            dynamic slide = slides.Add(1, 1); // 1 = ppLayoutTitle

            // Set a title on the slide
            dynamic shapes = slide.Shapes;
            dynamic titleShape = shapes[1]; // First shape is typically the title placeholder
            dynamic textFrame = titleShape.TextFrame;
            dynamic textRange = textFrame.TextRange;
            textRange.Text = "Test Slide Title";

            // Act
            string titleText = powerPointService.GetSlideTitle(slide);

            // Assert
            Assert.Equal("Test Slide Title", titleText);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReadAllTextContentFromSlide()
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

            // Get the first slide and add a title and content layout
            dynamic pres = presentation;
            dynamic slides = pres.Slides;

            // Delete the blank slide and add a slide with title and content layout
            slides[1].Delete();
            dynamic slide = slides.Add(1, 2); // 2 = ppLayoutText (Title and Content)

            // Set a title and body text on the slide
            dynamic shapes = slide.Shapes;
            dynamic titleShape = shapes[1]; // First shape is typically the title placeholder
            dynamic titleTextFrame = titleShape.TextFrame;
            dynamic titleTextRange = titleTextFrame.TextRange;
            titleTextRange.Text = "Test Title";

            dynamic contentShape = shapes[2]; // Second shape is typically the content placeholder
            dynamic contentTextFrame = contentShape.TextFrame;
            dynamic contentTextRange = contentTextFrame.TextRange;
            contentTextRange.Text = "Line 1\nLine 2\nLine 3";

            // Act
            string allText = powerPointService.GetAllTextFromSlide(slide);

            // Assert
            Assert.Contains("Test Title", allText);
            Assert.Contains("Line 1", allText);
            Assert.Contains("Line 2", allText);
            Assert.Contains("Line 3", allText);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetAllShapesOnSlide()
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

            // Get the first slide and add a title and content layout
            dynamic pres = presentation;
            dynamic slides = pres.Slides;

            // Delete the blank slide and add a slide with title and content layout
            slides[1].Delete();
            dynamic slide = slides.Add(1, 2); // 2 = ppLayoutText (Title and Content)

            // Act
            var shapes = powerPointService.GetShapesOnSlide(slide);

            // Assert
            Assert.NotNull(shapes);
            Assert.True(shapes.Count > 0, "Slide should have at least one shape");
            // Title and Content layout typically has 2 placeholders (title and content)
            Assert.Equal(2, shapes.Count);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReturnShapeType()
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

            // Get the first slide and add a title and content layout
            dynamic pres = presentation;
            dynamic slides = pres.Slides;

            // Delete the blank slide and add a slide with title and content layout
            slides[1].Delete();
            dynamic slide = slides.Add(1, 2); // 2 = ppLayoutText (Title and Content)

            // Get the shapes on the slide
            var shapes = powerPointService.GetShapesOnSlide(slide);
            Assert.NotNull(shapes);
            Assert.True(shapes.Count > 0);

            // Act - Get the type of the first shape (should be a placeholder/text box)
            string shapeType = powerPointService.GetShapeType(shapes[0]);

            // Assert
            Assert.NotNull(shapeType);
            Assert.NotEmpty(shapeType);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }
}
