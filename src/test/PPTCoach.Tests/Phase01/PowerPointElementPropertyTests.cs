using PPTCoach.Core;
using PPTCoach.Tests.Utils;
using System.Runtime.Versioning;

namespace PPTCoach.Tests.Phase01;

/// <summary>
/// Tests for Phase 1.5: Element Property Reading
/// </summary>
public class PowerPointElementPropertyTests
{
    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetFontNameFromTextShape()
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
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(
                powerPointService, System.Diagnostics.ProcessWindowStyle.Minimized);

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a text shape with specific font
            dynamic shapes = slide.Shapes;
            dynamic textBox = shapes.AddTextbox(1, 100, 100, 200, 50); // 1 = msoTextOrientationHorizontal
            dynamic textFrame = textBox.TextFrame;
            dynamic textRange = textFrame.TextRange;
            textRange.Text = "Test Text";

            // Set font name to a known value
            dynamic font = textRange.Font;
            font.Name = "Arial";

            // Act - Get font name from the text shape
            string fontName = powerPointService.GetFontName(textBox);

            // Assert
            Assert.Equal("Arial", fontName);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetFontSizeFromTextShape()
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
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(
                powerPointService, System.Diagnostics.ProcessWindowStyle.Minimized);

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a text shape with specific font size
            dynamic shapes = slide.Shapes;
            dynamic textBox = shapes.AddTextbox(1, 100, 100, 200, 50); // 1 = msoTextOrientationHorizontal
            dynamic textFrame = textBox.TextFrame;
            dynamic textRange = textFrame.TextRange;
            textRange.Text = "Test Text";

            // Set font size to a known value
            dynamic font = textRange.Font;
            font.Size = 24;

            // Act - Get font size from the text shape
            float fontSize = powerPointService.GetFontSize(textBox);

            // Assert
            Assert.Equal(24f, fontSize);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetFontColorFromTextShape()
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
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(
                powerPointService, System.Diagnostics.ProcessWindowStyle.Minimized);

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a text shape with specific font color
            dynamic shapes = slide.Shapes;
            dynamic textBox = shapes.AddTextbox(1, 100, 100, 200, 50); // 1 = msoTextOrientationHorizontal
            dynamic textFrame = textBox.TextFrame;
            dynamic textRange = textFrame.TextRange;
            textRange.Text = "Test Text";

            // Set font color to red (RGB: 255, 0, 0)
            dynamic font = textRange.Font;
            font.Color.RGB = 255; // Red color in COM RGB format (0x0000FF)

            // Act - Get font color (RGB) from the text shape
            (int red, int green, int blue) color = powerPointService.GetFontColor(textBox);

            // Assert
            Assert.Equal(255, color.red);
            Assert.Equal(0, color.green);
            Assert.Equal(0, color.blue);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetShapeFillColor()
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
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(
                powerPointService, System.Diagnostics.ProcessWindowStyle.Minimized);

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a rectangle shape with a fill color
            dynamic shapes = slide.Shapes;
            // AddShape(Type, Left, Top, Width, Height)
            // 1 = msoShapeRectangle
            dynamic rectangle = shapes.AddShape(1, 100, 100, 200, 100);

            // Set fill color to blue (RGB: 0, 0, 255)
            dynamic fill = rectangle.Fill;
            fill.Visible = -1; // msoTrue
            fill.Solid();
            fill.ForeColor.RGB = 16711680; // Blue color in COM RGB format (0xFF0000)

            // Act - Get fill color from the shape
            (int red, int green, int blue) fillColor = powerPointService.GetShapeFillColor(rectangle);

            // Assert
            Assert.Equal(0, fillColor.red);
            Assert.Equal(0, fillColor.green);
            Assert.Equal(255, fillColor.blue);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetShapePosition()
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
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(
                powerPointService,
                System.Diagnostics.ProcessWindowStyle.Minimized);

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a rectangle shape at a specific position
            dynamic shapes = slide.Shapes;
            // AddShape(Type, Left, Top, Width, Height)
            // 1 = msoShapeRectangle
            dynamic rectangle = shapes.AddShape(1, 150, 200, 100, 50);

            // Act - Get position (Left, Top) from the shape
            (float left, float top) position = powerPointService.GetShapePosition(rectangle);

            // Assert
            Assert.Equal(150f, position.left);
            Assert.Equal(200f, position.top);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetShapeSize()
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
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(
                powerPointService,
                System.Diagnostics.ProcessWindowStyle.Minimized);

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a rectangle shape with specific size
            dynamic shapes = slide.Shapes;
            // AddShape(Type, Left, Top, Width, Height)
            // 1 = msoShapeRectangle
            dynamic rectangle = shapes.AddShape(1, 100, 100, 250, 150);

            // Act - Get size (Width, Height) from the shape
            (float width, float height) size = powerPointService.GetShapeSize(rectangle);

            // Assert
            Assert.Equal(250f, size.width);
            Assert.Equal(150f, size.height);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetTableCellContent()
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
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(
                powerPointService, System.Diagnostics.ProcessWindowStyle.Minimized);

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a table (2 rows, 3 columns) to the slide
            dynamic shapes = slide.Shapes;
            // AddTable(NumRows, NumColumns, Left, Top, Width, Height)
            dynamic table = shapes.AddTable(2, 3, 100, 100, 300, 100);

            // Set text in a specific cell (row 1, column 2)
            dynamic cell = table.Table.Cell(1, 2);
            dynamic cellShape = cell.Shape;
            dynamic textFrame = cellShape.TextFrame;
            dynamic textRange = textFrame.TextRange;
            textRange.Text = "Test Cell Content";

            // Act - Get table cell content
            string cellContent = powerPointService.GetTableCellContent(table, 1, 2);

            // Assert
            Assert.Equal("Test Cell Content", cellContent);
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldHandleShapesWithoutTextGracefully()
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
            (pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(
                powerPointService, System.Diagnostics.ProcessWindowStyle.Minimized);

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a shape without text (rectangle without text)
            dynamic shapes = slide.Shapes;
            // AddShape(Type, Left, Top, Width, Height)
            // 1 = msoShapeRectangle
            dynamic rectangle = shapes.AddShape(1, 100, 100, 200, 100);

            // Act - Try to get font name from a shape without text
            // This should not throw an exception but return null or empty string
            string? fontName = powerPointService.GetFontName(rectangle);

            // Assert - Should handle gracefully by returning null or empty
            Assert.True(string.IsNullOrEmpty(fontName),
                "GetFontName should return null or empty string for shapes without text");
        }
        finally
        {
            // Cleanup - close PowerPoint
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
        }
    }
}
