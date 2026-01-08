using PPTCoach.Core;
using PPTCoach.Tests.Utils;
using PPTCoach.Tests.Constants;
using System.Runtime.Versioning;

namespace PPTCoach.Tests.Phase01;

/// <summary>
/// Tests for Phase 1.5: Element Property Reading
/// </summary>
public class PowerPointElementPropertyTests : PowerPointTestBase
{
    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetFontNameFromTextShape()
    {
        // Arrange
        SetupPowerPointWithPresentation();
        var shapes = GetFirstSlideShapes();

        // Add a text box with specific font
        var textBox = PowerPointTestHelpers.AddTestTextBox(shapes);
        dynamic font = textBox.TextFrame.TextRange.Font;
        font.Name = "Arial";

        // Act - Get font name from the text shape
        string fontName = PowerPointService.GetFontName(textBox);

        // Assert
        Assert.Equal("Arial", fontName);
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetFontSizeFromTextShape()
    {
        // Arrange
        SetupPowerPointWithPresentation();
        var shapes = GetFirstSlideShapes();

        // Add a text box with specific font size
        var textBox = PowerPointTestHelpers.AddTestTextBox(shapes);
        dynamic font = textBox.TextFrame.TextRange.Font;
        font.Size = 24;

        // Act - Get font size from the text shape
        float fontSize = PowerPointService.GetFontSize(textBox);

        // Assert
        Assert.Equal(24f, fontSize);
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetFontColorFromTextShape()
    {
        // Arrange
        SetupPowerPointWithPresentation();
        var shapes = GetFirstSlideShapes();

        // Add a text box with specific font color
        var textBox = PowerPointTestHelpers.AddTestTextBox(shapes);
        dynamic font = textBox.TextFrame.TextRange.Font;
        font.Color.RGB = TestColors.Red;

        // Act - Get font color (RGB) from the text shape
        (int red, int green, int blue) color = PowerPointService.GetFontColor(textBox);

        // Assert
        Assert.Equal(255, color.red);
        Assert.Equal(0, color.green);
        Assert.Equal(0, color.blue);
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetShapeFillColor()
    {
        // Arrange
        SetupPowerPointWithPresentation();
        var shapes = GetFirstSlideShapes();

        // Add a rectangle with a fill color
        var rectangle = PowerPointTestHelpers.AddTestRectangle(shapes);

        // Set fill color to blue
        dynamic fill = rectangle.Fill;
        fill.Visible = -1; // msoTrue
        fill.Solid();
        fill.ForeColor.RGB = TestColors.Blue;

        // Act - Get fill color from the shape
        (int red, int green, int blue) fillColor = PowerPointService.GetShapeFillColor(rectangle);

        // Assert
        Assert.Equal(0, fillColor.red);
        Assert.Equal(0, fillColor.green);
        Assert.Equal(255, fillColor.blue);
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetShapePosition()
    {
        // Arrange
        SetupPowerPointWithPresentation();
        var shapes = GetFirstSlideShapes();

        // Add a rectangle at a specific position
        var rectangle = PowerPointTestHelpers.AddTestRectangle(shapes, left: 150f, top: 200f);

        // Act - Get position (Left, Top) from the shape
        (float left, float top) position = PowerPointService.GetShapePosition(rectangle);

        // Assert
        Assert.Equal(150f, position.left);
        Assert.Equal(200f, position.top);
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetShapeSize()
    {
        // Arrange
        SetupPowerPointWithPresentation();
        var shapes = GetFirstSlideShapes();

        // Add a rectangle with specific size
        var rectangle = PowerPointTestHelpers.AddTestRectangle(shapes, width: 250f, height: 150f);

        // Act - Get size (Width, Height) from the shape
        (float width, float height) size = PowerPointService.GetShapeSize(rectangle);

        // Assert
        Assert.Equal(250f, size.width);
        Assert.Equal(150f, size.height);
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetTableCellContent()
    {
        // Arrange
        SetupPowerPointWithPresentation();
        var shapes = GetFirstSlideShapes();

        // Add a table (2 rows, 3 columns) to the slide
        var table = PowerPointTestHelpers.AddTestTable(shapes, numRows: 2, numColumns: 3);

        // Set text in a specific cell (row 1, column 2)
        dynamic cell = table.Table.Cell(ComIndexing.FirstIndex, 2);
        dynamic cellShape = cell.Shape;
        dynamic textFrame = cellShape.TextFrame;
        dynamic textRange = textFrame.TextRange;
        textRange.Text = "Test Cell Content";

        // Act - Get table cell content
        string cellContent = PowerPointService.GetTableCellContent(table, ComIndexing.FirstIndex, 2);

        // Assert
        Assert.Equal("Test Cell Content", cellContent);
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldHandleShapesWithoutTextGracefully()
    {
        // Arrange
        SetupPowerPointWithPresentation();
        var shapes = GetFirstSlideShapes();

        // Add a shape without text (rectangle without text)
        var rectangle = PowerPointTestHelpers.AddTestRectangle(shapes);

        // Act - Try to get font name from a shape without text
        // This should not throw an exception but return null or empty string
        string? fontName = PowerPointService.GetFontName(rectangle);

        // Assert - Should handle gracefully by returning null or empty
        Assert.True(string.IsNullOrEmpty(fontName),
            "GetFontName should return null or empty string for shapes without text");
    }
}
