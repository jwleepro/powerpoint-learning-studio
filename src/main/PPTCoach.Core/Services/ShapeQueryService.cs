using System.Runtime.Versioning;
using PPTCoach.Core.Constants;
using PPTCoach.Core.Interfaces;
using PPTCoach.Core.Utilities;

namespace PPTCoach.Core.Services;

/// <summary>
/// Service for querying shape properties
/// </summary>
public class ShapeQueryService : IShapeQuery
{
    [SupportedOSPlatform("windows")]
    public string GetShapeType(object shape)
    {
        dynamic shapeObj = shape;
        int typeValue = shapeObj.Type;

        // Map PowerPoint shape type constants to readable names
        // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoautoshapetype
        return typeValue switch
        {
            MsoShapeType.AutoShape => "AutoShape",
            MsoShapeType.Picture => "Picture",
            MsoShapeType.Placeholder => "Placeholder",
            MsoShapeType.TextBox => "TextBox",
            MsoShapeType.Table => "Table",
            _ => "Unknown"
        };
    }

    [SupportedOSPlatform("windows")]
    public (int red, int green, int blue) GetShapeFillColor(object shape)
    {
        dynamic shapeObj = shape;
        dynamic fill = shapeObj.Fill;
        int rgb = fill.ForeColor.RGB;
        return ComRgbConverter.ConvertToRgbTuple(rgb);
    }

    [SupportedOSPlatform("windows")]
    public (float left, float top) GetShapePosition(object shape)
    {
        dynamic shapeObj = shape;
        float left = shapeObj.Left;
        float top = shapeObj.Top;
        return (left, top);
    }

    [SupportedOSPlatform("windows")]
    public (float width, float height) GetShapeSize(object shape)
    {
        dynamic shapeObj = shape;
        float width = shapeObj.Width;
        float height = shapeObj.Height;
        return (width, height);
    }

    [SupportedOSPlatform("windows")]
    public string GetTableCellContent(object tableShape, int row, int column)
    {
        dynamic table = tableShape;
        dynamic cell = table.Table.Cell(row, column);
        dynamic cellShape = cell.Shape;
        dynamic textFrame = cellShape.TextFrame;
        dynamic textRange = textFrame.TextRange;
        return textRange.Text;
    }
}
