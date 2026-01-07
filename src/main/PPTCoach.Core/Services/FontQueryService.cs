using System.Runtime.Versioning;
using PPTCoach.Core.Constants;
using PPTCoach.Core.Interfaces;
using PPTCoach.Core.Utilities;

namespace PPTCoach.Core.Services;

/// <summary>
/// Service for querying font properties
/// </summary>
public class FontQueryService : IFontQuery
{
    [SupportedOSPlatform("windows")]
    public string? GetFontName(object shape)
    {
        try
        {
            dynamic? font = GetFontFromShape(shape);
            if (font == null)
            {
                return null;
            }
            return font.Name;
        }
        catch
        {
            return null;
        }
    }

    [SupportedOSPlatform("windows")]
    public float GetFontSize(object shape)
    {
        dynamic? font = GetFontFromShape(shape);
        if (font == null)
        {
            return 0f;
        }
        return font.Size;
    }

    [SupportedOSPlatform("windows")]
    public (int red, int green, int blue) GetFontColor(object shape)
    {
        dynamic? font = GetFontFromShape(shape);
        if (font == null)
        {
            return (0, 0, 0);
        }
        int rgb = font.Color.RGB;
        return ComRgbConverter.ConvertToRgbTuple(rgb);
    }

    [SupportedOSPlatform("windows")]
    private dynamic? GetFontFromShape(object shape)
    {
        try
        {
            dynamic shapeObj = shape;

            // Check if shape has a text frame
            if (shapeObj.HasTextFrame == MsoTriState.False)
            {
                return null;
            }

            dynamic textFrame = shapeObj.TextFrame;

            // Check if text frame has text
            if (textFrame.HasText == MsoTriState.False)
            {
                return null;
            }

            dynamic textRange = textFrame.TextRange;
            return textRange.Font;
        }
        catch
        {
            return null;
        }
    }
}
