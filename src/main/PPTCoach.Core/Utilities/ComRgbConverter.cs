namespace PPTCoach.Core.Utilities;

/// <summary>
/// Utility class for converting COM RGB values to RGB tuples
/// </summary>
public static class ComRgbConverter
{
    /// <summary>
    /// Converts COM RGB format (BGR little-endian: 0xBBGGRR) to RGB tuple
    /// </summary>
    /// <param name="rgb">COM RGB value</param>
    /// <returns>RGB tuple (red, green, blue)</returns>
    public static (int red, int green, int blue) ConvertToRgbTuple(int rgb)
    {
        int red = rgb & 0xFF;
        int green = (rgb >> 8) & 0xFF;
        int blue = (rgb >> 16) & 0xFF;
        return (red, green, blue);
    }
}
