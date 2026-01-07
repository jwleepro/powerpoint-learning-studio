namespace PPTCoach.Core.Constants;

/// <summary>
/// PowerPoint slide layout types (ppSlideLayout enumeration)
/// </summary>
public static class PpSlideLayout
{
    public const int Blank = 12;
    public const int Title = 1;
    public const int Text = 2;
}

/// <summary>
/// PowerPoint view types (ppViewType enumeration)
/// </summary>
public static class PpViewType
{
    public const int Normal = 1;
    public const int SlideSorter = 3;
}

/// <summary>
/// PowerPoint selection types (ppSelectionType enumeration)
/// </summary>
public static class PpSelectionType
{
    public const int None = 0;
    public const int Slides = 1;
    public const int Shapes = 2;
}

/// <summary>
/// Microsoft Office tri-state values (msoTriState enumeration)
/// </summary>
public static class MsoTriState
{
    public const int True = -1;
    public const int False = 0;
}

/// <summary>
/// Microsoft Office shape types (msoShapeType enumeration)
/// </summary>
public static class MsoShapeType
{
    public const int AutoShape = 1;
    public const int Picture = 13;
    public const int Placeholder = 14;
    public const int TextBox = 17;
    public const int Table = 19;
}

/// <summary>
/// PowerPoint COM identifiers
/// </summary>
public static class PowerPointProgId
{
    public const string Application = "PowerPoint.Application";
    public const string Executable = "powerpnt.exe";
}
