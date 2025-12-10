using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace PPTCoach.Core;

public class PowerPointService
{
    // Event for slide change detection
    public event EventHandler<int>? OnSlideChanged;

    // Event for shape selection change detection
    public event EventHandler<string>? OnShapeSelectionChanged;

    private object? _monitoredPresentation;
    private int _lastSlideNumber;
    private string? _lastSelectedShapeName;

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid pclsid);

    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [SupportedOSPlatform("windows")]
    public bool IsPowerPointInstalled()
    {
        try
        {
            Type? type = Type.GetTypeFromProgID("PowerPoint.Application");
            return type != null;
        }
        catch
        {
            return false;
        }
    }

    [SupportedOSPlatform("windows")]
    public object? GetRunningPowerPointInstance()
    {
        try
        {
            return ConnectToPowerPointOrThrow();
        }
        catch
        {
            // PowerPoint is not running or connection failed
            return null;
        }
    }

    [SupportedOSPlatform("windows")]
    public object ConnectToPowerPointOrThrow()
    {
        // Get CLSID from ProgID
        int clsidHr = CLSIDFromProgID("PowerPoint.Application", out Guid clsid);
        if (clsidHr != 0)
        {
            throw new COMException("Failed to get CLSID for PowerPoint.Application", clsidHr);
        }

        // Try to get active object from Running Object Table (ROT)
        int hr = GetActiveObject(ref clsid, IntPtr.Zero, out object obj);

        // S_OK = 0 means success
        const int S_OK = 0;

        if (hr != S_OK)
        {
            throw new COMException("Failed to connect to running PowerPoint instance", hr);
        }

        if (obj == null)
        {
            throw new COMException("PowerPoint instance is null");
        }

        return obj;
    }

    [SupportedOSPlatform("windows")]
    public object CreateNewPresentation(object powerPointInstance)
    {
        try
        {
            dynamic pptApp = powerPointInstance;
            dynamic presentations = pptApp.Presentations;
            dynamic presentation = presentations.Add();

            // Add a default blank slide
            dynamic slides = presentation.Slides;
            slides.Add(1, 12); // 12 = ppLayoutBlank

            return presentation;
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
        {
            throw new InvalidOperationException("Failed to create presentation. Invalid PowerPoint instance provided.", ex);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Failed to create presentation due to COM error.", ex);
        }
    }

    [SupportedOSPlatform("windows")]
    public int GetCurrentSlideNumber(object presentation)
    {
        dynamic pres = presentation;
        dynamic pptApp = pres.Application;
        dynamic activeWindow = pptApp.ActiveWindow;
        dynamic selection = activeWindow.Selection;

        // Check selection type first
        // ppSelectionNone = 0, ppSelectionSlides = 1
        int selectionType = selection.Type;

        // Get view type: 1 = Normal, 2 = Outline, 3 = SlideSorter, 7 = NotesPage, 9 = ReadingView
        int viewType = activeWindow.ViewType;

        // In Normal view, get the current slide being viewed
        if (viewType == 1)
        {
            try
            {
                dynamic view = activeWindow.View;
                dynamic slide = view.Slide;
                return slide.SlideIndex;
            }
            catch
            {
                // Fallback: if we can't get the slide from view, assume first slide
                return 1;
            }
        }

        // In Slide Sorter or other views, check if there's a slide selection
        if (selectionType == 0)
        {
            // No selection - but in Slide Sorter view, this means nothing is selected
            // In other views, we might still want to return the first slide if it exists
            if (viewType == 3) // Slide Sorter view
            {
                return 0;
            }

            // For other views, fallback to first slide if any exist
            dynamic slides = pres.Slides;
            return slides.Count > 0 ? 1 : 0;
        }

        // ppSelectionSlides = 1 means slides are selected
        if (selectionType == 1)
        {
            try
            {
                dynamic slideRange = selection.SlideRange;
                return slideRange.SlideNumber;
            }
            catch
            {
                // Couldn't get slide range
                return 0;
            }
        }

        // Other selection types
        return 0;
    }

    [SupportedOSPlatform("windows")]
    public int GetTotalSlideCount(object presentation)
    {
        dynamic pres = presentation;
        dynamic slides = pres.Slides;
        return slides.Count;
    }

    [SupportedOSPlatform("windows")]
    public string GetSlideTitle(object slide)
    {
        dynamic slideObj = slide;
        dynamic shapes = slideObj.Shapes;
        dynamic titleShape = shapes[1];
        dynamic textFrame = titleShape.TextFrame;
        dynamic textRange = textFrame.TextRange;
        return textRange.Text;
    }

    [SupportedOSPlatform("windows")]
    public string GetAllTextFromSlide(object slide)
    {
        dynamic slideObj = slide;
        dynamic shapes = slideObj.Shapes;
        var allText = new System.Text.StringBuilder();

        for (int i = 1; i <= shapes.Count; i++)
        {
            try
            {
                dynamic shape = shapes[i];
                if (shape.HasTextFrame == -1) // -1 = msoTrue in Office constants
                {
                    dynamic textFrame = shape.TextFrame;
                    if (textFrame.HasText == -1)
                    {
                        dynamic textRange = textFrame.TextRange;
                        string text = textRange.Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            allText.AppendLine(text.Trim());
                        }
                    }
                }
            }
            catch
            {
                // Skip shapes without text or that can't be accessed
                continue;
            }
        }

        return allText.ToString().Trim();
    }

    [SupportedOSPlatform("windows")]
    public List<object> GetShapesOnSlide(object slide)
    {
        dynamic slideObj = slide;
        dynamic shapes = slideObj.Shapes;
        var shapesList = new List<object>();

        for (int i = 1; i <= shapes.Count; i++)
        {
            shapesList.Add(shapes[i]);
        }

        return shapesList;
    }

    [SupportedOSPlatform("windows")]
    public string GetShapeType(object shape)
    {
        dynamic shapeObj = shape;
        int typeValue = shapeObj.Type;

        // Map PowerPoint shape type constants to readable names
        // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoautoshapetype
        return typeValue switch
        {
            1 => "AutoShape",
            13 => "Picture",
            14 => "Placeholder",
            17 => "TextBox",
            19 => "Table",
            _ => "Unknown"
        };
    }

    [SupportedOSPlatform("windows")]
    public object AddBlankSlide(object presentation)
    {
        return AddSlideWithLayout(presentation, 12); // 12 = ppLayoutBlank
    }

    [SupportedOSPlatform("windows")]
    public object AddSlideWithLayout(object presentation, int layoutType)
    {
        dynamic pres = presentation;
        dynamic slides = pres.Slides;
        int newIndex = slides.Count + 1;
        dynamic newSlide = slides.Add(newIndex, layoutType);
        return newSlide;
    }

    [SupportedOSPlatform("windows")]
    public void DeleteSlideByIndex(object presentation, int index)
    {
        dynamic pres = presentation;
        dynamic slides = pres.Slides;

        if (slides.Count == 1)
        {
            throw new InvalidOperationException("Cannot delete the last slide in the presentation.");
        }

        slides[index].Delete();
    }

    [SupportedOSPlatform("windows")]
    public void MoveSlide(object presentation, int fromIndex, int toIndex)
    {
        dynamic pres = presentation;
        dynamic slides = pres.Slides;
        dynamic slide = slides[fromIndex];
        slide.MoveTo(toIndex);
    }

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
        return ConvertComRgbToTuple(rgb);
    }

    [SupportedOSPlatform("windows")]
    public (int red, int green, int blue) GetShapeFillColor(object shape)
    {
        dynamic shapeObj = shape;
        dynamic fill = shapeObj.Fill;
        int rgb = fill.ForeColor.RGB;
        return ConvertComRgbToTuple(rgb);
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

    [SupportedOSPlatform("windows")]
    private dynamic? GetFontFromShape(object shape)
    {
        try
        {
            dynamic shapeObj = shape;

            // Check if shape has a text frame
            if (shapeObj.HasTextFrame == 0) // 0 = msoFalse
            {
                return null;
            }

            dynamic textFrame = shapeObj.TextFrame;

            // Check if text frame has text
            if (textFrame.HasText == 0) // 0 = msoFalse
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

    private static (int red, int green, int blue) ConvertComRgbToTuple(int rgb)
    {
        // COM RGB format is BGR (little-endian): 0xBBGGRR
        int red = rgb & 0xFF;
        int green = (rgb >> 8) & 0xFF;
        int blue = (rgb >> 16) & 0xFF;
        return (red, green, blue);
    }

    [SupportedOSPlatform("windows")]
    public void StartMonitoringSlideChanges(object presentation)
    {
        _monitoredPresentation = presentation;
        _lastSlideNumber = GetCurrentSlideNumber(presentation);
    }

    [SupportedOSPlatform("windows")]
    public void StopMonitoringSlideChanges()
    {
        _monitoredPresentation = null;
    }

    [SupportedOSPlatform("windows")]
    public void CheckForSlideChange()
    {
        if (_monitoredPresentation == null) return;

        try
        {
            int currentSlideNumber = GetCurrentSlideNumber(_monitoredPresentation);
            if (currentSlideNumber != _lastSlideNumber && currentSlideNumber > 0)
            {
                _lastSlideNumber = currentSlideNumber;
                OnSlideChanged?.Invoke(this, currentSlideNumber);
            }
        }
        catch
        {
            // Ignore errors during monitoring (presentation might be closed)
        }
    }

    /// <summary>
    /// Starts monitoring for shape selection changes
    /// </summary>
    [SupportedOSPlatform("windows")]
    public void StartMonitoringShapeSelection(object presentation)
    {
        _monitoredPresentation = presentation;
        _lastSelectedShapeName = GetCurrentSelectedShapeName(presentation);
    }

    /// <summary>
    /// Stops monitoring for shape selection changes
    /// </summary>
    [SupportedOSPlatform("windows")]
    public void StopMonitoringShapeSelection()
    {
        _lastSelectedShapeName = null;
    }

    /// <summary>
    /// Checks if the selected shape has changed and fires event if it has
    /// This should be called periodically to detect shape selection changes
    /// </summary>
    [SupportedOSPlatform("windows")]
    public void CheckForShapeSelectionChange()
    {
        if (_monitoredPresentation == null) return;

        try
        {
            string? currentShapeName = GetCurrentSelectedShapeName(_monitoredPresentation);
            if (currentShapeName != _lastSelectedShapeName && currentShapeName != null)
            {
                _lastSelectedShapeName = currentShapeName;
                OnShapeSelectionChanged?.Invoke(this, currentShapeName);
            }
        }
        catch
        {
            // Ignore errors during monitoring (presentation might be closed)
        }
    }

    /// <summary>
    /// Gets the name of the currently selected shape
    /// </summary>
    [SupportedOSPlatform("windows")]
    private string? GetCurrentSelectedShapeName(object presentation)
    {
        try
        {
            dynamic pres = presentation;
            dynamic app = pres.Application;
            dynamic activeWindow = app.ActiveWindow;
            dynamic selection = activeWindow.Selection;

            // Check if a shape is selected (Type 2 = ppSelectionShapes)
            if (selection.Type == 2)
            {
                dynamic shapeRange = selection.ShapeRange;
                if (shapeRange.Count > 0)
                {
                    return shapeRange[1].Name;
                }
            }

            return null;
        }
        catch
        {
            return null;
        }
    }
}
