using System.Runtime.Versioning;
using PPTCoach.Core.Constants;
using PPTCoach.Core.Interfaces;

namespace PPTCoach.Core.Services;

/// <summary>
/// Service for monitoring presentation events
/// </summary>
public class PresentationMonitorService : IPresentationMonitor
{
    private readonly ISlideQuery _slideQuery;
    private object? _monitoredPresentation;
    private int _lastSlideNumber;
    private string? _lastSelectedShapeName;
    private DateTime? _lastSaveTime;

    public event EventHandler<int>? OnSlideChanged;
    public event EventHandler<string>? OnShapeSelectionChanged;
    public event EventHandler<string>? OnPresentationSaved;

    public PresentationMonitorService(ISlideQuery slideQuery)
    {
        _slideQuery = slideQuery ?? throw new ArgumentNullException(nameof(slideQuery));
    }

    [SupportedOSPlatform("windows")]
    public void StartMonitoringSlideChanges(object presentation)
    {
        _monitoredPresentation = presentation;
        _lastSlideNumber = _slideQuery.GetCurrentSlideNumber(presentation);
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
            int currentSlideNumber = _slideQuery.GetCurrentSlideNumber(_monitoredPresentation);
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

    [SupportedOSPlatform("windows")]
    public void StartMonitoringShapeSelection(object presentation)
    {
        _monitoredPresentation = presentation;
        _lastSelectedShapeName = GetCurrentSelectedShapeName(presentation);
    }

    [SupportedOSPlatform("windows")]
    public void StopMonitoringShapeSelection()
    {
        _lastSelectedShapeName = null;
    }

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

    [SupportedOSPlatform("windows")]
    public void StartMonitoringPresentationSave(object presentation)
    {
        _monitoredPresentation = presentation;
        _lastSaveTime = GetPresentationLastSaveTime(presentation);
    }

    [SupportedOSPlatform("windows")]
    public void StopMonitoringPresentationSave()
    {
        _lastSaveTime = null;
    }

    [SupportedOSPlatform("windows")]
    public void CheckForPresentationSave()
    {
        if (_monitoredPresentation == null) return;

        try
        {
            DateTime? currentSaveTime = GetPresentationLastSaveTime(_monitoredPresentation);
            if (currentSaveTime.HasValue && currentSaveTime != _lastSaveTime)
            {
                _lastSaveTime = currentSaveTime;
                string? filePath = GetPresentationFilePath(_monitoredPresentation);
                if (!string.IsNullOrEmpty(filePath))
                {
                    OnPresentationSaved?.Invoke(this, filePath);
                }
            }
        }
        catch
        {
            // Ignore errors during monitoring (presentation might be closed)
        }
    }

    [SupportedOSPlatform("windows")]
    private string? GetCurrentSelectedShapeName(object presentation)
    {
        try
        {
            dynamic pres = presentation;
            dynamic app = pres.Application;
            dynamic activeWindow = app.ActiveWindow;
            dynamic selection = activeWindow.Selection;

            // Check if a shape is selected
            if (selection.Type == PpSelectionType.Shapes)
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

    [SupportedOSPlatform("windows")]
    private DateTime? GetPresentationLastSaveTime(object presentation)
    {
        try
        {
            dynamic pres = presentation;
            // Check if presentation has been saved (has a path)
            string path = pres.Path;
            if (string.IsNullOrEmpty(path))
            {
                return null;
            }

            string fullPath = pres.FullName;
            if (File.Exists(fullPath))
            {
                return File.GetLastWriteTime(fullPath);
            }

            return null;
        }
        catch
        {
            return null;
        }
    }

    [SupportedOSPlatform("windows")]
    private string? GetPresentationFilePath(object presentation)
    {
        try
        {
            dynamic pres = presentation;
            return pres.FullName;
        }
        catch
        {
            return null;
        }
    }
}
