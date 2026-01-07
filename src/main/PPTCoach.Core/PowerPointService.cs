using System.Runtime.Versioning;
using PPTCoach.Core.Interfaces;
using PPTCoach.Core.Services;

namespace PPTCoach.Core;

/// <summary>
/// Facade service that provides a unified interface to all PowerPoint operations.
/// Maintains backward compatibility while delegating to specialized services.
/// </summary>
public class PowerPointService : IPowerPointService
{
    private readonly IPowerPointConnection _connection;
    private readonly IPresentationManager _presentationManager;
    private readonly ISlideOperations _slideOperations;
    private readonly ISlideQuery _slideQuery;
    private readonly IShapeQuery _shapeQuery;
    private readonly IFontQuery _fontQuery;
    private readonly IPresentationMonitor _monitor;

    // Events from IPresentationMonitor
    public event EventHandler<int>? OnSlideChanged
    {
        add => _monitor.OnSlideChanged += value;
        remove => _monitor.OnSlideChanged -= value;
    }

    public event EventHandler<string>? OnShapeSelectionChanged
    {
        add => _monitor.OnShapeSelectionChanged += value;
        remove => _monitor.OnShapeSelectionChanged -= value;
    }

    public event EventHandler<string>? OnPresentationSaved
    {
        add => _monitor.OnPresentationSaved += value;
        remove => _monitor.OnPresentationSaved -= value;
    }

    /// <summary>
    /// Default constructor - creates instances of all service implementations.
    /// This maintains backward compatibility with existing code.
    /// </summary>
    public PowerPointService()
    {
        _connection = new PowerPointConnectionService();
        _presentationManager = new PresentationManagerService();
        _slideOperations = new SlideOperationsService();
        _slideQuery = new SlideQueryService();
        _shapeQuery = new ShapeQueryService();
        _fontQuery = new FontQueryService();
        _monitor = new PresentationMonitorService(_slideQuery);
    }

    /// <summary>
    /// Constructor with dependency injection support.
    /// This allows for testing and custom implementations.
    /// </summary>
    public PowerPointService(
        IPowerPointConnection connection,
        IPresentationManager presentationManager,
        ISlideOperations slideOperations,
        ISlideQuery slideQuery,
        IShapeQuery shapeQuery,
        IFontQuery fontQuery,
        IPresentationMonitor monitor)
    {
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _presentationManager = presentationManager ?? throw new ArgumentNullException(nameof(presentationManager));
        _slideOperations = slideOperations ?? throw new ArgumentNullException(nameof(slideOperations));
        _slideQuery = slideQuery ?? throw new ArgumentNullException(nameof(slideQuery));
        _shapeQuery = shapeQuery ?? throw new ArgumentNullException(nameof(shapeQuery));
        _fontQuery = fontQuery ?? throw new ArgumentNullException(nameof(fontQuery));
        _monitor = monitor ?? throw new ArgumentNullException(nameof(monitor));
    }

    // IPowerPointConnection methods
    [SupportedOSPlatform("windows")]
    public bool IsPowerPointInstalled() => _connection.IsPowerPointInstalled();

    [SupportedOSPlatform("windows")]
    public object? GetRunningPowerPointInstance() => _connection.GetRunningPowerPointInstance();

    [SupportedOSPlatform("windows")]
    public object ConnectToPowerPointOrThrow() => _connection.ConnectToPowerPointOrThrow();

    // IPresentationManager methods
    [SupportedOSPlatform("windows")]
    public object CreateNewPresentation(object powerPointInstance) =>
        _presentationManager.CreateNewPresentation(powerPointInstance);

    // ISlideOperations methods
    [SupportedOSPlatform("windows")]
    public object AddBlankSlide(object presentation) =>
        _slideOperations.AddBlankSlide(presentation);

    [SupportedOSPlatform("windows")]
    public object AddSlideWithLayout(object presentation, int layoutType) =>
        _slideOperations.AddSlideWithLayout(presentation, layoutType);

    [SupportedOSPlatform("windows")]
    public void DeleteSlideByIndex(object presentation, int index) =>
        _slideOperations.DeleteSlideByIndex(presentation, index);

    [SupportedOSPlatform("windows")]
    public void MoveSlide(object presentation, int fromIndex, int toIndex) =>
        _slideOperations.MoveSlide(presentation, fromIndex, toIndex);

    // ISlideQuery methods
    [SupportedOSPlatform("windows")]
    public int GetCurrentSlideNumber(object presentation) =>
        _slideQuery.GetCurrentSlideNumber(presentation);

    [SupportedOSPlatform("windows")]
    public int GetTotalSlideCount(object presentation) =>
        _slideQuery.GetTotalSlideCount(presentation);

    [SupportedOSPlatform("windows")]
    public string GetSlideTitle(object slide) =>
        _slideQuery.GetSlideTitle(slide);

    [SupportedOSPlatform("windows")]
    public string GetAllTextFromSlide(object slide) =>
        _slideQuery.GetAllTextFromSlide(slide);

    [SupportedOSPlatform("windows")]
    public List<object> GetShapesOnSlide(object slide) =>
        _slideQuery.GetShapesOnSlide(slide);

    // IShapeQuery methods
    [SupportedOSPlatform("windows")]
    public string GetShapeType(object shape) =>
        _shapeQuery.GetShapeType(shape);

    [SupportedOSPlatform("windows")]
    public (int red, int green, int blue) GetShapeFillColor(object shape) =>
        _shapeQuery.GetShapeFillColor(shape);

    [SupportedOSPlatform("windows")]
    public (float left, float top) GetShapePosition(object shape) =>
        _shapeQuery.GetShapePosition(shape);

    [SupportedOSPlatform("windows")]
    public (float width, float height) GetShapeSize(object shape) =>
        _shapeQuery.GetShapeSize(shape);

    [SupportedOSPlatform("windows")]
    public string GetTableCellContent(object tableShape, int row, int column) =>
        _shapeQuery.GetTableCellContent(tableShape, row, column);

    // IFontQuery methods
    [SupportedOSPlatform("windows")]
    public string? GetFontName(object shape) =>
        _fontQuery.GetFontName(shape);

    [SupportedOSPlatform("windows")]
    public float GetFontSize(object shape) =>
        _fontQuery.GetFontSize(shape);

    [SupportedOSPlatform("windows")]
    public (int red, int green, int blue) GetFontColor(object shape) =>
        _fontQuery.GetFontColor(shape);

    // IPresentationMonitor methods
    [SupportedOSPlatform("windows")]
    public void StartMonitoringSlideChanges(object presentation) =>
        _monitor.StartMonitoringSlideChanges(presentation);

    [SupportedOSPlatform("windows")]
    public void StopMonitoringSlideChanges() =>
        _monitor.StopMonitoringSlideChanges();

    [SupportedOSPlatform("windows")]
    public void CheckForSlideChange() =>
        _monitor.CheckForSlideChange();

    [SupportedOSPlatform("windows")]
    public void StartMonitoringShapeSelection(object presentation) =>
        _monitor.StartMonitoringShapeSelection(presentation);

    [SupportedOSPlatform("windows")]
    public void StopMonitoringShapeSelection() =>
        _monitor.StopMonitoringShapeSelection();

    [SupportedOSPlatform("windows")]
    public void CheckForShapeSelectionChange() =>
        _monitor.CheckForShapeSelectionChange();

    [SupportedOSPlatform("windows")]
    public void StartMonitoringPresentationSave(object presentation) =>
        _monitor.StartMonitoringPresentationSave(presentation);

    [SupportedOSPlatform("windows")]
    public void StopMonitoringPresentationSave() =>
        _monitor.StopMonitoringPresentationSave();

    [SupportedOSPlatform("windows")]
    public void CheckForPresentationSave() =>
        _monitor.CheckForPresentationSave();
}
