namespace PPTCoach.Core.Interfaces;

/// <summary>
/// Facade interface that combines all PowerPoint service capabilities
/// </summary>
public interface IPowerPointService :
    IPowerPointConnection,
    IPresentationManager,
    ISlideOperations,
    ISlideQuery,
    IShapeQuery,
    IFontQuery,
    IPresentationMonitor
{
}
