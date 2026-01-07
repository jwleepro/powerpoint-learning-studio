using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using PPTCoach.Core.Constants;
using PPTCoach.Core.Interfaces;

namespace PPTCoach.Core.Services;

/// <summary>
/// Service for creating and managing PowerPoint presentations
/// </summary>
public class PresentationManagerService : IPresentationManager
{
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
            slides.Add(1, PpSlideLayout.Blank);

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
}
