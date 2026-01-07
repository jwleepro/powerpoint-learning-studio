using System.Runtime.Versioning;
using PPTCoach.Core.Constants;
using PPTCoach.Core.Interfaces;

namespace PPTCoach.Core.Services;

/// <summary>
/// Service for slide manipulation operations
/// </summary>
public class SlideOperationsService : ISlideOperations
{
    [SupportedOSPlatform("windows")]
    public object AddBlankSlide(object presentation)
    {
        return AddSlideWithLayout(presentation, PpSlideLayout.Blank);
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
}
