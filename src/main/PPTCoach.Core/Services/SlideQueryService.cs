using System.Runtime.Versioning;
using PPTCoach.Core.Constants;
using PPTCoach.Core.Interfaces;

namespace PPTCoach.Core.Services;

/// <summary>
/// Service for querying slide information
/// </summary>
public class SlideQueryService : ISlideQuery
{
    [SupportedOSPlatform("windows")]
    public int GetCurrentSlideNumber(object presentation)
    {
        dynamic pres = presentation;
        dynamic pptApp = pres.Application;
        dynamic activeWindow = pptApp.ActiveWindow;
        dynamic selection = activeWindow.Selection;

        int selectionType = selection.Type;
        int viewType = activeWindow.ViewType;

        // In Normal view, get the current slide being viewed
        if (viewType == PpViewType.Normal)
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
        if (selectionType == PpSelectionType.None)
        {
            // No selection - but in Slide Sorter view, this means nothing is selected
            // In other views, we might still want to return the first slide if it exists
            if (viewType == PpViewType.SlideSorter)
            {
                return 0;
            }

            // For other views, fallback to first slide if any exist
            dynamic slides = pres.Slides;
            return slides.Count > 0 ? 1 : 0;
        }

        // ppSelectionSlides means slides are selected
        if (selectionType == PpSelectionType.Slides)
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
                if (shape.HasTextFrame == MsoTriState.True)
                {
                    dynamic textFrame = shape.TextFrame;
                    if (textFrame.HasText == MsoTriState.True)
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
}
