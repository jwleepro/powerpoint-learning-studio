using PPTCoach.Core;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.CSharp.RuntimeBinder;

namespace PPTCoach.Tests;

public class PowerPointConnectionTests
{
    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldDetectIfPowerPointIsInstalled()
    {
        // Arrange
        var powerPointService = new PowerPointService();

        // Act
        bool isInstalled = powerPointService.IsPowerPointInstalled();

        // Assert
        Assert.True(isInstalled, "PowerPoint should be detected as installed on this system");
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReturnNullWhenPowerPointIsNotRunning()
    {
        // Arrange
        var powerPointService = new PowerPointService();

        // Act
        var instance = powerPointService.GetRunningPowerPointInstance();

        // Assert
        Assert.Null(instance);
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldConnectToRunningPowerPointInstance()
    {
        // Arrange
        var powerPointService = new PowerPointService();

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        System.Diagnostics.Process? pptProcess = null;

        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint by launching the process
            // This simulates a user launching PowerPoint manually
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            // Assert
            Assert.NotNull(instance);

            // Verify it's a valid PowerPoint application object
            dynamic pptApp = instance;
            string name = pptApp.Name;
            Assert.Contains("PowerPoint", name);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldHandleMultiplePowerPointInstances()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess1 = null;
        System.Diagnostics.Process? pptProcess2 = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Act - Start first PowerPoint instance
            pptProcess1 = new System.Diagnostics.Process();
            pptProcess1.StartInfo.FileName = "powerpnt.exe";
            pptProcess1.StartInfo.UseShellExecute = true;
            pptProcess1.Start();

            // Wait for first PowerPoint instance with retry logic
            object? instance1 = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance1 = powerPointService.GetRunningPowerPointInstance();
                if (instance1 != null)
                {
                    break;
                }
            }

            // Start second PowerPoint instance
            pptProcess2 = new System.Diagnostics.Process();
            pptProcess2.StartInfo.FileName = "powerpnt.exe";
            pptProcess2.StartInfo.UseShellExecute = true;
            pptProcess2.Start();

            // Wait for second instance
            System.Threading.Thread.Sleep(5000);

            // Get instance again (should still work with multiple instances running)
            var instance2 = powerPointService.GetRunningPowerPointInstance();

            // Assert
            Assert.NotNull(instance1);
            Assert.NotNull(instance2);

            // Both should be valid PowerPoint application objects
            dynamic pptApp1 = instance1;
            dynamic pptApp2 = instance2;

            string name1 = pptApp1.Name;
            string name2 = pptApp2.Name;

            Assert.Contains("PowerPoint", name1);
            Assert.Contains("PowerPoint", name2);
        }
        finally
        {
            // Cleanup - close all PowerPoint instances
            if (pptProcess1 != null && !pptProcess1.HasExited)
            {
                pptProcess1.Kill();
                pptProcess1.Dispose();
            }

            if (pptProcess2 != null && !pptProcess2.HasExited)
            {
                pptProcess2.Kill();
                pptProcess2.Dispose();
            }

            // Give time for cleanup
            System.Threading.Thread.Sleep(2000);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldThrowExceptionWhenCOMConnectionFails()
    {
        // Arrange
        var powerPointService = new PowerPointService();

        // Act & Assert
        // Attempting to connect with an invalid ProgID should throw COMException
        var exception = Assert.Throws<COMException>(() =>
            powerPointService.ConnectToPowerPointOrThrow());

        // Verify the exception contains meaningful information
        Assert.NotNull(exception);
        Assert.NotEmpty(exception.Message);
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldCreateNewPowerPointPresentation()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Act
            var presentation = powerPointService.CreateNewPresentation(instance);

            // Assert
            Assert.NotNull(presentation);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReturnPresentationObjectAfterCreation()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Act
            var presentation = powerPointService.CreateNewPresentation(instance);

            // Assert
            Assert.NotNull(presentation);

            // Verify it's a valid presentation object by checking its properties
            dynamic pres = presentation;

            // Should have Slides collection
            dynamic slides = pres.Slides;
            Assert.NotNull(slides);

            // Should have a Name property
            string name = pres.Name;
            Assert.NotNull(name);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldCreatePresentationWithDefaultBlankSlide()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Act
            var presentation = powerPointService.CreateNewPresentation(instance);

            // Assert
            Assert.NotNull(presentation);

            // Verify it has a default blank slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            int slideCount = slides.Count;

            Assert.Equal(1, slideCount);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldHandleCreationFailureGracefully()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        var invalidInstance = new object(); // Invalid PowerPoint instance

        // Act & Assert
        var exception = Assert.Throws<InvalidOperationException>(() =>
            powerPointService.CreateNewPresentation(invalidInstance));

        // Verify the exception contains meaningful information about the failure
        Assert.NotNull(exception);
        Assert.Contains("presentation", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetCurrentSlideNumber()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Act
            int currentSlideNumber = powerPointService.GetCurrentSlideNumber(presentation);

            // Assert
            Assert.Equal(1, currentSlideNumber);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReturn0WhenNoSlideIsSelected()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation (without selecting any slide in slide sorter view)
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Deselect any slide by going to slide sorter view without selection
            dynamic pptApp = instance;
            dynamic activeWindow = pptApp.ActiveWindow;

            // Switch to slide sorter view (where no slide might be selected)
            // ViewType: 1 = Normal, 2 = Outline, 3 = SlideSorter, 7 = NotesPage, 9 = ReadingView
            activeWindow.ViewType = 3; // Slide Sorter View

            // Clear selection in slide sorter view
            dynamic selection = activeWindow.Selection;
            selection.Unselect();

            // Act
            int currentSlideNumber = powerPointService.GetCurrentSlideNumber(presentation);

            // Assert
            Assert.Equal(0, currentSlideNumber);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetTotalSlideCount()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Act
            int totalSlideCount = powerPointService.GetTotalSlideCount(presentation);

            // Assert
            Assert.Equal(1, totalSlideCount);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReadSlideTitleText()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide and replace it with a title slide layout
            dynamic pres = presentation;
            dynamic slides = pres.Slides;

            // Delete the blank slide and add a slide with title layout
            slides[1].Delete();
            dynamic slide = slides.Add(1, 1); // 1 = ppLayoutTitle

            // Set a title on the slide
            dynamic shapes = slide.Shapes;
            dynamic titleShape = shapes[1]; // First shape is typically the title placeholder
            dynamic textFrame = titleShape.TextFrame;
            dynamic textRange = textFrame.TextRange;
            textRange.Text = "Test Slide Title";

            // Act
            string titleText = powerPointService.GetSlideTitle(slide);

            // Assert
            Assert.Equal("Test Slide Title", titleText);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReadAllTextContentFromSlide()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide and add a title and content layout
            dynamic pres = presentation;
            dynamic slides = pres.Slides;

            // Delete the blank slide and add a slide with title and content layout
            slides[1].Delete();
            dynamic slide = slides.Add(1, 2); // 2 = ppLayoutText (Title and Content)

            // Set a title and body text on the slide
            dynamic shapes = slide.Shapes;
            dynamic titleShape = shapes[1]; // First shape is typically the title placeholder
            dynamic titleTextFrame = titleShape.TextFrame;
            dynamic titleTextRange = titleTextFrame.TextRange;
            titleTextRange.Text = "Test Title";

            dynamic contentShape = shapes[2]; // Second shape is typically the content placeholder
            dynamic contentTextFrame = contentShape.TextFrame;
            dynamic contentTextRange = contentTextFrame.TextRange;
            contentTextRange.Text = "Line 1\nLine 2\nLine 3";

            // Act
            string allText = powerPointService.GetAllTextFromSlide(slide);

            // Assert
            Assert.Contains("Test Title", allText);
            Assert.Contains("Line 1", allText);
            Assert.Contains("Line 2", allText);
            Assert.Contains("Line 3", allText);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetAllShapesOnSlide()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide and add a title and content layout
            dynamic pres = presentation;
            dynamic slides = pres.Slides;

            // Delete the blank slide and add a slide with title and content layout
            slides[1].Delete();
            dynamic slide = slides.Add(1, 2); // 2 = ppLayoutText (Title and Content)

            // Act
            var shapes = powerPointService.GetShapesOnSlide(slide);

            // Assert
            Assert.NotNull(shapes);
            Assert.True(shapes.Count > 0, "Slide should have at least one shape");
            // Title and Content layout typically has 2 placeholders (title and content)
            Assert.Equal(2, shapes.Count);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldReturnShapeType()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide and add a title and content layout
            dynamic pres = presentation;
            dynamic slides = pres.Slides;

            // Delete the blank slide and add a slide with title and content layout
            slides[1].Delete();
            dynamic slide = slides.Add(1, 2); // 2 = ppLayoutText (Title and Content)

            // Get the shapes on the slide
            var shapes = powerPointService.GetShapesOnSlide(slide);
            Assert.NotNull(shapes);
            Assert.True(shapes.Count > 0);

            // Act - Get the type of the first shape (should be a placeholder/text box)
            string shapeType = powerPointService.GetShapeType(shapes[0]);

            // Assert
            Assert.NotNull(shapeType);
            Assert.NotEmpty(shapeType);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldAddNewBlankSlide()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get initial slide count
            int initialSlideCount = powerPointService.GetTotalSlideCount(presentation);

            // Act
            var newSlide = powerPointService.AddBlankSlide(presentation);

            // Assert
            Assert.NotNull(newSlide);
            int finalSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(initialSlideCount + 1, finalSlideCount);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldAddSlideWithSpecificLayout()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get initial slide count
            int initialSlideCount = powerPointService.GetTotalSlideCount(presentation);

            // Act
            // 1 = ppLayoutTitle (Title slide layout)
            var newSlide = powerPointService.AddSlideWithLayout(presentation, 1);

            // Assert
            Assert.NotNull(newSlide);
            int finalSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(initialSlideCount + 1, finalSlideCount);

            // Verify the slide has the correct layout by checking for title placeholder
            dynamic slide = newSlide;
            dynamic shapes = slide.Shapes;
            Assert.True(shapes.Count > 0, "Title slide layout should have at least one placeholder");
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldDeleteSlideByIndex()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Add a second slide so we have multiple slides
            powerPointService.AddBlankSlide(presentation);

            // Get initial slide count (should be 2)
            int initialSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(2, initialSlideCount);

            // Act - Delete the second slide (index 2)
            powerPointService.DeleteSlideByIndex(presentation, 2);

            // Assert
            int finalSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(1, finalSlideCount);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldMoveSlideFromOnePositionToAnother()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Add slides with specific layouts to identify them
            // Slide 1: Default blank (already exists)
            // Slide 2: Title layout
            var slide2 = powerPointService.AddSlideWithLayout(presentation, 1); // Title layout
            // Slide 3: Title and Content layout
            var slide3 = powerPointService.AddSlideWithLayout(presentation, 2); // Title and Content layout

            // Verify we have 3 slides
            int totalSlides = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(3, totalSlides);

            // Get the layout of slide at position 1 before moving
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slideAtPos1Before = slides[1];
            int layoutBeforeMove = slideAtPos1Before.Layout;

            // Act - Move slide from position 2 to position 1
            powerPointService.MoveSlide(presentation, 2, 1);

            // Assert
            // Total count should remain the same
            int finalSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(3, finalSlideCount);

            // The slide that was at position 2 should now be at position 1
            dynamic slideAtPos1After = slides[1];
            int layoutAfterMove = slideAtPos1After.Layout;
            Assert.Equal(1, layoutAfterMove); // Title layout (which was at position 2)
            Assert.NotEqual(layoutBeforeMove, layoutAfterMove);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldPreventDeletingWhenOnlyOneSlideExists()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        // Ensure no PowerPoint is running initially
        var initialInstance = powerPointService.GetRunningPowerPointInstance();
        if (initialInstance != null)
        {
            dynamic app = initialInstance;
            app.Quit();
            Marshal.ReleaseComObject(app);
            System.Threading.Thread.Sleep(2000);
        }

        try
        {
            // Start PowerPoint
            pptProcess = new System.Diagnostics.Process();
            pptProcess.StartInfo.FileName = "powerpnt.exe";
            pptProcess.StartInfo.UseShellExecute = true;
            pptProcess.Start();

            // Wait for PowerPoint to fully initialize with retry logic
            object? instance = null;
            int maxRetries = 10;
            int retryDelayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation with only one slide
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Verify we have exactly 1 slide
            int initialSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(1, initialSlideCount);

            // Act & Assert - Attempting to delete the only slide should throw an exception
            var exception = Assert.Throws<InvalidOperationException>(() =>
                powerPointService.DeleteSlideByIndex(presentation, 1));

            // Verify the exception contains meaningful information
            Assert.NotNull(exception);
            Assert.Contains("last slide", exception.Message, StringComparison.OrdinalIgnoreCase);

            // Verify the slide was not deleted
            int finalSlideCount = powerPointService.GetTotalSlideCount(presentation);
            Assert.Equal(1, finalSlideCount);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetFontNameFromTextShape()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        try
        {
            // Start PowerPoint process
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powerpnt.exe",
                UseShellExecute = true,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
            };
            pptProcess = System.Diagnostics.Process.Start(psi);

            // Allow time for PowerPoint to initialize
            const int maxRetries = 20;
            const int retryDelayMs = 500;
            object? instance = null;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a text shape with specific font
            dynamic shapes = slide.Shapes;
            dynamic textBox = shapes.AddTextbox(1, 100, 100, 200, 50); // 1 = msoTextOrientationHorizontal
            dynamic textFrame = textBox.TextFrame;
            dynamic textRange = textFrame.TextRange;
            textRange.Text = "Test Text";

            // Set font name to a known value
            dynamic font = textRange.Font;
            font.Name = "Arial";

            // Act - Get font name from the text shape
            string fontName = powerPointService.GetFontName(textBox);

            // Assert
            Assert.Equal("Arial", fontName);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetFontSizeFromTextShape()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        try
        {
            // Start PowerPoint process
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powerpnt.exe",
                UseShellExecute = true,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
            };
            pptProcess = System.Diagnostics.Process.Start(psi);

            // Allow time for PowerPoint to initialize
            const int maxRetries = 20;
            const int retryDelayMs = 500;
            object? instance = null;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a text shape with specific font size
            dynamic shapes = slide.Shapes;
            dynamic textBox = shapes.AddTextbox(1, 100, 100, 200, 50); // 1 = msoTextOrientationHorizontal
            dynamic textFrame = textBox.TextFrame;
            dynamic textRange = textFrame.TextRange;
            textRange.Text = "Test Text";

            // Set font size to a known value
            dynamic font = textRange.Font;
            font.Size = 24;

            // Act - Get font size from the text shape
            float fontSize = powerPointService.GetFontSize(textBox);

            // Assert
            Assert.Equal(24f, fontSize);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetFontColorFromTextShape()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        try
        {
            // Start PowerPoint process
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powerpnt.exe",
                UseShellExecute = true,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
            };
            pptProcess = System.Diagnostics.Process.Start(psi);

            // Allow time for PowerPoint to initialize
            const int maxRetries = 20;
            const int retryDelayMs = 500;
            object? instance = null;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a text shape with specific font color
            dynamic shapes = slide.Shapes;
            dynamic textBox = shapes.AddTextbox(1, 100, 100, 200, 50); // 1 = msoTextOrientationHorizontal
            dynamic textFrame = textBox.TextFrame;
            dynamic textRange = textFrame.TextRange;
            textRange.Text = "Test Text";

            // Set font color to red (RGB: 255, 0, 0)
            dynamic font = textRange.Font;
            font.Color.RGB = 255; // Red color in COM RGB format (0x0000FF)

            // Act - Get font color (RGB) from the text shape
            (int red, int green, int blue) color = powerPointService.GetFontColor(textBox);

            // Assert
            Assert.Equal(255, color.red);
            Assert.Equal(0, color.green);
            Assert.Equal(0, color.blue);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetShapeFillColor()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        try
        {
            // Start PowerPoint process
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powerpnt.exe",
                UseShellExecute = true,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
            };
            pptProcess = System.Diagnostics.Process.Start(psi);

            // Allow time for PowerPoint to initialize
            const int maxRetries = 20;
            const int retryDelayMs = 500;
            object? instance = null;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a rectangle shape with a fill color
            dynamic shapes = slide.Shapes;
            // AddShape(Type, Left, Top, Width, Height)
            // 1 = msoShapeRectangle
            dynamic rectangle = shapes.AddShape(1, 100, 100, 200, 100);

            // Set fill color to blue (RGB: 0, 0, 255)
            dynamic fill = rectangle.Fill;
            fill.Visible = -1; // msoTrue
            fill.Solid();
            fill.ForeColor.RGB = 16711680; // Blue color in COM RGB format (0xFF0000)

            // Act - Get fill color from the shape
            (int red, int green, int blue) fillColor = powerPointService.GetShapeFillColor(rectangle);

            // Assert
            Assert.Equal(0, fillColor.red);
            Assert.Equal(0, fillColor.green);
            Assert.Equal(255, fillColor.blue);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetShapePosition()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        try
        {
            // Start PowerPoint process
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powerpnt.exe",
                UseShellExecute = true,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
            };
            pptProcess = System.Diagnostics.Process.Start(psi);

            // Allow time for PowerPoint to initialize
            const int maxRetries = 20;
            const int retryDelayMs = 500;
            object? instance = null;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a rectangle shape at a specific position
            dynamic shapes = slide.Shapes;
            // AddShape(Type, Left, Top, Width, Height)
            // 1 = msoShapeRectangle
            dynamic rectangle = shapes.AddShape(1, 150, 200, 100, 50);

            // Act - Get position (Left, Top) from the shape
            (float left, float top) position = powerPointService.GetShapePosition(rectangle);

            // Assert
            Assert.Equal(150f, position.left);
            Assert.Equal(200f, position.top);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetShapeSize()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        try
        {
            // Start PowerPoint process
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powerpnt.exe",
                UseShellExecute = true,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
            };
            pptProcess = System.Diagnostics.Process.Start(psi);

            // Allow time for PowerPoint to initialize
            const int maxRetries = 20;
            const int retryDelayMs = 500;
            object? instance = null;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a rectangle shape with specific size
            dynamic shapes = slide.Shapes;
            // AddShape(Type, Left, Top, Width, Height)
            // 1 = msoShapeRectangle
            dynamic rectangle = shapes.AddShape(1, 100, 100, 250, 150);

            // Act - Get size (Width, Height) from the shape
            (float width, float height) size = powerPointService.GetShapeSize(rectangle);

            // Assert
            Assert.Equal(250f, size.width);
            Assert.Equal(150f, size.height);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldGetTableCellContent()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        try
        {
            // Start PowerPoint process
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powerpnt.exe",
                UseShellExecute = true,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
            };
            pptProcess = System.Diagnostics.Process.Start(psi);

            // Allow time for PowerPoint to initialize
            const int maxRetries = 20;
            const int retryDelayMs = 500;
            object? instance = null;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a table (2 rows, 3 columns) to the slide
            dynamic shapes = slide.Shapes;
            // AddTable(NumRows, NumColumns, Left, Top, Width, Height)
            dynamic table = shapes.AddTable(2, 3, 100, 100, 300, 100);

            // Set text in a specific cell (row 1, column 2)
            dynamic cell = table.Table.Cell(1, 2);
            dynamic cellShape = cell.Shape;
            dynamic textFrame = cellShape.TextFrame;
            dynamic textRange = textFrame.TextRange;
            textRange.Text = "Test Cell Content";

            // Act - Get table cell content
            string cellContent = powerPointService.GetTableCellContent(table, 1, 2);

            // Assert
            Assert.Equal("Test Cell Content", cellContent);
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void ShouldHandleShapesWithoutTextGracefully()
    {
        // Arrange
        var powerPointService = new PowerPointService();
        System.Diagnostics.Process? pptProcess = null;

        try
        {
            // Start PowerPoint process
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powerpnt.exe",
                UseShellExecute = true,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
            };
            pptProcess = System.Diagnostics.Process.Start(psi);

            // Allow time for PowerPoint to initialize
            const int maxRetries = 20;
            const int retryDelayMs = 500;
            object? instance = null;

            for (int i = 0; i < maxRetries; i++)
            {
                System.Threading.Thread.Sleep(retryDelayMs);
                instance = powerPointService.GetRunningPowerPointInstance();
                if (instance != null)
                {
                    break;
                }
            }

            Assert.NotNull(instance);

            // Create a new presentation
            var presentation = powerPointService.CreateNewPresentation(instance);
            Assert.NotNull(presentation);

            // Get the first slide
            dynamic pres = presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides[1];

            // Add a shape without text (rectangle without text)
            dynamic shapes = slide.Shapes;
            // AddShape(Type, Left, Top, Width, Height)
            // 1 = msoShapeRectangle
            dynamic rectangle = shapes.AddShape(1, 100, 100, 200, 100);

            // Act - Try to get font name from a shape without text
            // This should not throw an exception but return null or empty string
            string? fontName = powerPointService.GetFontName(rectangle);

            // Assert - Should handle gracefully by returning null or empty
            Assert.True(string.IsNullOrEmpty(fontName),
                "GetFontName should return null or empty string for shapes without text");
        }
        finally
        {
            // Cleanup - close PowerPoint
            if (pptProcess != null && !pptProcess.HasExited)
            {
                var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
                if (cleanupInstance != null)
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                pptProcess.Kill();
                pptProcess.Dispose();
            }
        }
    }
}
