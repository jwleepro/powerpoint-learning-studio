using PPTCoach.Core;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace PPTCoach.Tests;

/// <summary>
/// Tests for Phase 1.1: PPT Instance Connection
/// </summary>
public class PowerPointInstanceConnectionTests
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
            PowerPointTestHelpers.CleanupPowerPoint(powerPointService, pptProcess);
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
}
