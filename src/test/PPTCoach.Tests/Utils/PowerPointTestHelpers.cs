using PPTCoach.Core;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace PPTCoach.Tests.Utils;

/// <summary>
/// Helper methods for PowerPoint test setup and cleanup
/// </summary>
public static class PowerPointTestHelpers
{
    /// <summary>
    /// Clears PowerPoint's "Resiliency" registry keys to prevent safe mode dialog
    /// This removes the flag that PowerPoint sets when it crashes
    /// </summary>
    private static void ClearPowerPointResiliencyKeys()
    {
        try
        {
            // PowerPoint stores crash recovery information in these registry paths
            string[] resiliencyPaths = new[]
            {
                @"Software\Microsoft\Office\16.0\PowerPoint\Resiliency",
                @"Software\Microsoft\Office\15.0\PowerPoint\Resiliency",
                @"Software\Microsoft\Office\14.0\PowerPoint\Resiliency"
            };

            foreach (var path in resiliencyPaths)
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(path, writable: true))
                    {
                        if (key != null)
                        {
                            // Delete subkeys that store crash information
                            foreach (var subKeyName in key.GetSubKeyNames())
                            {
                                key.DeleteSubKeyTree(subKeyName, throwOnMissingSubKey: false);
                            }
                        }
                    }
                }
                catch
                {
                    // Ignore if key doesn't exist or can't be accessed
                }
            }
        }
        catch
        {
            // Ignore registry errors - not critical for test execution
        }
    }

    /// <summary>
    /// Ensures no PowerPoint instance is running before starting a test
    /// </summary>
    public static void EnsureNoPowerPointRunning(PowerPointService powerPointService)
    {
        var existingInstance = powerPointService.GetRunningPowerPointInstance();
        if (existingInstance != null)
        {
            try
            {
                dynamic app = existingInstance;
                app.Quit();
                Marshal.ReleaseComObject(app);
            }
            catch
            {
                // Ignore cleanup errors
            }
            System.Threading.Thread.Sleep(2000);
        }

        // Clear resiliency keys to prevent safe mode dialog
        ClearPowerPointResiliencyKeys();
    }

    /// <summary>
    /// Starts PowerPoint and waits for it to be ready
    /// </summary>
    /// <returns>Tuple of (Process, COM Instance)</returns>
    public static (System.Diagnostics.Process process, object instance) StartPowerPointAndWait(
        PowerPointService powerPointService,
        System.Diagnostics.ProcessWindowStyle windowStyle = System.Diagnostics.ProcessWindowStyle.Normal)
    {
        var pptProcess = new System.Diagnostics.Process();
        pptProcess.StartInfo.FileName = "powerpnt.exe";
        pptProcess.StartInfo.UseShellExecute = true;
        pptProcess.StartInfo.WindowStyle = windowStyle;
        pptProcess.Start();

        // Wait for PowerPoint to fully initialize with retry logic
        object? instance = null;
        int maxRetries = 20; // Increased from 10 to 20
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

        if (instance == null)
        {
            throw new InvalidOperationException(
                $"Failed to connect to PowerPoint after {maxRetries} retries. " +
                "PowerPoint may not be starting correctly or may require more time to initialize.");
        }

        return (pptProcess, instance);
    }

    /// <summary>
    /// Cleans up PowerPoint instance in finally block
    /// </summary>
    public static void CleanupPowerPoint(PowerPointService powerPointService, System.Diagnostics.Process? pptProcess)
    {
        if (pptProcess != null && !pptProcess.HasExited)
        {
            var cleanupInstance = powerPointService.GetRunningPowerPointInstance();
            if (cleanupInstance != null)
            {
                try
                {
                    dynamic app = cleanupInstance;
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }

            try
            {
                pptProcess.Kill();
                pptProcess.Dispose();
            }
            catch
            {
                // Ignore if process already exited
            }
        }
    }
}
