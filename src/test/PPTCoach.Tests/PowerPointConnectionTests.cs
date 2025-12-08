using PPTCoach.Core;
using System.Runtime.Versioning;

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
}
