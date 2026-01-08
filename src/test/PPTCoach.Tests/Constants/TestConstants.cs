namespace PPTCoach.Tests.Constants;

/// <summary>
/// Timeout values for PowerPoint test operations
/// </summary>
public static class TestTimeouts
{
    /// <summary>
    /// Time to wait for PowerPoint to shutdown (ms)
    /// </summary>
    public const int PowerPointShutdownMs = 2000;

    /// <summary>
    /// Delay between retries when waiting for PowerPoint to initialize (ms)
    /// </summary>
    public const int PowerPointInitRetryDelayMs = 1000;

    /// <summary>
    /// Maximum number of retries when waiting for PowerPoint to initialize
    /// </summary>
    public const int PowerPointInitMaxRetries = 20;

    /// <summary>
    /// Time to wait when testing multiple PowerPoint instances (ms)
    /// </summary>
    public const int MultiInstanceWaitMs = 5000;

    /// <summary>
    /// Delay for event detection in monitoring tests (ms)
    /// </summary>
    public const int EventDetectionDelayMs = 500;

    /// <summary>
    /// Time to wait during cleanup operations (ms)
    /// </summary>
    public const int CleanupDelayMs = 2000;
}

/// <summary>
/// MSO text orientation constants
/// </summary>
public static class MsoTextOrientation
{
    public const int Horizontal = 1;
}

/// <summary>
/// MSO AutoShape type constants
/// </summary>
public static class MsoAutoShapeType
{
    public const int Rectangle = 1;
}

/// <summary>
/// Default values for test shape creation
/// </summary>
public static class TestShapeDefaults
{
    public const float DefaultLeft = 100f;
    public const float DefaultTop = 100f;
    public const float DefaultWidth = 200f;
    public const float DefaultHeight = 100f;
    public const float TextBoxHeight = 50f;
}

/// <summary>
/// COM RGB color values for testing
/// COM uses BGR format (Blue-Green-Red)
/// </summary>
public static class TestColors
{
    /// <summary>
    /// Red color: RGB(255, 0, 0) = 0x0000FF in COM BGR format
    /// </summary>
    public const int Red = 0x0000FF;

    /// <summary>
    /// Green color: RGB(0, 255, 0) = 0x00FF00 in COM BGR format
    /// </summary>
    public const int Green = 0x00FF00;

    /// <summary>
    /// Blue color: RGB(0, 0, 255) = 0xFF0000 in COM BGR format
    /// </summary>
    public const int Blue = 0xFF0000;

    /// <summary>
    /// Black color: RGB(0, 0, 0) = 0x000000 in COM BGR format
    /// </summary>
    public const int Black = 0x000000;

    /// <summary>
    /// White color: RGB(255, 255, 255) = 0xFFFFFF in COM BGR format
    /// </summary>
    public const int White = 0xFFFFFF;
}

/// <summary>
/// COM collection indexing constants
/// </summary>
public static class ComIndexing
{
    /// <summary>
    /// COM collections are 1-based (first item is at index 1)
    /// </summary>
    public const int FirstIndex = 1;
}
