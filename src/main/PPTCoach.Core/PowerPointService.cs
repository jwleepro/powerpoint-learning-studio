using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace PPTCoach.Core;

public class PowerPointService
{
    [SupportedOSPlatform("windows")]
    public bool IsPowerPointInstalled()
    {
        try
        {
            Type? type = Type.GetTypeFromProgID("PowerPoint.Application");
            return type != null;
        }
        catch
        {
            return false;
        }
    }

    [SupportedOSPlatform("windows")]
    public object? GetRunningPowerPointInstance()
    {
        try
        {
            // Attempt to get the active PowerPoint application instance
            // from the Running Object Table (ROT)
            // If PowerPoint is not running, this will return null
            
            Type? pptType = Type.GetTypeFromProgID("PowerPoint.Application");
            if (pptType == null)
                return null;

            // Try using reflection to invoke GetActiveObject if available
            var marshalType = typeof(Marshal);
            var method = marshalType.GetMethod("GetActiveObject", 
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static,
                null, 
                new[] { typeof(string) }, 
                null);

            if (method != null)
            {
                return method.Invoke(null, new object[] { "PowerPoint.Application" });
            }

            // Fallback: return null if method doesn't exist or fails
            return null;
        }
        catch
        {
            return null;
        }
    }
}
