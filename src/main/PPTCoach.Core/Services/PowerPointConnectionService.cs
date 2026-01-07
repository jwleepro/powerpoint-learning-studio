using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using PPTCoach.Core.Constants;
using PPTCoach.Core.Interfaces;

namespace PPTCoach.Core.Services;

/// <summary>
/// Service for PowerPoint connection and instance management
/// </summary>
public class PowerPointConnectionService : IPowerPointConnection
{
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid pclsid);

    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [SupportedOSPlatform("windows")]
    public bool IsPowerPointInstalled()
    {
        try
        {
            Type? type = Type.GetTypeFromProgID(PowerPointProgId.Application);
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
            return ConnectToPowerPointOrThrow();
        }
        catch
        {
            // PowerPoint is not running or connection failed
            return null;
        }
    }

    [SupportedOSPlatform("windows")]
    public object ConnectToPowerPointOrThrow()
    {
        // Get CLSID from ProgID
        int clsidHr = CLSIDFromProgID(PowerPointProgId.Application, out Guid clsid);
        if (clsidHr != 0)
        {
            throw new COMException("Failed to get CLSID for PowerPoint.Application", clsidHr);
        }

        // Try to get active object from Running Object Table (ROT)
        int hr = GetActiveObject(ref clsid, IntPtr.Zero, out object obj);

        // S_OK = 0 means success
        const int S_OK = 0;

        if (hr != S_OK)
        {
            throw new COMException("Failed to connect to running PowerPoint instance", hr);
        }

        if (obj == null)
        {
            throw new COMException("PowerPoint instance is null");
        }

        return obj;
    }
}
