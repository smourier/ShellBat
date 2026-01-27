namespace ShellBat;

#pragma warning disable CA1822 // Mark members as static
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)]
public class ShellBatSysInfo(Window? window, string separator = " | ")
{
    private readonly Assembly _assembly = Assembly.GetExecutingAssembly();

    public const string ApplicationCategoryName = "Application";
    public const string SystemCategoryName = "System";
    public const string SoftwareCategoryName = "Software";
    public const string GraphicsCategoryName = "Graphics";
    public const string ProcessCategoryName = "Process";

    [LocalizedCategory(SoftwareCategoryName)]
    [DisplayName("DirectN Version")]
    public Version DirectNVersion
    {
        get
        {
            var informationalVersion = typeof(IDXGIFactory).Assembly.GetInformationalVersion();
            if (informationalVersion != null)
            {
                try
                {
                    return new Version(informationalVersion);
                }
                catch
                {
                }
            }

            return new Version(0, 0);
        }
    }

    [LocalizedCategory(SoftwareCategoryName)]
    [DisplayName("WebView2 Version")]
    public string WebView2Version => Program.WebViewVersion ?? "Not Available";

    [LocalizedCategory(GraphicsCategoryName)]
    public uint WindowDpi => (window ?? Application.AllWindows.FirstOrDefault())?.Dpi.width ?? DirectN.Constants.USER_DEFAULT_SCREEN_DPI;

    [LocalizedCategory(SystemCategoryName)]
    public bool IsPenExtensionsInstalled => DirectN.Functions.GetSystemMetrics(SYSTEM_METRICS_INDEX.SM_TABLETPC) != 0;

    [LocalizedCategory(GraphicsCategoryName)]
    public bool HasNonWarpAdapter
    {
        get
        {
            using var factory = DXGIFunctions.CreateDXGIFactory1();
            var adapters = factory.EnumAdapters().ToArray();
            try
            {
                // https://learn.microsoft.com/windows/win32/direct3ddxgi/d3d10-graphics-programming-guide-dxgi#new-info-about-enumerating-adapters-for-windows-8
                const int WarpDeviceId = 0x8c;
                return adapters.Any(adapter => adapter.GetDesc().DeviceId != WarpDeviceId);
            }
            finally
            {
                adapters.Dispose();
            }
        }
    }

    [LocalizedCategory(ProcessCategoryName)]
    public string InstanceName => ShellBatInstance.Current?.IsUnspecified == true ? ShellBatInstance.Current.Name : $"'{ShellBatInstance.Current?.Name}'";

    [LocalizedCategory(ProcessCategoryName)]
    public bool IsLocalServer => ShellBatInstance.Current?.HttpServer.LocalServer != null;

    [LocalizedCategory(ProcessCategoryName)]
    public string? ServerProcessId
    {
        get
        {
            if (ShellBatInstance.Current?.HttpServer.ProcessId == Environment.ProcessId)
                return "(this)";

            return ShellBatInstance.Current?.HttpServer.ProcessId.ToString();
        }
    }

    [LocalizedCategory(ApplicationCategoryName)]
    public string ApplicationCompany => _assembly.GetCompany() ?? string.Empty;

    [LocalizedCategory(ApplicationCategoryName)]
    public string ApplicationProduct => _assembly.GetProduct() ?? string.Empty;

    [LocalizedCategory(ApplicationCategoryName)]
    public string ApplicationCopyright => _assembly.GetCopyright() ?? string.Empty;

    [LocalizedCategory(SoftwareCategoryName)]
    [DisplayName("OS Version")]
    public string OSVersion => Environment.OSVersion.VersionString;

    [LocalizedCategory(SoftwareCategoryName)]
    public string? KernelVersion => WindowsVersionUtilities.KernelVersion?.ToString();

    [LocalizedCategory(SystemCategoryName)]
    public bool IsTabletPC => DirectN.Functions.GetSystemMetrics(SYSTEM_METRICS_INDEX.SM_TABLETPC) != 0;

    [LocalizedCategory(SystemCategoryName)]
    public bool IsRemotelyControlled => DirectN.Functions.GetSystemMetrics(SYSTEM_METRICS_INDEX.SM_REMOTECONTROL) != 0;

    [LocalizedCategory(SystemCategoryName)]
    public bool IsRemoteSession => _isRemoteSession.Value;
    private static readonly Lazy<bool> _isRemoteSession = new(GetIsRemoteSession, true);

    [LocalizedCategory(SystemCategoryName)]
    public string VirtualMachineType => GetVirtualMachineType() ?? Res.NoneDetected;

    [LocalizedCategory(SystemCategoryName)]
    public int ProcessorCount => Environment.ProcessorCount;

    [LocalizedCategory(SystemCategoryName)]
    public string ProcessorArchitecture => SystemUtilities.GetProcessorArchitecture().ToString();

    [LocalizedCategory(SoftwareCategoryName)]
    public string? DotNetVersion => typeof(Uri).Assembly.GetInformationalVersion();

    [LocalizedCategory(ProcessCategoryName)]
    public string TokenElevationType => SystemUtilities.GetTokenElevationType().ToString();

    [LocalizedCategory(ProcessCategoryName)]
    public string Bitness => GetBitness();

    [LocalizedCategory(ProcessCategoryName)]
    public string Culture => CultureInfo.CurrentCulture.Name.Nullify() ?? CultureInfo.CurrentCulture.NativeName.Nullify() ?? CultureInfo.CurrentCulture.TwoLetterISOLanguageName + ", UI: " + CultureInfo.CurrentUICulture.Name.Nullify() ?? CultureInfo.CurrentUICulture.NativeName.Nullify() ?? CultureInfo.CurrentUICulture.TwoLetterISOLanguageName + ", Installed: " + CultureInfo.InstalledUICulture.Name;

    [LocalizedCategory(ProcessCategoryName)]
    public string Now => DateTime.Now + ", Utc: " + DateTime.UtcNow;

    [LocalizedCategory(ProcessCategoryName)]
    [DisplayName("Languages / keyboards")]
    public string Input
    {
        get
        {
            var tag = global::Windows.Globalization.Language.CurrentInputMethodLanguageTag;
            var userLanguages = KeyboardUtilities.GetUserLanguages();
            return string.Join(", ", userLanguages.Select(l => $"{(l.Name.EqualsIgnoreCase(tag) ? "✔ " : string.Empty)}{l.Name} / {l.KeyboardLayout}"));
        }
    }

    [LocalizedCategory(GraphicsCategoryName)]
    public int TextScaleFactor => DpiUtilities.TextScaleFactor;

    [LocalizedCategory(GraphicsCategoryName)]
#pragma warning disable CA1826 // Do not use Enumerable methods on indexable collections
    public string? WindowDpiAwareness => (window ?? Application.AllWindows.FirstOrDefault())?.DpiAwarenessDescription;

    [LocalizedCategory(GraphicsCategoryName)]
    public string? WindowMonitor
    {
        get
        {
            var monitor = (window ?? Application.AllWindows.FirstOrDefault())?.GetMonitor();
            if (monitor == null)
                return null;

            var s = monitor.DeviceName;
            var dd = monitor.DisplayDevice;
            if (dd != null)
            {
                s += " (" + dd.MonitorName + ")";
            }
            return s;
        }
    }
#pragma warning restore CA1826 // Do not use Enumerable methods on indexable collections

    [LocalizedCategory(GraphicsCategoryName)]
    public string DesktopDpi
    {
        get
        {
            var dpi = DpiUtilities.GetDpiForDesktopF();
            return dpi.width + " x " + dpi.height;
        }
    }

    [LocalizedCategory(GraphicsCategoryName)]
    [DisplayName("Graphic Adapter(s)")]
    public string Adapters => GetAdapters(separator);

    [LocalizedCategory(GraphicsCategoryName)]
    [DisplayName("Display(s)")]
    public string Displays => string.Join(separator ?? Environment.NewLine, DisplayConfigQuery());

    [LocalizedCategory(ApplicationCategoryName)]
    public string FileVersion => _assembly.GetFileVersion() ?? "0.0.0";

    [LocalizedCategory(ApplicationCategoryName)]
    public string BuildVersion => _assembly.GetInformationalVersion() ?? "0.0.0.0";

#if DEBUG
    [LocalizedCategory(ApplicationCategoryName)]
    public string AssemblyConfiguration => _assembly.GetConfiguration() ?? string.Empty;
#endif

    public override string ToString()
    {
        var sb = new StringBuilder();
        foreach (var prop in GetType().GetProperties().OrderBy(p => p.Name))
        {
            var value = prop.GetValue(this);
            sb.AppendLine($"{prop.Name}: {value}");
        }

        sb.AppendLine($"ConfigurationDirectoryPath: {Settings.ConfigurationDirectoryPath}");
        sb.AppendLine($"ConfigurationFilePath: {Settings.ConfigurationFilePath}");
        sb.AppendLine($"MonacoPath: {Settings.MonacoPath}");
        sb.AppendLine($"TempDirectoryPath: {Settings.TempDirectoryPath}");
        sb.AppendLine($"ThemesPath: {Settings.ThemesPath}");
        sb.AppendLine($"WebView2UserDataPath: {Settings.WebView2UserDataPath}");
        sb.AppendLine($"WicDecoderFileExtensions: {string.Join(", ", WicImagingComponent.DecoderFileExtensions)}");
        return sb.ToString();
    }

    private static string GetAdapters(string? separator)
    {
        using var fac = DXGIFunctions.CreateDXGIFactory1();
        var adapters = fac.EnumAdapters1().ToArray();
        var str = string.Join(separator ?? Environment.NewLine, adapters.Select(a => a.GetDesc().Description));
        adapters.Dispose();
        return str;
    }

    private static bool GetIsRemoteSession()
    {
        if (DirectN.Functions.GetSystemMetrics(SYSTEM_METRICS_INDEX.SM_REMOTESESSION) != 0)
            return true;

        // https://docs.microsoft.com/en-us/windows/win32/termserv/detecting-the-terminal-services-environment
        using var key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Terminal Server", false);
        if (key != null)
        {
            var value = key.GetValue("GlassSessionId")?.ToString();
            if (value != null)
            {
                if (DirectN.Functions.ProcessIdToSessionId((uint)Environment.ProcessId, out var sessionId) &&
                    uint.TryParse(value, out var id) &&
                    id != sessionId)
                    return true;
            }
        }

        // citrix
        if (Environment.GetEnvironmentVariable("SESSIONNAME")?.StartsWith("ICA", StringComparison.OrdinalIgnoreCase) == true)
            return true;

        return false;
    }

    private static IEnumerable<string> DisplayConfigQuery()
    {
        var dd = DisplayDevice.All.ToList();
        foreach (var path in DisplayConfig.Query())
        {
            var tar = DisplayConfig.GetTargetName(path);
            var src = DisplayConfig.GetSourceName(path);
            var display = dd.FirstOrDefault(m => m.DeviceName.EqualsIgnoreCase(src.viewGdiDeviceName.ToString()));
            if (display == null)
                continue;

            string? dpi = null;
            var mon = display.Monitor;
            if (mon != null)
            {
                dpi = " Dpi(a:" + mon.AngularDpi.width + " e:" + mon.EffectiveDpi.width + " r:" + mon.RawDpi.width + ")";
            }

            if (display.DeviceName == null)
                yield return tar.monitorFriendlyDeviceName + " " + src.viewGdiDeviceName + dpi;

            yield return tar.monitorFriendlyDeviceName + " " + display.DeviceName + dpi;
        }
    }

    private static string GetBitness()
    {
        if (IntPtr.Size == 8)
            return "64-bit";

        if (Environment.Is64BitOperatingSystem)
            return "32-bit on a 64-bit OS";

        return "32-bit";
    }

    private static bool SearchASCIICaseInsensitive(byte[] bytes, string asciiString)
    {
        if (bytes == null || bytes.Length == 0)
            return false;

        var s = Encoding.ASCII.GetBytes(asciiString);
        if (s.Length > bytes.Length)
            return false;

        for (var i = 0; i < bytes.Length; i++)
        {
            var equals = true;
            for (var j = 0; j < s.Length; j++)
            {
                var c1 = (char)bytes[i + j];
                var c2 = (char)s[j];
                if (char.ToLowerInvariant(c1) != char.ToLowerInvariant(c2))
                {
                    equals = false;
                    break;
                }
            }

            if (equals)
                return true;
        }
        return false;
    }

    private static bool IsWindowsSandbox()
    {
        using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Edge\WindowsSandbox", false);
        if (key == null)
            return false;

        return key.GetValue("SandboxEnabled") is int i && i > 0;
    }

    private static string? GetVirtualMachineType()
    {
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\services\mssmbios\Data", false);
            if (key?.GetValue("SMBiosData") is byte[] bytes)
            {
                // detect known emulators
                if (SearchASCIICaseInsensitive(bytes, "Hyper-V"))
                {
                    if (IsWindowsSandbox())
                        return "Windows Sandbox";

                    return "Hyper-V";
                }

                if (SearchASCIICaseInsensitive(bytes, "Microsoft"))
                    return "Virtual PC";

                if (SearchASCIICaseInsensitive(bytes, "VMWare"))
                    return "VMWare";

                if (SearchASCIICaseInsensitive(bytes, "VBox"))
                    return "Virtual Box";

                if (SearchASCIICaseInsensitive(bytes, "Bochs"))
                    return "Bochs";

                if (SearchASCIICaseInsensitive(bytes, "QEMU"))
                    return "QEMU";

                if (SearchASCIICaseInsensitive(bytes, "Plex86"))
                    return "Plex86";

                if (SearchASCIICaseInsensitive(bytes, "Parallels"))
                    return "Parallels";

                if (SearchASCIICaseInsensitive(bytes, "Xen"))
                    return "Xen";

                if (SearchASCIICaseInsensitive(bytes, "Virtual"))
                    return "Generic Virtual Machine";
            }
        }
        catch
        {
        }
        return null;
    }

#pragma warning restore CA1822 // Mark members as static
}
