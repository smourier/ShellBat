namespace ShellBat;

internal static class Program
{
    public const string AppId = "ShellBat";
    public const string ProgId = "ShellBat.File";

    public const string FromArgumentName = "from";
    public const string WindowRectArgumentName = "wr";
    public const string QuitArgumentName = "quit";
    public const string ParsingNameArgumentName = "d";
    public const string ScreenDevicePathArgumentName = "s";
    public const string ThemeArgumentName = "theme";
    public const string UpdateArgumentName = "update";
    public const string UpdatedArgumentName = "updated";

    public static bool DevEnabled { get; internal set; }
#if DEBUG
        = true;
#endif

    public static string? WebViewVersion { get; private set; }
    public static string WebView2DownloadUrl { get; } = "https://go.microsoft.com/fwlink/p/?LinkId=2124703";
    public static ShellBatWindow? MainWindow { get; private set; }
    public static bool IsSingleFileDeployment { get; } = Assembly.GetEntryAssembly()?.GetCustomAttributes<AssemblyMetadataAttribute>().Where(a => a.Key == "PublishSingleFile").FirstOrDefault()?.Value == "true";
    internal static bool RestartAsAdminRequested { get; set; }

    [STAThread]
    static void Main()
    {
        if (CommandLine.Current.HasArgument("register"))
        {
            if (SystemUtilities.GetTokenElevationType() == TokenElevationType.Limited)
            {
                ShellBatInstance.RegisterComObjects(Registry.CurrentUser);
            }
            else
            {
                ShellBatInstance.RegisterComObjects(Registry.LocalMachine);
            }
            return;
        }

        if (CommandLine.Current.HasArgument("unregister"))
        {
            if (SystemUtilities.GetTokenElevationType() == TokenElevationType.Limited)
            {
                ShellBatInstance.UnregisterComObjects(Registry.CurrentUser);
            }
            else
            {
                ShellBatInstance.UnregisterComObjects(Registry.LocalMachine);
            }
            return;
        }

        if (CommandLine.Current.HasArgument("diags"))
        {
            var sysInfo = new ShellBatSysInfo(null);
            MessageBox.Show(sysInfo.ToString(), $"{AppId} Info", MESSAGEBOX_STYLE.MB_OK | MESSAGEBOX_STYLE.MB_ICONINFORMATION);
            return;
        }

        // this seems the only reliable way to set the default background color to 'nothing/transparent' for WebView2
        // so we can use RootVisual's brush as background
        // note: strangely this work w/o this in normal VS runs (debug and release), but not after publishing as AOT (maybe some race condition?)
        // cf https://learn.microsoft.com/en-us/dotnet/api/microsoft.web.webview2.core.corewebview2controller.defaultbackgroundcolor#remarks
        Environment.SetEnvironmentVariable("WEBVIEW2_DEFAULT_BACKGROUND_COLOR", "00000000");

        _ = WebRootResources.EnsureFilesAsync();
        _ = ThemeResources.EnsureFilesAsync();
        _ = MonacoResources.EnsureFilesAsync();

        ShellBatInstance.WithInstance(instance =>
        {
            if (!WindowsVersionUtilities.IsWindows10OrGreater())
            {
                var td = new TaskDialog
                {
                    Title = typeof(Program).Namespace,
                    MainIcon = new HICON { Value = TaskDialog.TD_WARNING_ICON },
                    MainInstruction = Res.Error,
                    Content = string.Format(Res.UnsupportedWindowsVersion, WindowsVersionUtilities.KernelVersion),
                };
                td.Show(HWND.Null, false);
                return;
            }

            WebView2Utilities.Initialize(Assembly.GetEntryAssembly());
            WebViewVersion = WebView2Utilities.GetAvailableCoreWebView2BrowserVersionString();

            if (string.IsNullOrWhiteSpace(WebViewVersion))
            {
                unsafe
                {
                    delegate* unmanaged<HWND, uint, WPARAM, LPARAM, nint, HRESULT> callback = &TaskDialogCallback;
                    var td = new TaskDialog
                    {
                        Title = typeof(Program).Namespace,
                        MainIcon = new HICON { Value = TaskDialog.TD_WARNING_ICON },
                        Flags = TASKDIALOG_FLAGS.TDF_ENABLE_HYPERLINKS,
                        MainInstruction = Res.WebView2NotFound,
                        Content = string.Format(Res.WebView2Link, WebView2DownloadUrl),
                        Callback = (nint)callback,
                    };
                    td.Show(HWND.Null, false);
                }
                return;
            }

            AppDomain.CurrentDomain.FirstChanceException += (s, e) =>
            {
                if (e.Exception is DllNotFoundException dll && dll.Message.Contains("dxgidebug"))
                    return;

                // HttpListener might throw this (as 1st chance) when closing
                if (e.Exception is HttpListenerException)
                    return;

                if (e.Exception is ObjectDisposedException ode)
                {
                    if (ode.ObjectName?.Contains("System.Net.") == true)
                        return;
                }

                if (e.Exception is PlatformNotSupportedException pnse)
                {
                    if (pnse.Message.Contains("OID"))
                        return;
                }

                ShellBatInstance.LogError("FirstChanceException", e.Exception.ToString());
            };

            _ = Task.Run(Settings.Current.Backup);

            using var app = new ShellBatApplication();
            using var win = new ShellBatWindow();
            MainWindow = win;
            var maximized = false;
            var monitor = WebScreen.Get(CommandLine.Current.GetNullifiedArgument(ScreenDevicePathArgumentName))?.Monitor;
            if (monitor != null)
            {
                StartInMonitor(win, monitor);
            }
            else
            {
                var wr = CommandLine.Current.GetNullifiedArgument(WindowRectArgumentName);
                if (wr != null && RECT.TryParse(wr, null, out var rc))
                {
                    win.ResizeAndMove(rc);
                }
                else
                {
                    var noAutomaticPlacement = CommandLine.Current.HasArgument("nap") || (VIRTUAL_KEY.VK_SHIFT.IsPressed() && VIRTUAL_KEY.VK_CONTROL.IsPressed());
                    var initialRect = instance.Settings.GetInitialRect();
                    if (initialRect == null || noAutomaticPlacement)
                    {
                        monitor = win.GetMonitor();
                        StartInMonitor(win, monitor);
                    }
                    else
                    {
                        maximized = instance.Settings.MainPlacement?.showCmd == (uint)SHOW_WINDOW_CMD.SW_MAXIMIZE;
                    }
                }
            }

            if (maximized)
            {
                win.Show(SHOW_WINDOW_CMD.SW_MAXIMIZE);
            }
            else
            {
                win.Show();
            }

            app.Run();
            MainWindow = null;

            // if we serve HTTP, we need to keep running until all othere instances depending on us are closed
            if (!RestartAsAdminRequested && instance.Options.HasFlag(ShellBatInstanceOptions.IsLocalHttpServer))
            {
                instance.WaitForHttpServerStop();
                ShellBatInstance.LogInfo($"Local HTTP server stopped at {instance.HttpServer.Url}");
                instance.FlushLogs();
            }
            else
            {
                instance.FlushLogs();
            }
        });
    }

    private static void StartInMonitor(Window window, DirectN.Extensions.Utilities.Monitor? monitor)
    {
        monitor ??= DirectN.Extensions.Utilities.Monitor.Primary;
        if (monitor != null)
        {
            var area = monitor.WorkingArea;
            var width = area.Width * 3 / 4;
            var height = area.Height * 3 / 4;
            window.ResizeAndMove(RECT.Sized(
                area.left + (area.Width - width) / 2,
                area.top + (area.Height - height) / 2,
                width,
                height));
        }
        else // huh?
        {
            window.ResizeClient(800, 600);
            window.Center();
        }
    }

    internal static Icon? LoadIcon()
    {
        var exeHandle = DirectN.Functions.GetModuleHandleW(PWSTR.From(Environment.ProcessPath));
        var icon = Icon.FromHandle(DirectN.Functions.LoadImageW(new HINSTANCE { Value = exeHandle.Value }, new(ShellBatInstance.ShellBatIconIndex), GDI_IMAGE_TYPE.IMAGE_ICON, 16, 16, 0).Value, true);
        return icon;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(TaskDialogCallback))]
    private static HRESULT TaskDialogCallback(HWND hwnd, uint msg, WPARAM wParam, LPARAM lParam, nint lpRefData)
    {
        if (msg == (uint)TASKDIALOG_NOTIFICATIONS.TDN_HYPERLINK_CLICKED)
        {
            Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = WebView2DownloadUrl });
            return DirectN.Constants.E_FAIL;
        }
        return DirectN.Constants.S_OK;
    }
}