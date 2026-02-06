namespace ShellBat;

public partial class ShellBatApplication : CompositionApplication
{
    public ShellBatApplication()
    {
        // custom fatal error handler
        ShowFatalErrorFunc = ShowFatalError;
    }

    public static IEnumerable<ShellBatWindow> ShellBatWindows => AllWindows.OfType<ShellBatWindow>();

    public static bool ShowFatalError(HWND hwnd) => ShowFatalError(hwnd, options =>
    {
        if (Settings.Current.SendErrorReport)
        {
            options.Dialog.Flags |= TASKDIALOG_FLAGS.TDF_VERIFICATION_FLAG_CHECKED;
        }

        options.Dialog.VerificationText = Res.SendErrorReport;
        options.ShownFunc = (options, result) =>
        {
            var info = new ShellBatSysInfo(null);
            var dic = new Dictionary<string, object?>
            {
                { "AppVersion", typeof(ShellBatWindow).Assembly.GetInformationalVersion() },
                { "DirectNVersion", info.DirectNVersion.ToString() },
                { "Adapters", info.Adapters },
                { "Monitor", info.WindowMonitor },
                { "Displays", info.Displays },
                { "WindowDpi", info.WindowDpi },
                { "IsRemoteSession", info.IsRemoteSession },
                { "VirtualMachineType", info.VirtualMachineType },
                { "HasNonWarpAdapter", info.HasNonWarpAdapter },
                { "TextScaleFactor", info.TextScaleFactor },
                { "TokenElevationType", info.TokenElevationType },
                { "OSVersion", Environment.OSVersion.VersionString },
                { "64Bit", Environment.Is64BitProcess },
                { "ProcessorCount", Environment.ProcessorCount },
                { "Version", Environment.Version },
                { "RuntimeIdentifier", RuntimeInformation.RuntimeIdentifier },
                { "FrameworkDescription", RuntimeInformation.FrameworkDescription },
                { "OSDescription", RuntimeInformation.OSDescription },
                { "OSArchitecture", RuntimeInformation.OSArchitecture },
                { "ProcessArchitecture", RuntimeInformation.ProcessArchitecture },
                { "WorkingSet64", Environment.WorkingSet },
                { "MemoryMB", GC.GetGCMemoryInfo().TotalAvailableMemoryBytes / 1024 / 1024 },
                { "UserInteractive", Environment.UserInteractive },
                { "Uptime", TimeSpan.FromMilliseconds(Environment.TickCount64).ToString() },
            };

            var check = options.Dialog.ResultVerificationFlagChecked;
            if (check != Settings.Current.SendErrorReport)
            {
                Settings.Current.SendErrorReport = check;
                try
                {
                    Settings.Current.SerializeToConfiguration();
                }
                catch
                {
                    // hmmm... just go on
                }
            }

            // cancel means it wasn't shown
            if (result != MESSAGEBOX_RESULT.IDCANCEL && options.Dialog.ResultVerificationFlagChecked)
            {
                foreach (var win in AllWindows)
                {
                    win.Cursor = Cursor.Wait.Handle;
                }

                var errors = GetErrors();
                if (errors.Count > 0)
                {
                    var sw = Stopwatch.StartNew();
                    var error = errors[0];
                    var cts = new CancellationTokenSource();
                    var task = Task.Run(() => error.SendErrorReport(dic), cts.Token);
                    var delay = Task.Delay(10000, cts.Token);

                    Task.WaitAny(task, delay);
                    cts.Cancel();
                }
            }

            foreach (var win in AllWindows)
            {
                win.Close();
            }
            ShellBatInstance.Current?.Quit();
        };
    }, false);
}
