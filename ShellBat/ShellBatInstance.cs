namespace ShellBat;

public partial class ShellBatInstance : IShellBatInstance, IDisposable
{
    // sync this with ShellBat.csproj
    public const int ShellBatIconIndex = 1;
    public const int ShellBatThemeIconIndex = 2;

    public static Guid MonikerPrefix { get; } = new("7a5feab0-84a2-453f-bee0-6c907c7b2d90");
    public static Guid ConnectionPointInterfaceId { get; } = new("9f7f3b40-7826-4e3b-9510-958f1a899a1d");
    public static Guid CommandGroupId { get; } = new("5d8713f8-eb1d-4cca-8d4c-c7bdc03f5ec6");

    public static ShellBatInstance Current { get; private set; } = null!;
    public static string LogsDirectoryPath => Path.Combine(Path.GetTempPath(), Program.AppId, "Logs");

    public event EventHandler<ValueEventArgs<ShellBatInstanceCommand>>? Command;

    // write all traces on a single thread to avoid garbling the output
    private static string? _updateAuthId;
    private static bool _updatingMode;
    private int _threadId;
    private TextWriter? _logWriter;
    private readonly SingleThreadTaskScheduler? _logWriterScheduler;
    private readonly EventWaitHandle _notifyEvent;
    private readonly AutoResetEvent _serverStopEvent;
    private readonly AutoResetEvent _notifyStopEvent;
    private readonly Thread _thread;
    private readonly int _instancesUpdateInterval;
    private bool _fromHandled;
    private CommandTarget? _commandTarget;
    private CommandTargetConnectionPoint? _connectionPoint;
    private ShellBatRemoteInstance[] _remoteInstances;
    private Timer? _updateInstancesTimer;
    private ComServer? _comServer;
    private ShellBatTheme _theme;

    private ShellBatInstance()
    {
        Current = this;
        FromProcessId = CommandLine.Current.GetArgument<int>(Program.FromArgumentName);

        string? instanceName = null;
        var firstArg = GetFirstArgumentFilePath();
        if (firstArg != null &&
            Path.GetExtension(firstArg).EqualsIgnoreCase(FileExtension.ShellBatExtensionName) &&
            IOUtilities.PathIsFile(firstArg))
        {
            SettingsFilePath = firstArg;
            instanceName = Path.GetFileNameWithoutExtension(SettingsFilePath);
            CommandLine.Current.PositionedArguments.Remove(0);
        }
        else
        {
            // build settings
            SettingsFilePath = Path.Combine(ShellBat.Settings.ConfigurationDirectoryPath, $"{Name}.settings.json");
        }

        if (IOUtilities.PathIsFile(SettingsFilePath))
        {
            Settings = ShellBatInstanceSettings.Deserialize(SettingsFilePath, JsonSourceNoDefaultGenerationContext.Default, true) ?? new();
            Settings.FilePath = SettingsFilePath;
        }
        else
        {
            Settings = new() { FilePath = SettingsFilePath };
            SerializeSettings();
        }

        if (Settings.ShowDevTools != null)
        {
            Program.DevEnabled = Settings.ShowDevTools.Value;
        }

        // we need to handle our position if it was'nt specified from another instance
        _fromHandled = FromProcessId == 0 || CommandLine.Current.HasArgument(Program.ScreenDevicePathArgumentName) || CommandLine.Current.HasArgument(Program.WindowRectArgumentName);

        // note there can be multiple instances with unspecified names
        Name = Settings.InstanceName ?? instanceName ?? Res.UnspecifiedInstanceName;
        if (IOUtilities.NameToValidFileName(Name) != Name)
            throw new ShellBatException($"0010: Instance name '{Name}' is invalid.");

        if (Settings.WebView2UserDataPath != null)
        {
            ShellBat.Settings.WebView2UserDataPath = Settings.WebView2UserDataPath;
            ShellBat.Settings.HttpServerCachePath = Path.Combine(Settings.WebView2UserDataPath, ShellBat.Settings.CacheDirectoryName);
        }

        SetHistory();

        Theme = ShellBatTheme.LoadTheme(firstArg);

        _serverStopEvent = new AutoResetEvent(false);
        _notifyStopEvent = new AutoResetEvent(false);
        _notifyEvent = NotifyEvent.CreateEvent(MonikerPrefix.ToString("N"));
        if (_notifyEvent is null)
            throw new ShellBatException($"0002: Cannot create notify event named '{MonikerPrefix}'.");

        if (Settings.LogToFile)
        {
            _logWriterScheduler = new(thread =>
            {
                _threadId = thread.ManagedThreadId;
                thread.Name = $"{Program.AppId} Trace Writer";

                var dir = LogsDirectoryPath;
                if (!IOUtilities.PathIsDirectory(dir))
                {
                    try
                    {
                        Directory.CreateDirectory(dir);
                    }
                    catch
                    {
                        // ignore
                    }
                }

                _logWriter = new StreamWriter(Path.Combine(dir, $"{Guid.NewGuid():N}.{Program.AppId}.{Name}.log")) { AutoFlush = true };
                _logWriter?.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][{_threadId}]Log starting, bitness: {nint.Size} process: {SystemUtilities.CurrentProcess.MainModule?.FileName ?? SystemUtilities.CurrentProcess.ProcessName} id: {Environment.ProcessId} lowbox: {SystemUtilities.IsAppContainer()}");
                _logWriter?.WriteLine($"Architecture: {SystemUtilities.GetProcessorArchitecture()} Windows: {WindowsVersionUtilities.KernelVersion}");
                _logWriter?.WriteLine($"Command line: {Environment.CommandLine}");
                return true;
            });
        }

        _thread = new Thread(NotifyEventsReceivingThread)
        {
            IsBackground = true,
            Name = $"{Program.AppId}NotifyEvents"
        };
        _thread.Start();

        _commandTarget = new CommandTarget($"{Program.AppId}:{MonikerPrefix:N}:{Name}");
        _commandTarget.Command += OnCommand;
        _commandTarget.Register();

        _connectionPoint = new CommandTargetConnectionPoint(ConnectionPointInterfaceId);
        _commandTarget.AddConnectionPoint(_connectionPoint);

        _remoteInstances = ShellBatRemoteInstance.UpdateInstances();

        _instancesUpdateInterval = ShellBat.Settings.Current.InstancesUpdateInterval;
        if (_instancesUpdateInterval > 0)
        {
            _updateInstancesTimer = new Timer(_ => UpdateRemoteInstances(), null, _instancesUpdateInterval, Timeout.Infinite);
        }

        HttpServer = ShellBatHttpLocalServer.Get(this);
        if (HttpServer.LocalServer != null)
        {
            Task.Run(HttpServer.LocalServer.Start);
            LogInfo($"Instance name '{Name}'. Local HTTP server started at {HttpServer.Url}");
            Options = ShellBatInstanceOptions.IsLocalHttpServer;
        }
        else
        {
            LogInfo($"Instance name '{Name}'. Remote HTTP server detected at {HttpServer.Url}");
            Options = ShellBatInstanceOptions.None;
        }

        _comServer = new ComServer();
        _comServer.RegisterClassObject(() => new ShellBatShellExtension());
    }

    public string Name { get; } = Res.UnspecifiedInstanceName;
    public bool IsUnspecified => Name.EqualsIgnoreCase(Res.UnspecifiedInstanceName);
    public int ProcessId => Environment.ProcessId;
    public int FromProcessId { get; }
    public string SettingsFilePath { get; }
    public ShellBatInstanceSettings Settings { get; }
    public IReadOnlyList<ShellBatRemoteInstance> RemoteInstances => _remoteInstances;
    public ShellBatHttpServer HttpServer { get; }
    public string? CurrentParsingName => Program.MainWindow?.CurrentEntry?.FullDisplayName;
    public ShellBatInstanceOptions Options { get; protected set; }

    public ShellBatTheme Theme
    {
        get => _theme;
        [MemberNotNull(nameof(_theme))]
        set
        {
            Settings.ThemeFilePath = value.FilePath;
            ArgumentNullException.ThrowIfNull(value);
            _theme = value;
            foreach (var window in Application.AllWindows.OfType<ShellBatWindow>())
            {
                window.ApplyTheme(_theme);
            }
        }
    }

    internal void SetHistory()
    {
        if (Settings?.IsInstanceScopedHistory == true)
        {
            History.SetHistory(Path.Combine(Path.GetDirectoryName(SettingsFilePath)!, $"{Name}.history.json"));
        }
        else
        {
            History.SetHistory(Path.Combine(ShellBat.Settings.ConfigurationDirectoryPath, History.GlobalFileName));
        }
    }

    internal static string? GetFirstArgumentFilePath()
    {
        var firstArg = CommandLine.Current.GetNullifiedArgument(0);
        if (firstArg == null)
            return null;

        if (IOUtilities.IsPathRooted(firstArg) && IOUtilities.PathExists(firstArg))
            return firstArg;

        firstArg = Path.Combine(Environment.CurrentDirectory, firstArg);
        if (IOUtilities.IsPathRooted(firstArg) && IOUtilities.PathExists(firstArg))
            return firstArg;

        return null;
    }

    HRESULT IShellBatInstance.SetWindowRect(RECT rc, Guid? desktopId)
    {
        var window = Program.MainWindow;
        if (window is null)
            return DirectN.Constants.E_FAIL;

        window.WindowRect = rc;
        return DirectN.Constants.S_OK;
    }

    public IEnumerable<IShellBatInstance> All
    {
        get
        {
            yield return this;
            foreach (var instance in RemoteInstances)
            {
                yield return instance;
            }
        }
    }

    public ShellBatRemoteInstance? GetRemoteInstance(int processId) => _remoteInstances.FirstOrDefault(i => i.ProcessId == processId);

    public void BackupSettings() => Settings.Backup(TimeSpan.FromDays(ShellBat.Settings.Current.BackupsMaxDays));
    public void SerializeSettings(int deferTime = 300) => Settings.Serialize(JsonSourceNoDefaultGenerationContext.Default, deferTime);

    public void FlushLogs()
    {
        // don't use scheduler thread for this
        _logWriter?.Flush();
        _logWriterScheduler?.TriggerDequeue();
    }

    internal void WaitForHttpServerStop()
    {
        // are there any remote instances?
        var list = UpdateRemoteInstances();
        if (list.Length == 0)
            return;

        // are there any remote instances connected to our http server?
        var any = false;
        foreach (var instance in list)
        {
            var url = instance.GetConnectedToHttpServerUrl(null);
            if (url is null)
                continue;

            if (HttpServer.LocalServer!.Url.ToString().EqualsIgnoreCase(url.ToString()))
            {
                any = true;
                break;
            }
        }
        if (!any)
            return;

        Options |= ShellBatInstanceOptions.IsInHttpServerOnlyMode;
        NotifyOtherInstances();
        _serverStopEvent.WaitOne();
    }

    protected virtual void OnCommand(object? sender, CommandTargetEventArgs e)
    {
        var cmd = ShellBatInstanceCommand.Parse(e);
        if (cmd != null)
        {
            var evt = new ValueEventArgs<ShellBatInstanceCommand>(cmd, isCancellable: true);
            Command?.Invoke(this, evt);
            if (evt.Cancel)
                return;

            Window? mainWindow;
            switch (cmd.Id)
            {
                // do we serve HTTP files?
                case ShellBatInstanceCommandId.HttpServerUrl:
                    if (HttpServer.LocalServer != null)
                    {
                        e.Output = HttpServer.LocalServer.Url.ToString();
                    }
                    e.HResult = DirectN.Constants.S_OK;
                    break;

                // we don't server HTTP files, so where is the server?
                case ShellBatInstanceCommandId.ConnectedToHttpServerUrl:
                    if (HttpServer.LocalServer == null)
                    {
                        e.Output = HttpServer.Url.ToString();
                    }
                    e.HResult = DirectN.Constants.S_OK;
                    break;

                case ShellBatInstanceCommandId.Quit:
                    if (_updatingMode)
                    {
                        LogInfo($"Quit command received from process id {cmd.ProcessId} during updating mode, ignoring...");
                        e.HResult = DirectN.Constants.E_FAIL;
                        break;
                    }

                    LogInfo($"Quit command received from process id {cmd.ProcessId}, exiting...");
                    Quit();
                    e.HResult = DirectN.Constants.S_OK;
                    break;

                case ShellBatInstanceCommandId.GetCurrentParsingName:
                    if (Program.MainWindow is ShellBatWindow fw && fw.CurrentEntry is not null)
                    {
                        e.Output = fw.CurrentEntry.ParsingName;
                        e.HResult = DirectN.Constants.S_OK;
                    }
                    else
                    {
                        e.HResult = DirectN.Constants.E_FAIL;
                    }
                    break;

                case ShellBatInstanceCommandId.GetOptions:
                    e.Output = (int)Options;
                    e.HResult = DirectN.Constants.S_OK;
                    break;

                case ShellBatInstanceCommandId.GetMainWindowRect:
                    mainWindow = Program.MainWindow;
                    if (mainWindow != null)
                    {
                        e.Output = mainWindow.WindowRect.ToString();
                        e.HResult = DirectN.Constants.S_OK;
                    }
                    else
                    {
                        e.HResult = DirectN.Constants.E_FAIL;
                    }
                    break;

                case ShellBatInstanceCommandId.SetMainWindowRect:
                    mainWindow = Program.MainWindow;
                    if (mainWindow != null &&
                        cmd.Arguments.Length == 1 &&
                        cmd.Arguments[0] is string text &&
                        RECT.TryParse(text, null, out var rect) &&
                        rect.Width > 0 &&
                        rect.Height > 0)
                    {
                        mainWindow.WindowRect = rect;
                        e.HResult = DirectN.Constants.S_OK;
                        break;
                    }
                    e.HResult = DirectN.Constants.E_FAIL;
                    break;

                case ShellBatInstanceCommandId.SwitchTo:
                    mainWindow = Program.MainWindow;
                    if (mainWindow != null)
                    {
                        mainWindow.Show(SHOW_WINDOW_CMD.SW_SHOWNORMAL);
                        mainWindow.SetForeground();
                        e.HResult = DirectN.Constants.S_OK;
                        break;
                    }
                    e.HResult = DirectN.Constants.E_FAIL;
                    break;

                case ShellBatInstanceCommandId.ShowProperties:
                    if (Program.MainWindow is ShellBatWindow fw2 &&
                        cmd.Arguments.Length == 1 &&
                        cmd.Arguments[0] is string parsingName)
                    {
                        fw2.ShowProperties(parsingName, new WebWindowShowOptions
                        {
                            ForceOpen = true,
                        }, cmd);
                        e.HResult = DirectN.Constants.S_OK;
                        break;
                    }

                    e.HResult = DirectN.Constants.E_FAIL;
                    break;

                case ShellBatInstanceCommandId.ContinueShowProperties:
                    if (Program.MainWindow is ShellBatWindow fw3 &&
                        fw3.IsPropertiesLockedBy(cmd) &&
                        cmd.Arguments.Length == 1 &&
                        cmd.Arguments[0] is string parsingName2)
                    {
                        fw3.ShowProperties(parsingName2, new WebWindowShowOptions
                        {
                            ForceOpen = true,
                        }, cmd);
                        e.HResult = DirectN.Constants.S_OK;
                        break;
                    }

                    e.HResult = DirectN.Constants.E_FAIL;
                    break;

                case ShellBatInstanceCommandId.Update:
                    if (cmd.Arguments.Length != 1 ||
                        cmd.Arguments[0] is not string updateAuthId ||
                        updateAuthId != _updateAuthId)
                    {
                        e.HResult = DirectN.Constants.E_ACCESSDENIED;
                        break;
                    }

                    QuitAllRemoteInstances(); // this is supposed to be called before launching the updater, but we try again here
                    e.Output = Environment.ProcessPath;
                    e.HResult = DirectN.Constants.S_OK;
                    Task.Run(Quit);
                    break;
            }
        }
    }

    public virtual void SendEventToRemotes(ShellBatInstanceRemoteEventType type, params object?[]? parameters) => _connectionPoint?.InvokeMember((int)type, parameters);

    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            Settings?.SafeDispose();
            Interlocked.Exchange(ref _comServer, null)?.SafeDispose();
            Interlocked.Exchange(ref _updateInstancesTimer, null)?.SafeDispose();
            HttpServer.LocalServer?.SafeDispose();
            _notifyStopEvent.Set();
            _thread.Join(10000);
            _notifyEvent?.Dispose();
            Interlocked.Exchange(ref _connectionPoint, null)?.SafeDispose();
            Interlocked.Exchange(ref _commandTarget, null)?.SafeDispose();
        }
    }

    private void TraceToWriter(string? text = null, string? filePath = null)
    {
        if (Environment.CurrentManagedThreadId == _threadId)
        {
            trace();
            return;
        }

        var scheduler = _logWriterScheduler;
        if (scheduler != null)
        {
            _ = Task.Factory.StartNew(trace, CancellationToken.None, TaskCreationOptions.None, scheduler);
        }

        void trace()
        {
            _logWriter?.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][{Environment.CurrentManagedThreadId}]{(filePath != null ? Path.GetFileNameWithoutExtension(filePath) : null)}: {text}");
        }
    }

    // this is needed to be able to detect changes in instances' Options
    private sealed class InstanceComparer : IEqualityComparer<ShellBatRemoteInstance>
    {
        public static InstanceComparer Instance { get; } = new();

        public int GetHashCode([DisallowNull] ShellBatRemoteInstance obj) => HashCode.Combine(obj.ProcessId, obj.Options);
        public bool Equals(ShellBatRemoteInstance? x, ShellBatRemoteInstance? y)
        {
            if (x is null)
                return y is null;

            if (y is null)
                return false;

            return x.ProcessId == y.ProcessId && x.Options == y.Options;
        }
    }

    private ShellBatRemoteInstance[] UpdateRemoteInstances()
    {
        _updateInstancesTimer?.Change(_instancesUpdateInterval, Timeout.Infinite);
        var list = ShellBatRemoteInstance.UpdateInstances();
        var areEqual = list.SequenceEqual(_remoteInstances, InstanceComparer.Instance);

        Interlocked.Exchange(ref _remoteInstances, list);
        foreach (var instance in list)
        {
            instance.UpdateOptions();
            //LogVerbose($"Remote instance: {instance} options: {instance.Options}");
        }

        if (!areEqual)
        {
            foreach (var window in ShellBatApplication.ShellBatWindows)
            {
                window.UpdateRemoteInstances(list);
            }
        }

        if (!_fromHandled)
        {
            var fromInstance = list.FirstOrDefault(i => i.ProcessId == FromProcessId);
            if (fromInstance != null)
            {
                _fromHandled = true;
                var rc = fromInstance.GetMainWindowRect();
                if (rc != null)
                {
                    var newRc = rc.Value;
                    var padding = Math.Max(ShellBatInstanceSettings.DefaultDockPadding, Settings.DockPadding) * 2;
                    newRc.left += padding;
                    newRc.top += padding;
                    newRc.right += padding;
                    newRc.bottom += padding;
                    CommandLine.Current.NamedArguments[Program.WindowRectArgumentName] = newRc.ToString();
                }
            }
        }
        return list;
    }

    private void NotifyEventsReceivingThread()
    {
        do
        {
            var index = WaitHandle.WaitAny([_notifyStopEvent, _notifyEvent]);
            if (index == 0)
                break;

            // refresh if we're called
            var list = UpdateRemoteInstances();

            // if no instances and we're in http-server-only mode, exit
            if (list.Length == 0 && Options.HasFlag(ShellBatInstanceOptions.IsInHttpServerOnlyMode))
            {
                _serverStopEvent.Set();
                break;
            }
        }
        while (true);
    }

    // notify of life or death basically
    public static void NotifyOtherInstances() => NotifyEvent.RaiseEvent(MonikerPrefix.ToString("N"));
    public static void WithInstance(Action<ShellBatInstance> action)
    {
        ArgumentNullException.ThrowIfNull(action);

        var quit = CommandLine.Current.GetArgument<bool>(Program.QuitArgumentName);
        if (quit)
        {
            var remotes = ShellBatRemoteInstance.UpdateInstances();
            foreach (var other in remotes)
            {
                // drain results to ensure completion
                _ = other.Quit().ToArray();
            }
            return;
        }

        var instance = new ShellBatInstance();

        var updateAuthId = CommandLine.Current.GetNullifiedArgument(Program.UpdateArgumentName);
        if (updateAuthId != null)
        {
            try
            {
                ShellBatUpdateWindow.ShowProgress(updateAuthId);
            }
            catch (Exception ex)
            {
                LogError($"Cannot update to latest: {ex}");
            }
            finally
            {
                instance.Dispose();
            }
            return;
        }

        // if control is pressed, we open a new instance whatsover
        if (!VIRTUAL_KEY.VK_CONTROL.IsPressed() &&
            instance.Settings.InstanceName != null &&
            !instance.Settings.InstanceName.EqualsIgnoreCase(Res.UnspecifiedInstanceName))
        {
            var remoteInstance = instance.RemoteInstances.FirstOrDefault(i => i.Name.EqualsIgnoreCase(instance.Name));
            if (remoteInstance != null)
            {
                SwitchToInstance(remoteInstance.ProcessId);
                instance.Dispose();
                return;
            }
        }

        try
        {
            // tell others we're here
            NotifyOtherInstances();

            action(instance);
        }
        finally
        {
            instance.Dispose();

            // tell others we're gone
            NotifyOtherInstances();
        }
    }

    private static IEnumerable<string> GetOpenWithPaths()
    {
        yield return $@"Software\Classes\Folder\shell\OpenWith{Program.AppId}";
        yield return $@"Software\Classes\Directory\Background\shell\OpenWith{Program.AppId}";
    }

    public static void RegisterComObjects(RegistryKey key)
    {
        ArgumentNullException.ThrowIfNull(key);

        ComServer.Register<ShellBatShellExtension>(key);

        foreach (var path in GetOpenWithPaths())
        {
            using var shell = key.EnsureSubKey(path);
            if (shell != null)
            {
                shell.SetValue(null, Res.OpenWithShellBat);
                shell.SetValue("Icon", $"{Environment.ProcessPath},-{ShellBatIconIndex}");
                shell.SetValue("SeparatorBefore", 1);
                shell.SetValue("SeparatorAfter", 1);
                shell.SetValue("NeverDefault", string.Empty);
            }

            using var shell2 = key.EnsureSubKey($@"{path}\command");
            shell2?.SetValue("DelegateExecute", $"{typeof(ShellBatShellExtension).GUID:B}");
        }

        using var ShellBatFile = key.EnsureSubKey($@"Software\Classes\{FileExtension.ShellBatExtensionName}");
        ShellBatFile?.SetValue(null, Program.ProgId);

        using var progidKey = key.EnsureSubKey($@"Software\Classes\{Program.ProgId}");
        if (progidKey == null)
            return;

        progidKey.SetValue(null, Res.Document);
        using var defaultIconKey = progidKey.EnsureSubKey("DefaultIcon");
        defaultIconKey?.SetValue(null, $"{Environment.ProcessPath},-{ShellBatIconIndex}");

        using var shellKey = progidKey.EnsureSubKey(@"shell\open\command");
        shellKey?.SetValue(null, $"{Environment.ProcessPath} \"%1\"");

        using var ShellBatThemeFile = key.EnsureSubKey($@"Software\Classes\{FileExtension.ShellBatThemeExtensionName}");
        ShellBatThemeFile?.SetValue(null, Program.ProgId);

        using var themeDefaultIconKey = ShellBatThemeFile?.EnsureSubKey("DefaultIcon");
        themeDefaultIconKey?.SetValue(null, $"{Environment.ProcessPath},-{ShellBatThemeIconIndex}");
    }

    public static void UnregisterComObjects(RegistryKey key)
    {
        ArgumentNullException.ThrowIfNull(key);

        ComServer.Unregister<ShellBatShellExtension>(key);

        foreach (var path in GetOpenWithPaths())
        {
            key.DeleteSubKeyTree(path, false);
        }

        key.DeleteSubKeyTree($@"Software\Classes\{Program.ProgId}", false);
        key.DeleteSubKeyTree($@"Software\Classes\{FileExtension.ShellBatExtensionName}");
    }

    public static void ArrangeInstancesOnMonitor(WindowArrangement arrangement, int margin, DirectN.Extensions.Utilities.Monitor monitor, IReadOnlyList<IShellBatInstance> instances)
    {
        var rc = monitor.WorkingArea;
        margin = Math.Max(0, margin);

        var totalWidth = rc.Width - 2 * margin;
        var totalHeight = rc.Height - 2 * margin;

        switch (arrangement)
        {
            case WindowArrangement.Vertically:
                totalWidth -= margin * (instances.Count - 1);
                var left = rc.left + margin;
                var top = rc.top + margin;
                var width = totalWidth / instances.Count;
                for (var i = 0; i < instances.Count; i++)
                {
                    instances[i].SetWindowRect(new RECT(new POINT(left, top), new SIZE(width, totalHeight)));
                    left += width + margin;
                }
                break;

            case WindowArrangement.Horizontally:
                totalHeight -= margin * (instances.Count - 1);
                left = rc.left + margin;
                top = rc.top + margin;
                var height = totalHeight / instances.Count;
                for (var i = 0; i < instances.Count; i++)
                {
                    instances[i].SetWindowRect(new RECT(new POINT(left, top), new SIZE(totalWidth, height)));
                    top += height + margin;
                }
                break;

            case WindowArrangement.Tiled:
                var cols = (int)Math.Ceiling(Math.Sqrt(instances.Count));
                var rows = (int)Math.Ceiling((double)instances.Count / cols);
                totalWidth -= margin * (cols - 1);
                totalHeight -= margin * (rows - 1);
                left = rc.left + margin;
                top = rc.top + margin;
                width = totalWidth / cols;
                height = totalHeight / rows;
                for (var i = 0; i < instances.Count; i++)
                {
                    instances[i].SetWindowRect(new RECT(new POINT(left, top), new SIZE(width, height)));
                    if ((i + 1) % cols == 0)
                    {
                        left = rc.left + margin;
                        top += height + margin;
                    }
                    else
                    {
                        left += width + margin;
                    }
                }
                break;
        }
    }

    public static bool StartUpdate()
    {
        if (!Program.IsSingleFileDeployment)
            return false;

        _updateAuthId = Guid.NewGuid().ToString("N");

        var path = Path.Combine(Path.GetTempPath(), $"{Program.AppId}_{Guid.NewGuid():B}", Path.GetFileName(Environment.ProcessPath)!);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.Copy(Environment.ProcessPath!, path, true);
        var info = new ProcessStartInfo
        {
            FileName = path,
            Arguments = $"/{Program.UpdateArgumentName}:{_updateAuthId}",
        };

        //Application.TraceVerbose($"Starting updater instance '{path}' with auth id '{_updateAuthId}'.");
        var proc = Process.Start(info);
        if (proc == null)
        {
            Application.TraceError($"Cannot start updater instance '{path}'.");
            return false;
        }

        DirectN.Functions.AllowSetForegroundWindow((uint)proc.Id);

        // do the update: copy the current executable to a temp file and start it with the update argument
        // the restarted one will call this instance to quit and replace the executable
        QuitAllRemoteInstances((int[])[proc.Id]);

        //Application.TraceVerbose($"Updater instance started with process id {proc.Id}.");
        return true;
    }

    public void Quit()
    {
        Options &= ~ShellBatInstanceOptions.IsInHttpServerOnlyMode;
        _serverStopEvent.Set();
        foreach (var win in Application.AllWindows)
        {
            win.Close();
        }

        Application.AllExit();
    }

    public static void QuitAllRemoteInstances(IEnumerable<int>? exceptedProcessIds = null)
    {
        var remotes = ShellBatRemoteInstance.UpdateInstances();
        foreach (var other in remotes)
        {
            if (exceptedProcessIds != null && exceptedProcessIds.Contains(other.ProcessId))
                continue;

            // drain results to ensure completion
            _ = other.Quit().ToArray();
        }
    }

    public static void SwitchToInstance(int processId) => Current?.GetRemoteInstance(processId)?.SwitchTo();
    public static void QuitInstance(int processId) => Current?.GetRemoteInstance(processId)?.Quit().ToArray();
    public static void OpenNewInstance(OpenNewInstanceOptions? options = null)
    {
        options ??= new OpenNewInstanceOptions();
        var args = $"/{Program.FromArgumentName}:{SystemUtilities.CurrentProcess.Id}";
        if (!string.IsNullOrWhiteSpace(options.ParsingName))
        {
            args += $" /{Program.ParsingNameArgumentName}:\"{options.ParsingName}\"";
        }

        if (!string.IsNullOrWhiteSpace(options.ScreenDevicePath))
        {
            args += $" /{Program.ScreenDevicePathArgumentName}:\"{options.ScreenDevicePath}\"";
        }

        if (options.Rect != null)
        {
            args += $" /{Program.WindowRectArgumentName}:\"{options.Rect.Value}\"";
        }

        var psi = new ProcessStartInfo
        {
            FileName = Environment.ProcessPath,
            Arguments = args,
            UseShellExecute = true,
        };

        if (options.AsAdministrator)
        {
            psi.Verb = "runas";
        }
        else if (SystemUtilities.GetTokenElevationType() == TokenElevationType.Full)
        {
            ProcessUtilities.RestartAsNonAdmin();
            return;
        }

        var proc = Process.Start(psi);
        if (proc != null)
        {
            DirectN.Functions.AllowSetForegroundWindow((uint)proc.Id);
        }
    }

    // if we are here, there should be an available update, hence warnings
    public static async Task UpdateToLatest(string updateAuthId, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(updateAuthId);
        _updatingMode = true;
        LogInfo($"Starting update ...");

        var remotes = ShellBatRemoteInstance.UpdateInstances();
        if (remotes == null || remotes.Length == 0)
            throw new ShellBatException($"0015: No {Res.AppName} instance found, cannot update.");

        string? toBeUpdatedPath = null;
        foreach (var other in remotes)
        {
            toBeUpdatedPath = other.Update(updateAuthId);
            if (toBeUpdatedPath != null)
                break;
        }

        if (toBeUpdatedPath == null)
            throw new ShellBatException($"0016: No {Res.AppName} instance is ready for an update.");

        var fileName = Path.GetFileName(Environment.ProcessPath);
        var thisVersion = Version.Parse(Assembly.GetExecutingAssembly().GetFileVersion()!);
        var releases = await GitHubApi.ListReleasesAsync(cancellationToken);
        var last = releases.LastOrDefault(r => r.Assets != null && r.Assets.Any(a => a.Name.EqualsIgnoreCase(fileName)));
        if (last == null || thisVersion >= last.Version)
            throw new ShellBatException($"0017: No update available, current version is {thisVersion}, latest is {last?.Version.ToString() ?? "n/a"}.");

        var url = last.Assets.First(a => a.Name.EqualsIgnoreCase(fileName)).DownloadUrl ?? throw new ShellBatException($"0018: No download url found for release {last.Version}.");

        LogInfo($"Downloading release {last.Version} from '{url}' to '{toBeUpdatedPath}' ...");

        var temp = await GitHubApi.DownloadReleaseAsync(url, Path.GetFileName(toBeUpdatedPath), cancellationToken);

        if (cancellationToken.IsCancellationRequested)
        {
            LogWarning($"Update cancelled.");
            return;
        }

        MarkOfTheWeb.RemoveZone(temp, false);
        FileSystemUtilities.WrapSharingViolations(() => IOUtilities.FileOverwrite(temp, toBeUpdatedPath, true, false), maxRetryCount: 20, waitTime: 200);

        var screenDevicePath = WebScreen.All.FirstOrDefault(s => s.IsThis)?.DevicePath;
        var psi = new ProcessStartInfo
        {
            FileName = toBeUpdatedPath,
            Arguments = $"/{Program.UpdatedArgumentName} /{Program.ScreenDevicePathArgumentName}:\"{screenDevicePath}\""
        };

        if (cancellationToken.IsCancellationRequested)
        {
            LogWarning($"Update cancelled.");
            return;
        }

        LogInfo($"Starting updated instance '{toBeUpdatedPath}'.");
        var proc = Process.Start(psi);
        if (proc != null)
        {
            DirectN.Functions.AllowSetForegroundWindow((uint)proc.Id);
        }
    }

    public static void LogInfo(object? message = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null) => Log(TraceLevel.Info, message, methodName, filePath);
    public static void LogWarning(object? message = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null) => Log(TraceLevel.Warning, message, methodName, filePath);
    public static void LogVerbose(object? message = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null) => Log(TraceLevel.Verbose, message, methodName, filePath);
    public static void LogError(object? message = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null) => Log(TraceLevel.Error, message, methodName, filePath);
    public static void Log(TraceLevel level, object? message = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        if (!string.IsNullOrEmpty(methodName))
        {
            methodName += "|";
        }

        string? slevel = null;
        if (level == TraceLevel.Warning || level == TraceLevel.Error)
        {
            slevel = level.ToString() + "|";
        }

        var name = Thread.CurrentThread.Name.Nullify();
        if (name != null)
        {
            name = $"[{name}]";
        }

        var text = $"#{Environment.CurrentManagedThreadId}{name}|{methodName}{slevel}{message}";
        EventProvider.Default.WriteMessageEvent(text, (byte)level);
        if (level > Current.Settings.LogLevel)
            return;

        Current?.TraceToWriter(text, filePath);
    }
}
