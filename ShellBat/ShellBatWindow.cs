namespace ShellBat;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class ShellBatWindow : WebViewCompositionWindow
{
    private const string _assocHandlerId = "assochandler:";
    private static readonly ushort _clipboardCutFormatId1 = Clipboard.RegisterFormat($"{Program.AppId}{Guid.NewGuid():N}");
    private static readonly ushort _clipboardCutFormatId2 = Clipboard.RegisterFormat($"{Program.AppId}{Guid.NewGuid():N}");
    private static readonly ushort _clipboardDragFormatId = Clipboard.RegisterFormat($"{Program.AppId}{Guid.NewGuid():N}");

    private readonly DotNet _hostObject;
    private readonly Dictionary<string, WebTerminal> _terminals = [];
    private readonly ConcurrentDictionary<string, object?> _queries = new();
    private readonly List<PreviewViewer> _previewViewers = [];
    private PdfRenderer? _pdfRenderer;
    private EntryWatcher? _watcher;
    private SpriteVisual? _loadingVisual;
    private GitHubRelease? _availableUpdate;
    private readonly bool _created;
    private bool? _draggingIsUs;
    private bool _firstTimeNavigated = true;
    private RECT _captionRect; // in webbrowser coordinates
    private double _scale = 1;
    private float _direction = 1.1f;
    private int _terminalId;
    private Hsv _loadingHsv = new(0, 100, 100);
    private string[]? _draggingNames;
    private POINT _draggingPoint;
    private ushort _clipboardCutFormatId = _clipboardCutFormatId1;
    private int? _propertiesSourceProcessId;
    private int? _propertiesTargetProcessId;

    public ShellBatWindow()
        : base(
            Program.AppId,
            style: WINDOW_STYLE.WS_THICKFRAME,
            rect: ShellBatInstance.Current.Settings.GetInitialRect())
    {
        _created = true;
        ShellBatCommand.ProcessSettings(Handle);
        Cursor = DirectN.Extensions.Utilities.Cursor.AppStarting.Handle;
        DirectN.Functions.AddClipboardFormatListener(Handle);

        IsDropTarget = true;
        _hostObject = new DotNet(this)
        {
            ServerUrl = ShellBatInstance.Current.HttpServer.Url.ToString()
        };

        _hostObject.Event += (s, e) =>
        {
            HistoryEntry? historyEntry;
            KeyEventArgs? evt;
            string? name;
            string? parsingName;
            int processId;
            switch (e.Type)
            {
                case WebEventType.ReadyStateChange:
                    var state = e.Value.GetNullifiedProperty("state");
                    var msg = e.Value.GetNullifiedProperty("message");
                    //Application.TraceVerbose($"WebView ReadyStateChange: state={state} dir:{System.Environment.CurrentDirectory} msg:{msg}");
                    if (state.EqualsIgnoreCase("complete"))
                    {
                        ApplyTheme(ShellBatInstance.Current.Theme);

                        // if we're opening a dead entry, go to desktop instead
                        if ((VIRTUAL_KEY.VK_SHIFT.IsPressed() && VIRTUAL_KEY.VK_CONTROL.IsPressed()))
                        {
                            Navigate(string.Empty, false);
                            return;
                        }

                        // we could end up because of a F5
                        if (CurrentEntry is not null)
                        {
                            Navigate(CurrentEntry, false, null, true);
                            return;
                        }

                        // determine initial parsing name
                        var historyName = History.Current.Entries.FirstOrDefault()?.ParsingName;
                        var firstArg = ShellBatInstance.GetFirstArgumentFilePath();
                        var initialParsingName = firstArg ??
                            CommandLine.Current.GetNullifiedArgument(Program.ParsingNameArgumentName) ??
                            historyName;

                        //Application.TraceVerbose($"Initial parsing name: '{initialParsingName}'");
                        if (!Navigate(initialParsingName, false))
                        {
                            if (initialParsingName == historyName)
                            {
                                ShowError(new() { Text = string.Format(Res.CannotNavigate, initialParsingName), Timer = 3000 });
                            }

                            Navigate(Entry.Desktop, false);
                        }
                    }
                    break;

                case WebEventType.OpenNewInstance:
                case WebEventType.OpenNewInstanceAdministrator:
                case WebEventType.OpenNewInstanceOnScreen:
                    parsingName = e.Value.GetNullifiedProperty("parsingName") ?? CurrentEntry?.ParsingName;
                    var devicePath = e.Value.GetNullifiedProperty("devicePath");
                    bool asAdmin;
                    if (VIRTUAL_KEY.VK_SHIFT.IsPressed())
                    {
                        asAdmin = true;
                    }
                    else
                    {
                        asAdmin = e.Type == WebEventType.OpenNewInstanceAdministrator;
                    }

                    ShellBatInstance.OpenNewInstance(new OpenNewInstanceOptions { ParsingName = parsingName, ScreenDevicePath = devicePath, AsAdministrator = asAdmin });
                    break;

                case WebEventType.QuitAllInstances:
                    QuitAllInstances();
                    break;

                case WebEventType.ArrangeInstances:
                    var arrangement = e.Value.GetProperty<WindowArrangement>("arrangement");
                    var allScreens = e.Value.GetProperty<bool>("allScreens");
                    var margin = e.Value.GetProperty("margin", ShellBatInstance.Current.Settings.DockPadding);
                    ArrangeInstances(arrangement, margin, allScreens);
                    break;

                case WebEventType.SwitchToInstance:
                    processId = e.Value.GetProperty<int>("processId");
                    ShellBatInstance.SwitchToInstance(processId);
                    break;

                case WebEventType.QuitInstance:
                    processId = e.Value.GetProperty<int>("processId");
                    ShellBatInstance.QuitInstance(processId);
                    break;

                case WebEventType.CaptionSizeChanged:
                    _captionRect = e.Value.AsDOMRect();
                    break;

                case WebEventType.VisualViewportChanged:
                    _scale = e.Value.GetProperty<double>("scale", -1);
                    foreach (var viewer in _previewViewers)
                    {
                        _ = viewer.Refresh();
                    }
                    break;

                case WebEventType.WindowNotOpen:
                    name = e.Value.GetNullifiedProperty("name");
                    if (name.EqualsIgnoreCase(WebPropertiesWindow.WindowId)) // remote view only works on properties window, not on others
                    {
                        parsingName = e.Value.GetNullifiedProperty("parsingName");
                        if (parsingName != null && _propertiesTargetProcessId != null)
                        {
                            // try to continue with other instance
                            var instance = ShellBatInstance.Current.GetRemoteInstance(_propertiesTargetProcessId.Value);
                            if (instance == null)
                            {
                                _propertiesTargetProcessId = null;
                                break;
                            }

                            var hr = instance.ContinueShowProperties(parsingName);
                            if (hr.IsError)
                            {
                                _propertiesTargetProcessId = null;
                            }
                        }
                    }
                    break;

                case WebEventType.RemoveDeletedItemsFromFavorites:
                    if (Settings.Current.RemoveDeletedFavorites())
                    {
                        Settings.Current.SerializeToConfiguration();
                    }
                    break;

                case WebEventType.RemoveDeletedItemsFromHistory:
                    if (History.Current.RemoveDeletedEntries())
                    {
                        History.Current.SerializeToConfiguration();
                    }
                    break;

                case WebEventType.UpdateNow:
                    _ = CheckForUpdates(true);
                    break;

                case WebEventType.OpenLogsFolder:
                    if (IOUtilities.PathIsDirectory(ShellBatInstance.LogsDirectoryPath))
                    {
                        WindowsExtensions.Open(ShellBatInstance.LogsDirectoryPath);
                    }
                    break;

                case WebEventType.SetGlobalEvents:
                    var type = e.Value.GetNullifiedProperty("type");
                    switch (type)
                    {
                        case GlobalEventTypes.NewWindowReady:
                        case GlobalEventTypes.WindowClose:
                        case GlobalEventTypes.WindowMove:
                        case GlobalEventTypes.WindowResize:
                        case GlobalEventTypes.WindowUpdate:
                            _ = OnChildWindowsChanged();
                            break;
                    }
                    break;

                case WebEventType.Close:
                    Close();
                    break;

                case WebEventType.Minimize:
                    Show(SHOW_WINDOW_CMD.SW_SHOWMINIMIZED);
                    break;

                case WebEventType.MaximizeRestore:
                    Show(IsZoomed ? SHOW_WINDOW_CMD.SW_RESTORE : SHOW_WINDOW_CMD.SW_SHOWMAXIMIZED);
                    break;

                case WebEventType.KeyDown:
                    evt = e.Value.AsKeyEventArgs();
                    //Application.TraceInfo($"WebView KeyDown: {evt}");
                    if (evt != null)
                    {
                        OnKeyDown(this, evt);
                    }
                    break;

                case WebEventType.KeyUp:
                    evt = e.Value.AsKeyEventArgs();
                    if (evt != null)
                    {
                        OnKeyUp(this, evt);
                    }
                    break;

                case WebEventType.ConsoleLog:
                    if (e.Value.GetObjectProperty("arguments") is object[] arguments)
                    {
                        foreach (var argument in arguments)
                        {
                            Application.TraceVerbose("ConsoleLog: " + argument);
                        }
                    }
                    break;

                case WebEventType.Log:
                    if (e.Value.GetObjectProperty("arguments") is object[] logArguments)
                    {
                        foreach (var argument in logArguments)
                        {
                            ShellBatInstance.LogVerbose("WebViewLog: " + argument);
                        }
                    }
                    break;

                case WebEventType.MoveHistoryBack:
                    historyEntry = History.Current.MoveBack();
                    if (historyEntry != null)
                    {
                        Navigate(historyEntry.ParsingName, true);
                    }
                    break;

                case WebEventType.MoveHistoryForward:
                    parsingName = CurrentEntry?.ParsingName;
                    historyEntry = History.Current.MoveForward();
                    if (historyEntry != null)
                    {
                        Navigate(historyEntry.ParsingName, true);
                    }
                    break;

                case WebEventType.RestartAsAdministrator:
                    RestartAsAdministrator();
                    break;

                case WebEventType.EntryDoubleClicked:
                    parsingName = e.Value.GetNullifiedProperty("parsingName");
                    Run(parsingName, false);
                    break;

                case WebEventType.ExportAsCsv:
                    var entry = CurrentEntry;
                    if (entry is null)
                        return;

                    name = IOUtilities.NameToValidFileName(entry.DisplayName);
                    var csvPath = Path.Combine(Path.GetTempPath(), $"{name}.csv");
                    entry.ExportAsCsv(csvPath);
                    WindowsExtensions.Open(csvPath);
                    break;

                case WebEventType.ExportExtensionsAsCsv:
                    var extensionsCsvPath = Path.Combine(Path.GetTempPath(), $"{Program.AppId}_Extensions.csv");
                    FileExtension.ExportAsCsv(extensionsCsvPath);
                    WindowsExtensions.Open(extensionsCsvPath);
                    break;

                case WebEventType.DisposeChildWindows:
                    foreach (var child in ChildWindows)
                    {
                        if (child.ClassName.StartsWith(Program.AppId))
                        {
                            child.Dispose();
                        }
                    }
                    _propertiesSourceProcessId = null;
                    break;

                case WebEventType.OnSearchDoubleClicked:
                    parsingName = e.Value.GetNullifiedProperty("parsingName");
                    var searchType = e.Value.GetProperty<SearchType>("type");
                    if (parsingName != null)
                    {
                        SelectEntry(parsingName);
                        switch (searchType)
                        {
                            case SearchType.FindStrings:
                                var position = e.Value.GetProperty<long>("position");
                                ShowProperties(parsingName, new WebWindowShowOptions
                                {
                                    ForceOpen = true,
                                    ViewerId = "Binary",
                                    ViewerOptions = new Dictionary<string, object?> { ["position"] = position, }
                                });
                                break;

                            case SearchType.WindowsSearch:
                                break;
                        }
                    }
                    break;
            }
        };

        _hostObject.Error += (s, e) =>
        {
            ShellBatInstance.LogError($"WebView Error: {e.Message} (source: {e.Source}, line: {e.Line}, col: {e.Column} stack: {e.StackTrace})");
            Application.AddError(e.BuildException());
            Application.ShowFatalError(Handle);
        };

        _hostObject.Navigate += (s, e) =>
        {
            if (_firstTimeNavigated)
            {
                _firstTimeNavigated = false;
                RestoreChildWindows();

                if (CommandLine.Current.HasArgument(Program.UpdatedArgumentName))
                {
                    var version = Assembly.GetExecutingAssembly().GetFileVersion();
                    ShowSuccess(new() { Text = string.Format(Res.UpdatedSuccessfully, version), Timer = 5000 });
                }
            }

            if (CurrentEntry == e.Entry)
                return;

            CurrentEntry = e.Entry;
            var entry = e.Entry;

            if (entry != null)
            {
                Text = $"{Program.AppId} - {entry.FullDisplayName}";
            }
            else
            {
                Text = Program.AppId;
            }

            if (entry is null)
            {
                if (_watcher != null)
                {
                    _watcher.Event -= OnEntryEvent;
                    _watcher.Dispose();
                }
                return;
            }

            if (entry.IsFolder)
            {
                _ = History.Current.AddEntry(entry);

                if (_watcher != null)
                {
                    _watcher.Event -= OnEntryEvent;
                    _watcher.Dispose();
                }

                if (!entry.IsZipEntry)
                {
                    _watcher = new EntryWatcher(entry);
                    _watcher.Event += OnEntryEvent;
                    try
                    {
                        _watcher.Run();
                    }
                    catch (Exception ex)
                    {
                        ShellBatInstance.LogError($"Failed to start watcher for '{entry.ParsingName}': {ex}");
                        _watcher.Event -= OnEntryEvent;
                        _watcher.Dispose();
                        _watcher = null;
                        ShowError(new() { Text = string.Format(Res.EntryWatcherError, entry.ParsingName), Timer = 3000 });
                    }
                }

                foreach (var kv in _terminals)
                {
                    if (kv.Value.GetTerminalSync())
                    {
                        kv.Value.ChangeDirectoryInWebView(entry);
                    }
                }
            }
        };
    }

    public Entry? CurrentEntry { get; private set; }
    public double Scale => _scale;
    public Task<D2D_RECT_F> GetEntriesRect(bool throwOnError = true) => ExecuteScript("getEntriesRect();", JsonSourceStructGenerationContext.Default.D2D_RECT_F, throwOnError);

    public void RefreshEntries(string? parsingName) => RunTaskOnUIThread(() => ExecuteScript($"refreshEntries(\"{HttpUtility.JavaScriptStringEncode(parsingName)}\");"));
    private void OnEntryEvent(object? sender, EntryWatcherEventArgs e)
    {
        //Application.TraceInfo($"event: {e}");
        if (!ShellBatInstance.Current.Settings.NotifyOnFolderEvents)
            return;

        if (string.IsNullOrWhiteSpace(e.Entry?.DisplayName)) // desktop
            return;

        var reason = e.ToLocalizedString();
        RunTaskOnUIThread(() => ExecuteScript($"refreshEntries(null, \"{HttpUtility.JavaScriptStringEncode(reason)}\");"));
    }

    public virtual void UpdateRemoteInstances(IReadOnlyList<ShellBatRemoteInstance> instances)
    {
        ArgumentNullException.ThrowIfNull(instances);
        RunTaskOnUIThread(() => ExecuteScript("updateInstances();"));
    }

    public virtual void ClearCaches(CacheTypes types)
    {
        if (types.HasFlag(CacheTypes.HttpServer))
        {
            IOUtilities.DirectoryDeleteFiles(Settings.HttpServerCachePath, true, false);
        }

        if (types.HasFlag(CacheTypes.WebView2))
        {
            RunTaskOnUIThread(() => ClearBrowsingDataAll());
        }
    }

    public virtual void ShowPropertyGrid(PropertyGridType type) => ExecuteScript($"showPropertyGrid({(int)type});");
    public virtual void ShowProperties(string? parameters, WebWindowShowOptions? options = null, ShellBatInstanceCommand? command = null)
    {
        LockProperties(command);
        options ??= new WebWindowShowOptions();
        string id;
        if (options.NewWindow)
        {
            id = WebPropertiesWindow.GetNewId();
        }
        else
        {
            id = WebPropertiesWindow.WindowId;
        }
        ShowProperties(id, parameters, options);
    }

    private void ShowProperties(string? id, string? parameters, WebWindowShowOptions? options) =>
        ExecuteScript($"showWindow(\"{id}\", \"{HttpUtility.JavaScriptStringEncode(parameters)}\", {JsonSerializer.Serialize(options, JsonSourceGenerationContext.Default.WebWindowShowOptions)});");

    public virtual void ShowToast(string title, int timer = 3000)
    {
        if (string.IsNullOrWhiteSpace(title))
            return;

        RunTaskOnUIThread(() => ExecuteScript($"showToast(\"{HttpUtility.JavaScriptStringEncode(title)}\",{timer});"));
    }

    protected virtual void ShowAlert(WebAlert alert)
    {
        ArgumentNullException.ThrowIfNull(alert);
        RunTaskOnUIThread(() => ExecuteScript($"showAlert({JsonSerializer.Serialize(alert, JsonSourceGenerationContext.Default.WebAlert)});"));
    }

    public virtual void ShowError(WebAlert alert)
    {
        ArgumentNullException.ThrowIfNull(alert);
        alert.Icon = "error";
        ShowAlert(alert);
    }

    public virtual void ShowSuccess(WebAlert alert)
    {
        ArgumentNullException.ThrowIfNull(alert);
        alert.Icon = "success";
        ShowAlert(alert);
    }

    public virtual void ShowInfo(WebAlert alert)
    {
        ArgumentNullException.ThrowIfNull(alert);
        alert.Icon = "info";
        ShowAlert(alert);
    }

    public virtual void ShowQuestion(WebAlert alert)
    {
        ArgumentNullException.ThrowIfNull(alert);
        alert.Icon = "question";
        alert.ShowCancelButton = true;
        alert.ShowConfirmButton = true;
        alert.Timer = 30000;
        alert.TimerProgressBar = true;
        ShowAlert(alert);
    }

    public virtual WebContextMenu? GetContextMenu(IEnumerable<string> parsingNames)
    {
        ArgumentNullException.ThrowIfNull(parsingNames);
        //Application.TraceVerbose($"GetContextMenu: {string.Join(", ", parsingNames)}");
        var entries = parsingNames.Select(n => Entry.Get(null, n, ShellItemParsingOptions.DontThrowOnError)).WhereNotNull().ToArray();
        if (entries.Length == 0)
            return null;

        IShellBatInstance[]? instances = null;
        var currentEntry = CurrentEntry;
        var menu = new WebContextMenu(this, entries);
        WebMenuItem item;
        if (entries.Length == 1)
        {
            var entry = entries[0];

            item = new WebMenuItem { Html = entry.DisplayName, IsDisabled = true, ClassName = "title" };
            menu.MenuItems.Add(item);

            if (entry.IsShortcut)
            {
                item.Html += " (" + Res.Link + ")";
            }

            if (entry.IsRecycleBin)
            {
                AddRecycleBinCommands(menu.MenuItems);
                menu.MenuItems.Add(new WebMenuItem { IsSeparator = true });
            }

            item = new WebMenuItem { Id = nameof(Res.View), Html = Res.View, Icon = "fa-solid fa-binoculars", };
            menu.MenuItems.Add(item);

            instances ??= [.. ShellBatInstance.Current.All];
            if (instances.Length > 1)
            {
                item = new WebMenuItem { Id = nameof(Res.ViewIn), Html = Res.ViewIn, };
                menu.MenuItems.Add(item);
                foreach (var instance in instances.Where(i => !i.IsThis))
                {
                    var childItem = new WebMenuItem
                    {
                        Id = $"{nameof(Res.ViewIn)}:{instance.ProcessId}",
                        Html = instance.OtherDisplayName,
                    };
                    item.Items.Add(childItem);
                }
            }

            if (entry.VirtualDiskImagePath != null)
            {
                item = new WebMenuItem { Id = nameof(Res.Eject), Html = Res.Eject, Icon = "fa-solid fa-eject", };
                menu.MenuItems.Add(item);
            }

            if (entry.Extension.IsDiskImage)
            {
                item = new WebMenuItem { Id = nameof(Res.Mount), Html = Res.Mount, Icon = "fa-solid fa-compact-disc", };
                menu.MenuItems.Add(item);

                item = new WebMenuItem { Id = nameof(Res.MountInNewInstance), Html = Res.MountInNewInstance, Icon = "fa-solid fa-compact-disc", };
                menu.MenuItems.Add(item);

                foreach (var screen in WebScreen.All)
                {
                    if (screen.IsThis)
                        continue;

                    item = new WebMenuItem { Id = $"{nameof(Res.MountInNewInstanceScreen)}:{screen.DevicePath}", Html = string.Format(Res.MountInNewInstanceScreen, screen.DisplayName), Icon = "fa-solid fa-compact-disc" };
                    menu.MenuItems.Add(item);
                }
            }

            if (entry.IsShortcut)
            {
                var target = entry.GetShortcutTargetEntry();
                if (target is not null)
                {
                    item = new WebMenuItem { Id = nameof(Res.OpenTargetFolder), Html = Res.OpenTargetFolder, Icon = "fa-solid fa-up-right-from-square", };
                    menu.MenuItems.Add(item);
                }
            }

            item = new WebMenuItem { Id = nameof(Res.OpenWithShell), Html = Res.OpenWithShell };
            menu.MenuItems.Add(item);

            if (entry.CanBeOpenedAsFolder)
            {
                item = new WebMenuItem { Id = nameof(Res.OpenAsFolder), Html = Res.OpenAsFolder };
                menu.MenuItems.Add(item);

                item = new WebMenuItem { Id = nameof(Res.OpenInNewInstance), Html = Res.OpenInNewInstance };
                menu.MenuItems.Add(item);

                item = new WebMenuItem { Id = nameof(Res.OpenInNewInstanceAdministrator), Html = Res.OpenInNewInstanceAdministrator };
                menu.MenuItems.Add(item);

                foreach (var screen in WebScreen.All)
                {
                    if (screen.IsThis)
                        continue;

                    item = new WebMenuItem { Id = $"{nameof(Res.OpenInNewInstanceScreen)}:{screen.DevicePath}", Html = string.Format(Res.OpenInNewInstanceScreen, screen.DisplayName) };
                    menu.MenuItems.Add(item);
                }
            }

            if (entry.IsConsoleApplication())
            {
                item = new WebMenuItem { Id = nameof(Res.OpenWithNewTerminal), Html = Res.OpenWithNewTerminal };
                menu.MenuItems.Add(item);

                item.Items = [.. Settings.Current.GetTerminals().Select(t => new WebMenuItem
                {
                    Id = $"{nameof(Res.OpenWithNewTerminal)}:{t.Key}",
                    Html = t.DisplayName,
                    Icon = t.FinalIcon,
                })];
            }

            var assocHandlers = entry.Extension.GetAssocHandlers();
            try
            {
                if (assocHandlers.Count > 0)
                {
                    item = new WebMenuItem { Id = nameof(Res.OpenWith), Html = Res.OpenWith };
                    menu.MenuItems.Add(item);

                    foreach (var handler in assocHandlers)
                    {
                        var name = handler.UIName ?? handler.Name ?? handler.ProgId ?? handler.AppUserModelId ?? "???";
                        var handlerItem = new WebMenuItem
                        {
                            Id = $"{_assocHandlerId}{handler.Key}",
                            Html = $"<div class=openwith><img src='{_hostObject.ServerUrl}{UrlType.AssocHandler}/{ShellBatHttpLocalServer.EscapeUrl(entry.Extension.Name)}/{ShellBatHttpLocalServer.EscapeUrl(handler.Key)}?shil=1'>&nbsp;{HttpUtility.HtmlEncode(name)}</div>",
                        };
                        item.Items.Add(handlerItem);
                    }
                }
            }
            finally
            {
                assocHandlers.Dispose();
            }

            if (currentEntry?.IsRecycleBin == true && !entry.IsRecycleBin)
            {
                item = new WebMenuItem { Id = nameof(Res.Restore), Html = Res.Restore, Icon = "fa-solid fa-trash-arrow-up", };
                menu.MenuItems.Add(item);
            }

            menu.MenuItems.Add(new WebMenuItem { IsSeparator = true });
        }
        else
        {
            item = new WebMenuItem { Html = string.Format(Res.ContextMenuItems, entries.Length), IsDisabled = true, ClassName = "title" };
            menu.MenuItems.Add(item);

            if (entries.Length > 1 && currentEntry?.IsRecycleBin == true)
            {
                item = new WebMenuItem { Id = nameof(Res.Restore), Html = Res.Restore, Icon = "fa-solid fa-trash-arrow-up", };
                menu.MenuItems.Add(item);
                menu.MenuItems.Add(new WebMenuItem { IsSeparator = true });
            }
        }

        void addCopyMoveInstances(string cid, WebMenuItem menuItem)
        {
            foreach (var instance in instances.Where(i => !i.IsThis))
            {
                var item = new WebMenuItem
                {
                    Id = $"{cid}:{instance.ProcessId}",
                    Html = instance.OtherDisplayName,
                };
                menuItem.Items.Add(item);
            }
        }

        if (entries.Any(e => e.CanCopy))
        {
            item = new WebMenuItem { Id = nameof(Res.Copy), Html = Res.Copy, Icon = "fa-solid fa-copy", };
            menu.MenuItems.Add(item);

            instances ??= [.. ShellBatInstance.Current.All];
            if (instances.Length > 1)
            {
                item = new WebMenuItem { Id = nameof(Res.CopyTo), Html = Res.CopyTo, };
                menu.MenuItems.Add(item);
                addCopyMoveInstances(nameof(Res.CopyTo), item);
            }
        }

        if (entries.Any(e => e.CanMove))
        {
            instances ??= [.. ShellBatInstance.Current.All];
            if (instances.Length > 1)
            {
                item = new WebMenuItem { Id = nameof(Res.MoveTo), Html = Res.MoveTo, };
                menu.MenuItems.Add(item);
                addCopyMoveInstances(nameof(Res.MoveTo), item);
            }
        }

        item = new WebMenuItem { Id = nameof(Res.Zip), Html = Res.Zip, Icon = "fa-solid fa-file-zipper", };
        menu.MenuItems.Add(item);

        if (entries.Length == 1 && (entries[0].Extension.CanBeDecompressed || entries[0].Extension.IsCompressed))
        {
            item = new WebMenuItem { Id = nameof(Res.ExtractTo), Html = Res.ExtractTo, };
            menu.MenuItems.Add(item);
        }

        if (entries.Length == 1)
        {
            var dataObject = Clipboard.GetDataObject(false);
            if (dataObject.IsFilePasteEnabled())
            {
                menu.MenuItems.Add(new WebMenuItem { IsSeparator = true });
                item = new WebMenuItem { Id = nameof(Res.Paste), Html = Res.Paste, Icon = "fa-solid fa-paste", };
                menu.MenuItems.Add(item);
                item = new WebMenuItem { Id = nameof(Res.PasteShortcut), Html = Res.PasteShortcut, Icon = "fa-solid fa-link", };
                menu.MenuItems.Add(item);
            }
        }

        if (entries.Any(e => e.CanDelete))
        {
            menu.MenuItems.Add(new WebMenuItem { IsSeparator = true });
            item = new WebMenuItem { Id = nameof(Res.Recycle), Html = Res.Recycle, Icon = "fa-solid fa-recycle", };
            menu.MenuItems.Add(item);
            item = new WebMenuItem { Id = nameof(Res.Delete), Html = Res.Delete, Icon = "fa-solid fa-circle-xmark", };
            menu.MenuItems.Add(item);
            item = new WebMenuItem { Id = nameof(Res.Cut), Html = Res.Cut, Icon = "fa-solid fa-scissors", };
            menu.MenuItems.Add(item);
        }

        if (entries.Length == 1 && entries[0].CanRename)
        {
            item = new WebMenuItem { Id = nameof(Res.Rename), Html = Res.Rename, Icon = "fa-solid fa-font", };
            menu.MenuItems.Add(item);
        }

        menu.MenuItems.Add(new WebMenuItem { IsSeparator = true });
        if (entries.Length == 1 &&
            entries[0].FileSystemPath != null &&
            IOUtilities.PathIsFile(entries[0].FileSystemPath!))
        {
            var zone = MarkOfTheWeb.MapUrlToZone(entries[0].FileSystemPath!, false);
            if (zone != URLZONE.URLZONE_LOCAL_MACHINE)
            {
                item = new WebMenuItem { Id = nameof(Res.Unblock), Html = Res.Unblock, Icon = "fa-solid fa-unlock", };
                menu.MenuItems.Add(item);
            }
            else if (VIRTUAL_KEY.VK_SHIFT.IsPressed())
            {
                item = new WebMenuItem { Id = nameof(Res.Block), Html = Res.Block, Icon = "fa-solid fa-lock", };
                menu.MenuItems.Add(item);
            }
        }

        item = new WebMenuItem { Id = nameof(Res.ShellContextMenu), Html = Res.ShellContextMenu, Icon = "fa-brands fa-windows", };
        menu.MenuItems.Add(item);

        WebContextMenu.RemoveSpuriousSeparators(menu.MenuItems);
        return menu;
    }

    private void ExecuteActionInternal(IReadOnlyList<Entry> entries, string actionId)
    {
        ArgumentNullException.ThrowIfNull(entries);
        ArgumentNullException.ThrowIfNull(actionId);
        //Application.TraceInfo($"{actionId}");

        switch (actionId)
        {
            case nameof(Res.EmptyRecycleBin):
                {
                    using var rb = new RecycleBin();
                    using var fo = new FileOperation();
                    rb.PurgeAll(fo);
                    fo.PerformOperations();
                }
                break;

            case nameof(Res.RestoreAllRecycleBinItems):
                {
                    using var rb = new RecycleBin();
                    rb.RestoreAll();
                }
                break;

            case nameof(Res.View):
                if (entries.Count == 1)
                {
                    var entry = entries[0];
                    if (entry is not null)
                    {
                        var options = new WebWindowShowOptions { ForceOpen = true };
                        if (VIRTUAL_KEY.VK_SHIFT.IsPressed())
                        {
                            options.NewWindow = true;
                        }

                        ShowProperties(entry.ParsingName, options);
                    }
                }
                break;

            case nameof(Res.OpenWithShell):
                if (entries.Count == 1)
                {
                    var entry = entries[0];
                    if (entry is not null)
                    {
                        WindowsExtensions.Open(entry.ParsingName);
                    }
                }
                break;

            case nameof(Res.OpenAsFolder):
                if (entries.Count == 1)
                {
                    var entry = entries[0];
                    if (entry.CanBeOpenedAsFolder)
                    {
                        Navigate(entry, false);
                    }
                }
                break;

            case nameof(Res.OpenInNewInstanceAdministrator):
            case nameof(Res.OpenInNewInstance):
                openNewInstance();
                break;

            case nameof(Res.OpenTargetFolder):
                if (entries.Count == 1)
                {
                    var entry = entries[0];
                    var target = entry.GetShortcutTargetEntry();
                    if (target is not null && target.Parent is not null)
                    {
                        Navigate(target.Parent, false);
                    }
                }
                break;

            case nameof(Res.Unblock):
                foreach (var entry in entries.Select(e => e.FileSystemPath).WhereNotNull())
                {
                    MarkOfTheWeb.RemoveZone(entry, false);
                }
                break;

            case nameof(Res.Block):
                foreach (var entry in entries.Select(e => e.FileSystemPath).WhereNotNull())
                {
                    MarkOfTheWeb.SetZone(entry, URLZONE.URLZONE_INTERNET, false);
                }
                break;

            case nameof(Res.Rename):
                ExecuteScript($"renameEntry();");
                break;

            case nameof(Res.Zip):
                if (CurrentEntry != null)
                {
                    var zipItems = entries.Select(e => Entry.GetItem(e.ParsingName, ShellItemParsingOptions.DontThrowOnError)).WhereNotNull().ToArray();
                    if (zipItems.Length > 0)
                    {
                        try
                        {
                            var zipEntry = CurrentEntry.CreateNewItem(CurrentEntry.DisplayName + ".zip");
                            if (zipEntry != null)
                            {
                                zipEntry.CopyTo(zipItems, null, false);
                                RefreshEntries(zipEntry.ParsingName);
                            }
                        }
                        finally
                        {
                            zipItems.Dispose();
                        }
                    }
                }
                break;

            case nameof(Res.Copy):
            case nameof(Res.Cut):
                CopyOrCut(actionId, entries);
                break;

            case nameof(Res.Paste):
            case nameof(Res.PasteShortcut):
                PasteOrPasteAsShortcut(actionId);
                break;

            case nameof(Res.Mount):
            case nameof(Res.MountInNewInstance):
                mount();
                break;

            case nameof(Res.Eject):
                if (entries.Count == 1)
                {
                    var entry = entries[0];
                    if (entry.VirtualDiskImagePath != null)
                    {
                        // the eject verb (like "mount") doesn't work, this is painful
                        // so we use this hack
                        var hr = IOUtilities.EjectVirtualDisk(entry.ParsingName, Handle, out var isVirtual);
                        if (hr.IsSuccess)
                        {
                            // navigate back to the disk image entry's folder
                            var containingFolder = Entry.Get(null, entry.VirtualDiskImagePath, ShellItemParsingOptions.DontThrowOnError)?.Parent;
                            if (containingFolder != null)
                            {
                                Navigate(containingFolder, false, entry.VirtualDiskImagePath);
                            }
                        }
                    }
                }
                break;

            case nameof(Res.Restore):
                {
                    var restoreItems = entries.Select(e => Entry.GetItem(e.ParsingName, ShellItemParsingOptions.DontThrowOnError)).WhereNotNull().ToArray();
                    try
                    {
                        using var rb = new RecycleBin();
                        using var fo = new FileOperation();
                        rb.RestoreItems(restoreItems, fo);
                        fo.PerformOperations();
                    }
                    finally
                    {
                        restoreItems.Dispose();
                    }

                    // for some reason this doesn't work automatically if the reycle bin is not opened
                    RefreshEntries(null);
                }
                break;

            case nameof(Res.Recycle):
            case nameof(Res.Delete):
                if (actionId == nameof(Res.Delete) && VIRTUAL_KEY.VK_SHIFT.IsPressed())
                {
                    // bypass recycle bin
                    foreach (var entry in entries.Where(e => e.CanDelete))
                    {
                        entry.Delete();
                    }
                    return;
                }

                var items = entries.Where(e => e.CanDelete).Select(e => Entry.GetItem(e.ParsingName, ShellItemParsingOptions.DontThrowOnError)).WhereNotNull().ToArray();
                try
                {
                    if (items.Length > 0)
                    {
                        using var fo = new FileOperation();
                        if (actionId == nameof(Res.Recycle))
                        {
                            fo.SetOperationFlags(FILEOPERATION_FLAGS.FOFX_RECYCLEONDELETE);
                        }

                        fo.DeleteItems(items.Select(i => i.NativeObject));
                        fo.PerformOperations();
                    }
                }
                finally
                {
                    items.Dispose();
                }

                // for some reason this doesn't work automatically if the reycle bin is not opened
                RefreshEntries(null);
                break;

            case nameof(Res.ShellContextMenu):
                var list = new List<ItemIdList>();
                try
                {
                    foreach (var entry in entries)
                    {
                        var idl = entry.GetIdList(false);
                        if (idl is not null)
                        {
                            list.Add(idl);
                        }
                    }

                    using var site = new Site(this);
                    ShellItem.ShowContextMenu(list, site, flags: CMF.CMF_EXPLORE | CMF.CMF_EXTENDEDVERBS | CMF.CMF_CANRENAME);
                }
                finally
                {
                    list.Dispose();
                }
                break;

            default:
                if (entries.Count == 1 && actionId.StartsWith(_assocHandlerId, StringComparison.OrdinalIgnoreCase))
                {
                    var handlerKey = actionId[_assocHandlerId.Length..].Nullify();
                    if (handlerKey != null)
                    {
                        var entry = entries[0];
                        var assocHandler = entry.Extension.GetAssocHandler(h => h.Key.EqualsIgnoreCase(handlerKey));
                        if (assocHandler != null)
                        {
                            var dataObject = entry.CreateDataObject(false);
                            if (dataObject != null)
                            {
                                assocHandler.Invoke(dataObject.Object);
                            }
                        }
                    }
                    break;
                }

                if (entries.Count == 1 && actionId.StartsWith($"{nameof(Res.OpenWithNewTerminal)}:", StringComparison.OrdinalIgnoreCase))
                {
                    var terminalKey = actionId[$"{nameof(Res.OpenWithNewTerminal)}:".Length..].Nullify();
                    if (terminalKey != null)
                    {
                        var entry = entries[0];
                        var terminal = Settings.Current.GetTerminals().FirstOrDefault(t => t.Key.EqualsIgnoreCase(terminalKey));
                        if (terminal != null)
                        {
                            RunTerminal(terminalKey, new WebWindowShowOptions { CommandLine = entry.ParsingName });
                        }
                    }
                    break;
                }

                if (actionId.StartsWith($"{nameof(Res.OpenInNewInstanceScreen)}:", StringComparison.OrdinalIgnoreCase))
                {
                    var screenDevicePath = actionId[$"{nameof(Res.OpenInNewInstanceScreen)}:".Length..].Nullify();
                    if (screenDevicePath != null)
                    {
                        openNewInstance(screenDevicePath);
                    }
                    break;
                }

                if (actionId.StartsWith($"{nameof(Res.MountInNewInstanceScreen)}:", StringComparison.OrdinalIgnoreCase))
                {
                    var screenDevicePath = actionId[$"{nameof(Res.MountInNewInstanceScreen)}:".Length..].Nullify();
                    mount(screenDevicePath);
                    break;
                }

                if (entries.Count == 1 && actionId.StartsWith($"{nameof(Res.ViewIn)}:", StringComparison.OrdinalIgnoreCase))
                {
                    var parts = actionId.Split(':');
                    if (parts.Length == 2 && int.TryParse(parts[1], out var processId))
                    {
                        var targetInstance = ShellBatInstance.Current.GetRemoteInstance(processId);
                        if (targetInstance != null)
                        {
                            var hr = targetInstance.ShowProperties(entries[0].ParsingName);
                            if (hr.IsSuccess)
                            {
                                _propertiesTargetProcessId = processId;
                            }
                        }
                    }
                    break;
                }

                if (actionId.StartsWith($"{nameof(Res.CopyTo)}:", StringComparison.OrdinalIgnoreCase) ||
                    actionId.StartsWith($"{nameof(Res.MoveTo)}:", StringComparison.OrdinalIgnoreCase))
                {
                    var parts = actionId.Split(':');
                    if (parts.Length == 2 && int.TryParse(parts[1], out var processId))
                    {
                        var targetEntry = (ShellBatInstance.Current.GetRemoteInstance(processId) as IShellBatInstance)?.CurrentEntry;
                        if (targetEntry != null)
                        {
                            if (parts[0].EqualsIgnoreCase(Res.CopyTo))
                            {
                                var copyItems = entries.Where(e => e.CanCopy)
                                    .Select(e => Entry.GetItem(e.ParsingName, ShellItemParsingOptions.DontThrowOnError))
                                    .WhereNotNull()
                                    .ToArray();
                                try
                                {
                                    if (copyItems.Length > 0)
                                    {
                                        targetEntry.CopyTo(copyItems, null, false);
                                        RefreshEntries(null);
                                    }
                                }
                                finally
                                {
                                    copyItems.Dispose();
                                }
                            }
                            else
                            {
                                var moveItems = entries.Where(e => (e.CanCopy && e.CanDelete) || e.CanMove)
                                    .Select(e => Entry.GetItem(e.ParsingName, ShellItemParsingOptions.DontThrowOnError))
                                    .WhereNotNull()
                                    .ToArray();
                                try
                                {
                                    if (moveItems.Length > 0)
                                    {
                                        targetEntry.MoveTo(moveItems, null, false);
                                        RefreshEntries(null);
                                    }
                                }
                                finally
                                {
                                    moveItems.Dispose();
                                }
                            }
                        }
                    }
                }
                break;
        }

        void openNewInstance(string? screenDevicePath = null)
        {
            if (entries.Count != 1)
                return;

            var entry = entries[0];
            if (entry.CanBeOpenedAsFolder)
            {
                bool asAdmin;
                if (VIRTUAL_KEY.VK_SHIFT.IsPressed())
                {
                    asAdmin = true;
                }
                else
                {
                    asAdmin = actionId == nameof(Res.OpenInNewInstanceAdministrator);
                }
                ShellBatInstance.OpenNewInstance(new OpenNewInstanceOptions { ParsingName = entry.ParsingName, ScreenDevicePath = screenDevicePath, AsAdministrator = asAdmin });
            }
        }

        void mount(string? screenDevicePath = null)
        {
            if (entries.Count != 1)
                return;

            var entry = entries[0];
            if (!entry.Extension.IsDiskImage)
                return;

            // is it already mounted?
            foreach (var drive in Entry.MyComputer.EnumerateChildren())
            {
                if (drive.VirtualDiskImagePath != null && drive.VirtualDiskImagePath.EqualsIgnoreCase(entry.ParsingName))
                {
                    Navigate(drive, false);
                    return;
                }
            }

            var item = Entry.GetItem(entry.ParsingName, ShellItemParsingOptions.DontThrowOnError);
            if (item == null)
                return;

            using var idl = item.GetIdList();
            if (idl is null)
                return;

            // now, we want to get the mounted folder path once mounted to be able to navigate to it
            // and we'd like to prevent the shell view from opening its own window for that mounted folder
            // all this w/o polling...
            Application? _app = null;
            TaskUtilities.RunWithNewSTAThread(() =>
            {
                //Application.TraceInfo("Mount thread starting");
                string? mountedDisk = null;
                using var events = new WindowsEvents();
                events.WindowEvent += (s, e) =>
                {
                    if (e.IsRegistered)
                    {
                        foreach (var view in ShellView.Windows)
                        {
                            using var folder = view.GetFolder();
                            var fsp = folder?.SIGDN_FILESYSPATH;
                            if (fsp != null)
                            {
                                var vdi = IOUtilities.GetVirtualDiskImagePath(fsp, false);
                                if (vdi.EqualsIgnoreCase(entry.ParsingName))
                                {
                                    mountedDisk = fsp;
                                    //Application.TraceInfo($"Mounted disk image '{entry.ParsingName}' as '{fsp}'");
                                    var folderEntry = Entry.Get(null, fsp, ShellItemParsingOptions.DontThrowOnError);
                                    if (folderEntry != null)
                                    {
                                        RunTaskOnUIThread(() =>
                                        {
                                            switch (actionId)
                                            {
                                                case nameof(Res.Mount):
                                                    Navigate(folderEntry, false);
                                                    break;

                                                default:
                                                    ShellBatInstance.OpenNewInstance(new OpenNewInstanceOptions
                                                    {
                                                        ParsingName = folderEntry.ParsingName,
                                                        ScreenDevicePath = screenDevicePath,
                                                        AsAdministrator = false
                                                    });
                                                    break;
                                            }
                                        }, true);

                                        // hack: prevent the shell view from opening it's mounted folder window
                                        // IDispatch is an IWebBrowser2 interface here
                                        view.Dispatch.CallObjectMethod("Quit");
                                        _app?.Exit();
                                    }
                                    break;
                                }
                            }
                            view.Dispose();
                        }
                    }
                };

                _app = new Application();

                // kill this app/thread/pump after 5 seconds in case something goes wrong
                using var timer = new Timer(state =>
                {
                    _app?.Exit();
                }, null, 5000, 0);
                _app.Run();

                //Application.TraceInfo("Mount thread exiting");
            });

            ShellItem.InvokeContextMenuVerb([idl], "mount");
        }
    }

    public virtual Task<string[]?> GetSelection() => ExecuteScript($"getSelection();", JsonSourceGenerationContext.Default.StringArray);
    public virtual async Task<ShellBatChildWindow[]> GetChildWindows()
    {
        var ret = await ExecuteScript($"getWindows();", JsonSourceNoDefaultGenerationContext.Default.ShellBatChildWindowArray);
        ret ??= [];
        return ret;
    }

    protected virtual void RestoreChildWindows()
    {
        var windows = ShellBatInstance.Current.Settings.ChildWindows;
        if (windows == null || windows.Count == 0)
            return;

        foreach (var window in windows)
        {
            if (string.IsNullOrWhiteSpace(window.Id))
                continue;

            if (window.Id.StartsWith(WebPropertiesWindow.WindowId, StringComparison.OrdinalIgnoreCase))
            {
                var options = new WebWindowShowOptions
                {
                    ForceOpen = true,
                    Left = window.Left,
                    Top = window.Top,
                    Right = window.Right,
                    Bottom = window.Bottom,
                    ViewerId = window.ViewerId,
                    PinnedViewerId = window.PinnedViewerId,
                    ViewerOptions = window.ViewerOptions,
                };
                ShowProperties(window.Id, window.Parameters, options);
                continue;
            }

            if (window.Id.StartsWith(WebTerminalWindow.WindowId, StringComparison.OrdinalIgnoreCase))
            {
                if (window.Parameters == null)
                    continue;

                var options = new WebWindowShowOptions
                {
                    ForceOpen = true,
                    Left = window.Left,
                    Top = window.Top,
                    Right = window.Right,
                    Bottom = window.Bottom,
                    CommandLine = window.CommandLine,
                };
                RunTerminal(window.Parameters, options);
            }
        }
    }

    protected virtual void PasteOrPasteAsShortcut(string actionId)
    {
        ArgumentNullException.ThrowIfNull(actionId);

        Entry.ClearClipboardCuts();
        DropItems(Clipboard.GetDataObject(false)?.NativeObject, actionId == nameof(Res.PasteShortcut) ? DROPEFFECT.DROPEFFECT_LINK : DROPEFFECT.DROPEFFECT_COPY);
    }

    protected virtual async Task CopyOrCut(string actionId)
    {
        ArgumentNullException.ThrowIfNull(actionId);
        var selection = await GetSelection();
        if (selection == null || selection.Length == 0)
            return;

        CopyOrCut(actionId, selection.Select(n => Entry.Get(null, n, ShellItemParsingOptions.DontThrowOnError)).WhereNotNull());
    }

    protected virtual void CopyOrCut(string actionId, IEnumerable<Entry> entries)
    {
        ArgumentNullException.ThrowIfNull(actionId);
        if (entries == null || !entries.Any())
            return;

        using var dataObject = Entry.CreateDataObject(entries.Where(e => e.CanCopy), false);
        if (dataObject != null)
        {
            Entry.ClearClipboardCuts();

            if (actionId == nameof(Res.Cut))
            {
                dataObject.Object.SetPreferredDropEffect(DROPEFFECT.DROPEFFECT_MOVE);

                // also add a marker format to indicate an ongoing cut operation
                // switch between two format IDs because we monitor the clipboard changes
                _clipboardCutFormatId = _clipboardCutFormatId == _clipboardCutFormatId1 ? _clipboardCutFormatId2 : _clipboardCutFormatId1;
                dataObject.Object.SetData(
                    new FORMATETC
                    {
                        cfFormat = _clipboardCutFormatId,
                        tymed = (uint)TYMED.TYMED_HGLOBAL,
                    },
                    new STGMEDIUM
                    {
                        tymed = (uint)TYMED.TYMED_HGLOBAL,
                    },
                    true);
            }

            Clipboard.SetDataObject(dataObject);
        }

        // remember the cuts to display them differently
        if (actionId == nameof(Res.Cut))
        {
            foreach (var entry in entries.Where(e => (e.CanCopy && e.CanDelete) || e.CanMove))
            {
                Entry.AddToClipboardCut(entry);
            }
            RefreshEntries(null);
        }
    }

    protected virtual async Task OnChildWindowsChanged()
    {
        var windows = await GetChildWindows();
        var childWindows = windows.Where(w => !string.IsNullOrWhiteSpace(w.Id)).ToArray();
        if (childWindows.Length == 0 && (ShellBatInstance.Current.Settings.ChildWindows == null || ShellBatInstance.Current.Settings.ChildWindows.Count == 0))
            return;

        if (ShellBatInstance.Current.Settings.ChildWindows != null &&
            childWindows.Length == ShellBatInstance.Current.Settings.ChildWindows.Count &&
            !childWindows.Where((t, i) => !t.Equals(ShellBatInstance.Current.Settings.ChildWindows[i])).Any())
            return;

        ShellBatInstance.Current.Settings.ChildWindows = childWindows;
        ShellBatInstance.Current.SerializeSettings();
    }

    public virtual WebActions GetActions()
    {
        var actions = new WebActions(this, CurrentEntry ?? Entry.UserProfile);

        foreach (var view in ShellView.Windows)
        {
            using var folder = view.GetFolder();
            if (folder is null)
                continue;

            var pn = folder.SIGDN_DESKTOPABSOLUTEPARSING;
            if (pn is not null)
            {
                var name = folder.SIGDN_DESKTOPABSOLUTEEDITING ?? pn;
                actions.OpenFromExplorerList[pn] = name;
            }
        }

        actions.DetectsVisualStudioInstances = VisualStudioSolution.DetectsVisualStudioInstances();
        if (actions.DetectsVisualStudioInstances)
        {
            actions.RunScanVisualStudioSolutions();
        }
        return actions;
    }

    public virtual IReadOnlyList<WebMenuItem> GetEntryActions(Entry entry)
    {
        ArgumentNullException.ThrowIfNull(entry);
        var commands = new List<WebMenuItem>();
        if (entry.IsRecycleBin)
        {
            AddRecycleBinCommands(commands);
        }

        if (entry.VirtualDiskImagePath != null)
        {
            var item = new WebMenuItem { Id = nameof(Res.Eject), Html = Res.Eject, Icon = "fa-solid fa-eject", };
            commands.Add(item);
        }

        return commands;
    }

    protected virtual void AddRecycleBinCommands(IList<WebMenuItem> list)
    {
        ArgumentNullException.ThrowIfNull(list);

        var item = new WebMenuItem { Id = nameof(Res.EmptyRecycleBin), Html = Res.EmptyRecycleBin, Icon = "fa-regular fa-trash-can", };
        list.Add(item);

        item = new WebMenuItem { Id = nameof(Res.RestoreAllRecycleBinItems), Html = Res.RestoreAllRecycleBinItems, Icon = "fa-solid fa-trash-arrow-up", };
        list.Add(item);
        list.Add(new WebMenuItem { IsSeparator = true });
    }

    public void ExecuteAction(IEnumerable<string?>? parsingNames, string actionId)
    {
        if (parsingNames == null)
            return;

        var entries = new List<Entry>();
        foreach (var name in parsingNames)
        {
            if (name != null)
            {
                var entry = Entry.Get(null, name, ShellItemParsingOptions.DontThrowOnError);
                if (entry is not null)
                {
                    entries.Add(entry);
                }
            }
        }
        ExecuteAction(entries, actionId);
    }

    public virtual void ExecuteAction(IReadOnlyList<Entry> entries, string actionId)
    {
        try
        {
            ExecuteActionInternal(entries, actionId);
        }
        catch (Exception ex)
        {
            ShellBatInstance.LogError($"ExecuteAction '{actionId}' failed: {ex}");
            throw;
        }
    }

    public virtual void Run(string? parsingName, bool fromHistory)
    {
        var entry = Entry.Get(null, parsingName);
        if (entry is null || entry.IsFolder)
            return;

        WindowsExtensions.Open(entry.ParsingName);
    }

    internal void DisposePreviewViewer(PreviewViewer viewer) => _previewViewers.Remove(viewer);
    internal void RegisterPreviewViewer(PreviewViewer viewer)
    {
        if (_previewViewers.Contains(viewer))
            return;

        _previewViewers.Add(viewer);
    }

    internal void DisposeTerminal(string id)
    {
        if (_terminals.Remove(id))
        {
            CloseTerminal(id);
        }
    }

    public virtual WebTerminal? GetTerminal(string id, string key)
    {
        ArgumentNullException.ThrowIfNull(id);
        ArgumentNullException.ThrowIfNull(key);

        if (!_terminals.TryGetValue(id, out var terminal))
        {
            var entry = Settings.Current.GetTerminals().FirstOrDefault(t => t.Key.EqualsIgnoreCase(key));
            if (entry == null)
            {
                Application.TraceWarning($"Terminal '{key}' was not found.");
                return null;
            }

            terminal = new WebTerminal(entry, CurrentEntry ?? Entry.UserProfile, this, id);
            _terminals[id] = terminal;
        }
        return terminal;
    }

    public virtual WebWindow? GetWindow(string id, string? parameters)
    {
        ArgumentNullException.ThrowIfNull(id);

        if (id.StartsWith(WebPropertiesWindow.WindowId, StringComparison.OrdinalIgnoreCase))
        {
            var parsingName = parameters;
            var entry = Entry.Get(null, parsingName, ShellItemParsingOptions.DontThrowOnError);
            if (entry is null)
                return null;

            return new WebPropertiesWindow(this, entry, id);
        }

        if (CurrentEntry != null)
        {
            if (id.EqualsIgnoreCase(SearchType.FindStrings.ToString()))
                return new WebSearchWindow(this, SearchType.FindStrings, CurrentEntry);

            if (id.EqualsIgnoreCase(SearchType.WindowsSearch.ToString()))
            {
                using var ws = new WindowsSearch();
                var hr = ws.TryIncludedInCrawlScope($"file:///{CurrentEntry.ParsingName.Replace('\\', '/').Replace("'", "''")}", out var isIncluded, out var reason);
                if (hr.IsSuccess && !isIncluded)
                {
                    ShowError(new()
                    {
                        Text = string.Format(Res.WindowsSearchExcluded, CurrentEntry.ParsingName, Res.ResourceManager.GetString(reason.ToString()) ?? reason.ToString()),
                        Timer = 30000,
                        TimerProgressBar = true,
                        Width = "30vw",
                        ShowConfirmButton = true
                    });
                    return null;
                }

                return new WebSearchWindow(this, SearchType.WindowsSearch, CurrentEntry);
            }
        }

        if (id.StartsWith(WebTerminalWindow.WindowId, StringComparison.OrdinalIgnoreCase))
        {
            var terminal = Settings.Current.GetTerminals().FirstOrDefault(t => t.Key.EqualsIgnoreCase(parameters));
            if (terminal != null)
                return new WebTerminalWindow(this, id, CurrentEntry, terminal);

            Application.TraceWarning($"Terminal '{parameters}' was not found.");
            return null;
        }

        throw new NotSupportedException($"Window '{id}' is not supported.");
    }

    public virtual void CloseTerminal(string id)
    {
        ArgumentNullException.ThrowIfNull(id);
        var text = $"closeTerminal(\"{HttpUtility.JavaScriptStringEncode(id)}\");";
        RunTaskOnUIThread(() => ExecuteScript(text));
    }

    public virtual void RunTerminal(string terminalKey, WebWindowShowOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(terminalKey);
        var text = $"showTerminal(\"{WebTerminalWindow.WindowId}{_terminalId++}\", \"{HttpUtility.JavaScriptStringEncode(terminalKey)}\", {JsonSerializer.Serialize(options, JsonSourceGenerationContext.Default.WebWindowShowOptions)});";
        ExecuteScript(text);
    }

    public virtual void SelectEntry(string? parsingName)
    {
        if (parsingName == null)
            return;

        var text = $"selectParsingNameEntry(\"{HttpUtility.JavaScriptStringEncode(parsingName)}\");";
        ExecuteScript(text);
    }

    public virtual bool Navigate(string? parsingName, bool fromHistory, string? selectParsingName = null, bool refresh = false)
    {
        var entry = Entry.Get(null, parsingName, ShellItemParsingOptions.DontThrowOnError);
        if (entry is null)
            return false;

        return Navigate(entry, fromHistory, selectParsingName, refresh);
    }

    public virtual bool Navigate(Entry entry, bool fromHistory, string? selectParsingName = null, bool refresh = false)
    {
        ArgumentNullException.ThrowIfNull(entry);
        if (!refresh && entry == CurrentEntry)
            return false;

        if (!entry.CanBeOpenedAsFolder)
        {
            var parent = entry.Parent;
            if (parent is not null)
                return Navigate(parent, false, entry.ParsingName, refresh);

            return false;
        }

        //Application.TraceVerbose($"Navigating to '{entry.ParsingName}' (fromHistory={fromHistory}, selectParsingName={selectParsingName}, refresh={refresh})");
        string text;
        if (selectParsingName != null)
        {
            text = $"navigate(\"{HttpUtility.JavaScriptStringEncode(entry.ParsingName)}\", {fromHistory.ToString().ToLowerInvariant()}, \"{HttpUtility.JavaScriptStringEncode(selectParsingName)}\");";
        }
        else
        {
            text = $"navigate(\"{HttpUtility.JavaScriptStringEncode(entry.ParsingName)}\", {fromHistory.ToString().ToLowerInvariant()});";
        }
        ExecuteScript(text);
        return true;
    }

    public virtual void StopSearch(string? queryId)
    {
        queryId = queryId?.Nullify();
        if (queryId == null)
            return;

        if (_queries.Remove(queryId, out var query))
        {
            Application.TraceInfo($"Search '{queryId}' for '{query}' was stopped.");
        }
    }

    public virtual void StartSearch(SearchType type, string queryId, string query, object? options)
    {
        ArgumentNullException.ThrowIfNull(query);
        ArgumentNullException.ThrowIfNull(queryId);

        query = query.Trim();
        if (query == null)
            return;

        var path = CurrentEntry?.FileSystemPath;
        if (path == null)
            return;

        var disp = options as DirectN.IDispatch;
        var recursive = disp.GetProperty("recursive", false);
        var textFilesOnly = disp.GetProperty("textFilesOnly", false);

        switch (type)
        {
            case SearchType.FindStrings:
                Task.Run(() => FindStringsSearch(path, queryId, query, new FindStringsOptions
                {
                    Search = query,
                    IsRecursive = recursive,
                    InTextOnlyFiles = textFilesOnly,
                }));
                break;

            case SearchType.WindowsSearch:
                Task.Run(() => WindowsSearch(path, queryId, query, recursive));
                break;
        }
    }

    protected virtual void FindStringsSearch(string path, string queryId, string query, FindStringsOptions options)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentNullException.ThrowIfNull(queryId);
        ArgumentNullException.ThrowIfNull(query);
        ArgumentNullException.ThrowIfNull(options);

        _queries[queryId] = query;
        try
        {
            var index = 0;
            var i = 1; // start from 1 (0 will be undefined for js)
            options.SkipEntryFunc = skip;
            foreach (var result in FindStrings.Search(path, options))
            {
                if (!_queries.ContainsKey(queryId))
                {
                    Application.TraceInfo($"Search '{queryId}' for '{query}' was cancelled.");
                    break;
                }

                var wr = WebFindStringsResult.From(queryId, i++, result);
                wr.FilePath = Path.GetRelativePath(path, result.FilePath);
                var text = $"onSearchResult(\"{queryId}\", {JsonSerializer.Serialize(wr, JsonSourceGenerationContext.Default.WebFindStringsResult)});";
                RunTaskOnUIThread(() => ExecuteScript(text));
            }

            bool skip(string file)
            {
                index++;
                if (index % 10 == 0)
                {
                    var wr = new WebFindStringsResult
                    {
                        QueryId = queryId,
                        ParsingName = string.Empty,
                        Index = -2,
                        FilePath = string.Format(Res.SearchingCount, index),
                        Position = -1,
                        Text = string.Empty,
                    };
                    var text = $"onSearchResult(\"{queryId}\", {JsonSerializer.Serialize(wr, JsonSourceGenerationContext.Default.WebFindStringsResult)});";
                    RunTaskOnUIThread(() => ExecuteScript(text));
                }
                return false;
            }
        }
        finally
        {
            var last = new WebFindStringsResult
            {
                QueryId = queryId,
                ParsingName = string.Empty,
                Index = -1,
                FilePath = string.Empty,
                Position = -1,
                Text = string.Empty,
            };
            var text = $"onSearchResult(\"{queryId}\", {JsonSerializer.Serialize(last, JsonSourceGenerationContext.Default.WebFindStringsResult)});";
            RunTaskOnUIThread(() => ExecuteScript(text));

            _queries.Remove(queryId, out _);
        }
    }

    protected virtual void WindowsSearch(string path, string queryId, string query, bool isRecursive)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentNullException.ThrowIfNull(query);
        ArgumentNullException.ThrowIfNull(queryId);

        _queries[queryId] = query;
        try
        {
            // https://learn.microsoft.com/en-us/windows/win32/search/-search-3x-advancedquerysyntax
            // https://superuser.com/questions/1194876/an-extensive-list-of-windows-search-syntax
            using var ws = new WindowsSearch();

            var scope = isRecursive ? "scope" : "directory";
            ws.WhereRestrictions = $"AND {scope}='file:{path.Replace("'", "''")}'";
            ws.SelectColumns = "path";
            var hr = ws.TryGenerateSQLFromUserQuery(query, out var sql);
            if (hr.IsError)
            {
                Application.TraceError($"TryGenerateSQLFromUserQuery fail: {ComError.GetError(hr)}");
                return;
            }

            //Application.TraceVerbose($"Windows Search SQL: {sql}");
            var error = ws.TryExecuteSql(sql!, out var result);
            if (error != null)
            {
                Application.TraceError($"TryExecuteSql fail: {error}");
                return;
            }

            using (result)
            {
                var i = 1; // start from 1 (0 will be undefined for js)
                foreach (var row in result!.EnumerateRows())
                {
                    if (result.EnumerationError != null)
                        break;

                    var filePath = row.Store.GetNullifiedStringValue(ShellN.PropertyKeys.System.ParsingPath);
                    if (filePath == null || !Uri.TryCreate(filePath, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeFile)
                        continue;

                    var relativePath = Path.GetRelativePath(path, uri.AbsolutePath);
                    var wr = new WebWindowsSearchResult
                    {
                        QueryId = queryId,
                        ParsingName = uri.AbsolutePath,
                        Index = i,
                        FilePath = relativePath,
                    };

                    foreach (var kv in row.Store)
                    {
                        // remove some properties we don't want to expose
                        if (kv.Key == ShellN.PropertyKeys.System.ParsingPath ||
                            kv.Key == ShellN.PropertyKeys.System.ItemId)
                            continue;

                        var pd = kv.Key.ToDescription();
                        var name = pd?.DisplayName.Nullify() ?? pd?.CanonicalName.Nullify() ?? kv.Key.ToString();
                        wr.Properties[name] = kv.Value;
                    }

                    i++;
                    var text = $"onSearchResult(\"{queryId}\", {JsonSerializer.Serialize(wr, JsonSourceGenerationContext.Default.WebWindowsSearchResult)});";
                    RunTaskOnUIThread(() => ExecuteScript(text));

                    row.Dispose();
                }

                if (result.EnumerationError != null)
                {
                    Application.TraceError($"Enumeration error: {result.EnumerationError}");
                }
            }
        }
        finally
        {
            var last = new WebWindowsSearchResult
            {
                QueryId = queryId,
                ParsingName = string.Empty,
                Index = -1,
                FilePath = string.Empty,
            };

            var text = $"onSearchResult(\"{queryId}\", {JsonSerializer.Serialize(last, JsonSourceGenerationContext.Default.WebWindowsSearchResult)});";
            RunTaskOnUIThread(() => ExecuteScript(text));
            _queries.Remove(queryId, out _);
        }
    }

    public virtual void QuitAllInstances(bool includeThis = true)
    {
        ShellBatInstance.QuitAllRemoteInstances([System.Environment.ProcessId]);
        if (includeThis)
        {
            Close();
        }
    }

    public virtual void ArrangeInstances(WindowArrangement arrangement, int margin, bool allScreens)
    {
        var instances = ShellBatInstance.Current.All.ToArray();
        if (instances.Length <= 1) // includes this
            return;

        if (allScreens)
        {
            var monitors = DirectN.Extensions.Utilities.Monitor.All.ToArray();
            var instancesPerMonitor = (int)Math.Ceiling((double)instances.Length / monitors.Length);
            var index = 0;
            for (var i = 0; i < instances.Length; i += instancesPerMonitor)
            {
                var batch = instances.Skip(i).Take(instancesPerMonitor).ToArray();
                ShellBatInstance.ArrangeInstancesOnMonitor(arrangement, margin, monitors[index], batch);
                index++;
                if (index >= monitors.Length)
                    break;
            }
        }
        else
        {
            var monitor = GetMonitor(MONITOR_FROM_FLAGS.MONITOR_DEFAULTTONEAREST);
            if (monitor == null)
                return;

            ShellBatInstance.ArrangeInstancesOnMonitor(arrangement, margin, monitor, instances);
        }
    }

    public virtual void SetGlobalEvent(string eventType, string? json = null)
    {
        var text = $"setGlobalEvent(\"{eventType}\", {HttpUtility.JavaScriptStringEncode(json)});";
        ExecuteScript(text);
    }

    public virtual void SaveInstance()
    {
        using var fsd = new FileSaveDialog();
        fsd.SetDefaultExtension(FileExtension.ShellBatExtensionName);
        fsd.SetFileTypes([$"{Program.AppId} {Res.Instance} (*{FileExtension.ShellBatExtensionName})|*{FileExtension.ShellBatExtensionName}"]);
        fsd.SetTitle(Res.SaveInstance);
        if (!fsd.Show(Handle))
            return;

        using var targetItem = fsd.GetResult();
        if (targetItem == null)
            return;

        var path = targetItem.SIGDN_DESKTOPABSOLUTEPARSING;
        if (path == null)
            return;

        ShellBatInstance.Current.Settings.InstanceName = Path.GetFileNameWithoutExtension(path);
        ShellBatInstance.Current.Settings.Serialize(JsonSourceGenerationContext.Default);
    }

    public virtual void RestartAsAdministrator()
    {
        if (ProcessUtilities.RestartAsAdmin(true))
        {
            Program.RestartAsAdminRequested = true;
            Close();
        }
    }

    public virtual void RefreshCurrentTheme()
    {
        var filePath = ShellBatInstance.Current.Theme.FilePath;
        if (filePath != null)
        {
            LoadTheme(filePath);
        }
    }

    public virtual void LoadTheme(string? filePath = null)
    {
        ShellBatTheme? theme;
        filePath = filePath?.Nullify();
        if (filePath == null)
        {
            using var fod = new FileOpenDialog();
            fod.SetDefaultExtension(FileExtension.ShellBatThemeExtensionName);
            fod.SetFileTypes([$"{Program.AppId} {Res.Theme} (*{FileExtension.ShellBatThemeExtensionName})|*{FileExtension.ShellBatThemeExtensionName}"]);
            fod.SetTitle(Res.LoadTheme);
            if (!fod.Show(Handle))
                return;

            using var result = fod.GetResult();
            if (result == null)
                return;

            result.ComObject.Object.BindToHandler(null, ShellN.Constants.BHID_Stream, typeof(DirectN.IStream).GUID, out var unk);
            var strm = ComObject.FromPointer<DirectN.IStream>(unk);
            if (strm == null)
                return;

            using var stream = new StreamOnIStream(strm.Object, true);
            theme = ShellBatTheme.Deserialize(stream);
            theme?.FilePath = result.SIGDN_DESKTOPABSOLUTEPARSING!;
        }
        else
        {
            theme = ShellBatTheme.Deserialize(filePath);
            theme?.FilePath = filePath;
        }

        if (theme != null)
        {
            ShellBatInstance.Current.Theme = theme;
            ShellBatInstance.Current.SerializeSettings();
        }
    }

    public virtual void SaveTheme()
    {
        using var fsd = new FileSaveDialog();
        fsd.SetDefaultExtension(FileExtension.ShellBatThemeExtensionName);
        fsd.SetFileTypes([$"{Program.AppId} {Res.Theme} (*{FileExtension.ShellBatThemeExtensionName})|*{FileExtension.ShellBatThemeExtensionName}"]);
        fsd.SetTitle(Res.SaveInstance);
        if (!fsd.Show(Handle))
            return;

        using var result = fsd.GetResult();
        if (result == null || result.Extension == null || result.SIGDN_NORMALDISPLAY == null)
            return;

        using var folder = result.GetParent();
        if (folder == null)
            return;

        using var newItem = folder.CreateNewItem(result.SIGDN_NORMALDISPLAY, makeUniqueName: false);
        if (newItem == null)
            return;

        using var ctx = IBindCtxExtensions.CreateBindCtx(STGM.STGM_WRITE)!;
        newItem.ComObject.Object.BindToHandler(ctx.Object, ShellN.Constants.BHID_Stream, typeof(DirectN.IStream).GUID, out var unk);
        var strm = ComObject.FromPointer<DirectN.IStream>(unk);
        if (strm == null)
            return;

        using var stream = new StreamOnIStream(strm.Object, true);
        ShellBatInstance.Current.Theme.SerializeTo(stream);
    }

    public virtual async Task CheckForUpdates(bool doIfAvailable, CancellationToken cancellationToken = default)
    {
        if (_availableUpdate == null)
        {
            var fileName = Path.GetFileName(System.Environment.ProcessPath);
            var thisVersion = Version.Parse(Assembly.GetExecutingAssembly().GetFileVersion()!);
            var releases = await GitHubApi.ListReleasesAsync(cancellationToken);
            var last = releases.LastOrDefault(r => r.Assets != null && r.Assets.Any(a => a.Name.EqualsIgnoreCase(fileName)));
            if (last == null || thisVersion >= last.Version)
            {
                ShowInfo(new WebAlert { Text = Res.NoUpdates });
                return;
            }

            _availableUpdate = last;
        }

        if (!doIfAvailable)
        {
            ShowQuestion(new WebAlert
            {
                Text = string.Format(Res.UpdateAvailable, _availableUpdate.Version),
                ConfirmEventType = WebEventType.UpdateNow,
            });
            return;
        }

        if (!ShellBatInstance.StartUpdate())
        {
            ShowError(new WebAlert { Text = Res.UpdateIsNotSupported });
        }
    }

    public virtual async Task OpenViewers()
    {
        var selection = await GetSelection();
        if (selection != null && selection.Length > 0)
        {
            ShowProperties(selection[0], new WebWindowShowOptions { ForceOpen = true });
            return;
        }

        ShowToast(string.Format(Res.CommandNotAvailableInThisContext, nameof(ShellBatCommand.OpenViewers)));
    }

    public virtual void OpenMenu(MenuId menuId, IEnumerable<string>? selectPath = null)
    {
        var paths = new List<string>();
        if (selectPath != null)
        {
            paths.AddRange(selectPath.WhereNotNull());
        }

        var script = $"openMenu({(int)menuId}, {JsonSerializer.Serialize([.. paths], JsonSourceGenerationContext.Default.StringArray)});";
        ExecuteScript(script);
    }

    public virtual void OpenSearch(SearchType type)
    {
        var script = $"openSearch({(int)type});";
        ExecuteScript(script);
    }

    public virtual async Task OpenWithExplorerWithSelection()
    {
        var selection = await GetSelection();
        await OpenWithExplorer(selection);
    }

    public virtual async Task OpenWithExplorer(IEnumerable<string>? parsingNames = null)
    {
        var entry = CurrentEntry;
        if (entry is null)
            return;

        var parsingName = entry.ParsingName;
        if (string.IsNullOrWhiteSpace(parsingName))
        {
            ShellItem.OpenDesktop();
            return;
        }

        var names = new List<string> { parsingName };
        if (parsingNames != null)
        {
            foreach (var pn in parsingNames)
            {
                var name = pn.Nullify();
                if (name != null)
                {
                    names.Add(name);
                }
            }
        }
        ShellItem.OpenFoldersAndSelectItems(names);
    }

    public virtual void ShowLoading(bool show, string? info = null)
    {
        //Application.TraceVerbose($"ShowLoading: {show}, Info: {info}");
        if (show)
        {
            if (_loadingVisual == null)
            {
                _loadingVisual = Compositor.CreateSpriteVisual();
                _loadingVisual.Size = new Vector2(32, 32);
                _loadingVisual.AnchorPoint = new Vector2(0.5f, 0.5f);
                var cr = ClientRect;
                _loadingVisual.Offset = new Vector3(32, cr.Height - 64, 0);
                var brush = Compositor.CreateColorBrush(Color.FromArgb(255, 80, 80, 80));
                _loadingVisual.Brush = brush;

                RootVisual.Children.InsertAtTop(_loadingVisual);
                CompositorController.Commit();
            }
            else
            {
                var size = _loadingVisual.Size.X * _direction;
                if (_direction > 1 && size > 64)
                {
                    _direction = 0.9f;
                }
                else if (_direction < 1 && size < 32)
                {
                    _direction = 1.1f;
                }

                _loadingVisual.Size = new Vector2(size, size);

                _loadingVisual.Brush?.SafeDispose();
                _loadingHsv.Hue += 10;
                if (_loadingHsv.Hue >= 360)
                {
                    _loadingHsv.Hue -= 360;
                }

                var color = _loadingHsv.ToD3DCOLORVALUE();
                var brush = Compositor.CreateColorBrush(Color.FromArgb(255, color.BR, color.BG, color.BB));
                _loadingVisual.Brush = brush;
                _loadingVisual.RotationAngleInDegrees = (_loadingVisual.RotationAngleInDegrees + 10) % 360;

                CompositorController.Commit();
            }
        }
        else
        {
            if (_loadingVisual != null)
            {
                RootVisual.Children.Remove(_loadingVisual);
                _loadingVisual = null;
            }
        }
    }

    internal void ApplyTheme(ShellBatTheme theme)
    {
        if (GraphicsDevice == null)
            return;

        // transparency / acrylic is defined by having transparent main background color for webview
        var mainBackgroundColor = theme.MainBackgroundColor.Nullify();
        if (mainBackgroundColor != null &&
            D3DCOLORVALUE.TryParseFromName(mainBackgroundColor, out var color) &&
            color == D3DCOLORVALUE.Transparent)
        {
            var opacity = theme.Opacity.Clamp(0.1f, 1);
            var acrylicTintColor = theme.AcrylicTintColor.Nullify();
            if (acrylicTintColor != null && D3DCOLORVALUE.TryParseFromName(acrylicTintColor, out var tintColor))
            {
                WindowsExtensions.EnableBlurBehind(Handle);
                CompositionEffectBrush brush;
                try
                {
                    brush = AcrylicBrush.Create(
                        GraphicsDevice!,
                        tintColor,
                        theme.AcrylicTintOpacity,
                        theme.AcrylicTintLuminosityOpacity);
                }
                catch
                {
                    brush = AcrylicBrush.Create(
                        GraphicsDevice!,
                        D3DCOLORVALUE.White,
                        0.2f);
                }

                RootVisual.Brush = brush;
                RootVisual.Opacity = opacity;
            }
            else if (opacity != 1)
            {
                // old-style transparency
                ExtendedStyle |= WINDOW_EX_STYLE.WS_EX_LAYERED;
                DirectN.Functions.SetLayeredWindowAttributes(Handle, 0, (byte)(255 * opacity * 100 / 100), LAYERED_WINDOW_ATTRIBUTES_FLAGS.LWA_ALPHA);
                RootVisual.Brush = null;
            }
            else
            {
                ExtendedStyle &= ~WINDOW_EX_STYLE.WS_EX_LAYERED;
                DirectN.Functions.SetLayeredWindowAttributes(Handle, 0, 255, LAYERED_WINDOW_ATTRIBUTES_FLAGS.LWA_ALPHA);
                RootVisual.Brush = null;
            }
        }
        else
        {
            RootVisual.Brush = null;
        }

        if (theme.IsTopmost)
        {
            SetWindowPos(HWND.TOPMOST, 0, 0, 0, 0, SET_WINDOW_POS_FLAGS.SWP_NOMOVE | SET_WINDOW_POS_FLAGS.SWP_NOSIZE | SET_WINDOW_POS_FLAGS.SWP_NOACTIVATE);
        }
        else
        {
            SetWindowPos(HWND.NOTOPMOST, 0, 0, 0, 0, SET_WINDOW_POS_FLAGS.SWP_NOMOVE | SET_WINDOW_POS_FLAGS.SWP_NOSIZE | SET_WINDOW_POS_FLAGS.SWP_NOACTIVATE);
        }

        if (WebView != null)
        {
            var script = $"applyTheme({theme.SerializeForWebView()});";
            ExecuteScript(script);
        }
    }

    protected override void CreateDeviceResources()
    {
        base.CreateDeviceResources();
        if (_pdfRenderer == null)
        {
            using var device = Device.As<IDXGIDevice>()!;
            DirectN.Functions.PdfCreateRenderer(device.Object, out var obj).ThrowOnError();
            if (obj != null)
            {
                _pdfRenderer = new PdfRenderer(new ComObject<IPdfRendererNative>(obj));
            }
        }
    }

    protected override void OnMouseButtonDoubleClick(object? sender, MouseButtonEventArgs e)
    {
        if (e.Button == MouseButton.Left)
        {
            var caption = GetCaptionRect();
            if (caption != null && caption.Value.Contains(e.Point.ScreenToClient(Handle)))
            {
                ExecuteScript("editAddress();");
                return;
            }
        }
        base.OnMouseButtonDoubleClick(sender, e);
    }

    public virtual bool RunCommand(ShellBatCommand command)
    {
        ArgumentNullException.ThrowIfNull(command);
        switch (command.Name)
        {
            case nameof(ShellBatCommand.Copy):
                _ = CopyOrCut(nameof(Res.Copy));
                return true;

            case nameof(ShellBatCommand.Cut):
                _ = CopyOrCut(nameof(Res.Cut));
                return true;

            case nameof(ShellBatCommand.Paste):
                PasteOrPasteAsShortcut(nameof(Res.Paste));
                return true;

            case nameof(ShellBatCommand.GoBack):
                ExecuteScript("goBack();");
                return true;

            case nameof(ShellBatCommand.GoForward):
                ExecuteScript("goForward();");
                return true;

            case nameof(ShellBatCommand.GoUp):
                ExecuteScript("navigate(window.upParsingName);");
                return true;

            case nameof(ShellBatCommand.EditAddress):
                OnFocusChanged(true);
                ExecuteScript("editAddress();");
                return true;

            case nameof(ShellBatCommand.OpenKeyboardShortcuts):
                ShowPropertyGrid(PropertyGridType.KeyboardShortcuts);
                return true;

            case nameof(ShellBatCommand.OpenInstanceSettings):
                ShowPropertyGrid(PropertyGridType.InstanceSettings);
                return true;

            case nameof(ShellBatCommand.OpenSettings):
                ShowPropertyGrid(PropertyGridType.Settings);
                return true;

            case nameof(ShellBatCommand.CloseApplication):
                Close();
                return true;

            case nameof(ShellBatCommand.MinimizeWindow):
                Show(SHOW_WINDOW_CMD.SW_MINIMIZE);
                return true;

            case nameof(ShellBatCommand.MaximizeWindow):
                if (IsZoomed)
                {
                    Show(SHOW_WINDOW_CMD.SW_RESTORE);
                }
                else
                {
                    Show(SHOW_WINDOW_CMD.SW_MAXIMIZE);
                }
                return true;

            case nameof(ShellBatCommand.Reload):
                ExecuteScript("window.location.reload();");
                return true;

            case nameof(ShellBatCommand.OpenDevTools):
                WebView?.Object.OpenDevToolsWindow();
                return true;

            case nameof(ShellBatCommand.OpenWithExplorer):
                _ = OpenWithExplorerWithSelection();
                return true;

            case nameof(ShellBatCommand.OpenDesktop):
                Navigate(string.Empty, false);
                return true;

            case nameof(ShellBatCommand.OpenThisPC):
                Navigate($"::{ShellN.Constants.CLSID_MyComputer:B}", false);
                return true;

            case nameof(ShellBatCommand.OpenRecycleBin):
                Navigate($"::{ShellN.Constants.CLSID_RecycleBin:B}", false);
                return true;

            case nameof(ShellBatCommand.CloseWindow):
                ExecuteScript("closeWindow();");
                return true;

            case nameof(ShellBatCommand.FindStrings):
                OpenSearch(SearchType.FindStrings);
                return true;

            case nameof(ShellBatCommand.OpenSearch):
                OpenSearch(SearchType.WindowsSearch);
                return true;

            case nameof(ShellBatCommand.OpenTerminal):
                OpenMenu(MenuId.AppActions, [Res.NewTerminal, Res.CommandPrompt]);
                return true;

            case nameof(ShellBatCommand.OpenHistory):
                OpenMenu(MenuId.AppHistory);
                return true;

            case nameof(ShellBatCommand.OpenFavorites):
                OpenMenu(MenuId.AppFavorites);
                return true;

            case nameof(ShellBatCommand.OpenViewers):
                _ = OpenViewers();
                return true;

            case nameof(ShellBatCommand.OpenNewInstance):
                ShellBatInstance.OpenNewInstance(new OpenNewInstanceOptions { ParsingName = CurrentEntry?.ParsingName });
                return true;

            case nameof(ShellBatCommand.OpenNewInstanceOnOtherScreen):
                var screen = WebScreen.All.FirstOrDefault(s => !s.IsThis);
                if (screen != null)
                {
                    ShellBatInstance.OpenNewInstance(new OpenNewInstanceOptions { ParsingName = CurrentEntry?.ParsingName, ScreenDevicePath = screen.DevicePath });
                    return true;
                }
                break;

            case nameof(ShellBatCommand.RestartAsAdministrator):
                RestartAsAdministrator();
                return true;

            case nameof(ShellBatCommand.CopyFullPath):
                var fp = CurrentEntry?.ParsingName;
                if (fp != null)
                {
                    Clipboard.SetText(fp);
                    ShowToast(Res.FullPathCopiedToClipboard);
                }
                return true;
        }
        return false;
    }

    protected override void OnNavigationStarting(object? sender, NavigationEventArgs e)
    {
        // we don't want to navigate to external links inside the app
        if (e.Uri != null && Uri.TryCreate(e.Uri, UriKind.Absolute, out var uri) && uri.Scheme != Uri.UriSchemeFile)
        {
            e.Cancel = true;
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = e.Uri,
                    UseShellExecute = true,
                });
            }
            catch
            {
                // continue
            }
        }
    }

    protected override void OnKeyDown(object? sender, KeyEventArgs e)
    {
        // note: even if accelerators are disabled in WebView settings we won't
        // get here if dev tools are opened and focused
        switch (e.Key)
        {
            case VIRTUAL_KEY.VK_ESCAPE:
                if (Entry.ClearClipboardCuts() > 0)
                {
                    RefreshEntries(null);
                    Clipboard.Empty();
                    e.Handled = true;
                }
                break;

            default:
                var ck = ShellBatCommand.GetCommandByKey(e);
                if (ck != null)
                {
                    e.Handled = RunCommand(ck);
                    if (!e.Handled)
                    {
                        ShowToast(string.Format(Res.CommandNotAvailableInThisContext, ck.Name));
                    }
                }
                break;
        }

        base.OnKeyDown(sender, e);
    }

    // adjust by webview zoom level
    protected override RECT? GetCaptionRect() => ScaleRect(_captionRect);
    public virtual RECT ScaleRect(RECT rect)
    {
        rect.left = (int)(rect.left * _scale);
        rect.top = (int)(rect.top * _scale);
        rect.right = (int)(rect.right * _scale);
        rect.bottom = (int)(rect.bottom * _scale);
        return rect;
    }

    public virtual void CreateNewFolder()
    {
        CurrentEntry?.CreateNewFolder();
    }

    protected override void OnNavigationCompleted(object? sender, NavigationEventArgs e)
    {
        base.OnNavigationCompleted(sender, e);
        _ = MonacoExtensions.LoadLanguages(WebView!);
    }

    protected override CoreWebView2EnvironmentOptions? GetEnvironmentOptions()
    {
        var options = new CoreWebView2EnvironmentOptions();

        // disable CORS (for js module)
        options.put_AdditionalBrowserArguments(PWSTR.From("--allow-file-access-from-files --disable-web-security"));
        if (ShellBatInstance.Current.Settings.ScrollbarStyle == COREWEBVIEW2_SCROLLBAR_STYLE.COREWEBVIEW2_SCROLLBAR_STYLE_FLUENT_OVERLAY)
        {
            options.put_ScrollBarStyle(COREWEBVIEW2_SCROLLBAR_STYLE.COREWEBVIEW2_SCROLLBAR_STYLE_FLUENT_OVERLAY);
        }
        return options;
    }

    protected override void ControllerCreated()
    {
        // this is for a full support of .NET Task or Task<T> methods
        // unfortunately, uses undocumented (private) interfaces
        if (WebView!.Object is ICoreWebView2PrivatePartial partial)
        {
            partial.AddHostObjectHelper(new WebViewHostObjectHelper()).ThrowOnError();
            DispatchObject.ContinueOnAsync = true;
            DispatchObject.OneStepInvoke = true;
        }

        WebView.Object.get_Settings(out var settingsObj).ThrowOnError();
        using var settings = new ComObject<ICoreWebView2Settings3>(settingsObj);
        settingsObj.put_IsBuiltInErrorPageEnabled(false).ThrowOnError();
        settingsObj.put_AreDefaultContextMenusEnabled(false);
        settingsObj.put_IsStatusBarEnabled(false).ThrowOnError();
        settings.Object.put_AreBrowserAcceleratorKeysEnabled(false).ThrowOnError();

        // get IUnknown from the host object and wrap it in a VARIANT
        ComObject.WithComInstance(_hostObject, unk =>
        {
            using var variant = new Variant(unk, VARENUM.VT_UNKNOWN);
            var detached = variant.Detached;
            WebView.Object.AddHostObjectToScript(PWSTR.From("dotnet"), ref detached).ThrowOnError();
        }, true);

#if DEBUG
        if (Program.DevEnabled)
        {
            WebView.Object.OpenDevToolsWindow();
        }
#else
        settingsObj.put_AreDevToolsEnabled(Program.DevEnabled);
        if (Program.DevEnabled)
        {
            WebView.Object.OpenDevToolsWindow();
        }
#endif

        Task.Run(NavigateToIndex);
    }

    protected virtual async Task NavigateToIndex()
    {
        await WebRootResources.EnsureFilesAsync();
        await ThemeResources.EnsureFilesAsync();
        _ = RunTaskOnUIThread(() =>
        {
            var index = Path.Combine(Settings.WebRootPath, Settings.Index);
            WebView!.Object.Navigate(PWSTR.From(index)).ThrowOnError();
            OnFocusChanged(true);
        });
    }

    protected override HRESULT OnBeforeDragEnter(IDataObject dataObject, MODIFIERKEYS_FLAGS flags, POINTL point, ref DROPEFFECT effect, out bool handled)
    {
        var isUs = dataObject.Has(_clipboardDragFormatId);
        if (isUs)
        {
            //Application.TraceVerbose($"Can't d&d from the same instance");
            effect = DROPEFFECT.DROPEFFECT_NONE;
            _draggingIsUs = true;
        }

        return base.OnBeforeDragEnter(dataObject, flags, point, ref effect, out handled);
    }

    protected override HRESULT OnBeforeDragOver(MODIFIERKEYS_FLAGS flags, POINTL point, ref DROPEFFECT effect, out bool handled)
    {
        if (_draggingIsUs == true)
        {
            effect = DROPEFFECT.DROPEFFECT_NONE;
        }
        return base.OnBeforeDragOver(flags, point, ref effect, out handled);
    }

    protected override void OnAfterDragLeave()
    {
        if (_draggingIsUs != null)
        {
            _draggingIsUs = null;
        }
        base.OnAfterDragLeave();
    }

    protected override void OnAfterDrop(IDataObject dataObject, MODIFIERKEYS_FLAGS keyFlags, POINTL point, DROPEFFECT effect) => DropItems(dataObject, effect);
    protected virtual void DropItems(IDataObject? dataObject, DROPEFFECT effect)
    {
        if (dataObject is null || CurrentEntry is null)
            return;

        var dataObj = new ComObject<IDataObject>(dataObject, false);
        //var da = new DataObject(dataObj, false);
        //Application.TraceVerbose($"effect={effect}: {string.Join(", ", da.GetFormats())}");

        using var folder = Entry.GetItem(CurrentEntry.ParsingName, ShellItemParsingOptions.DontThrowOnError);
        if (folder is null)
            return;

        var isUs = dataObject.Has(_clipboardDragFormatId);
        if (isUs)
        {
            Application.TraceVerbose($"Can't paste from the same instance");
            return;
        }

        var preferred = dataObject.GetPreferredDropEffect(effect);

        if (effect == DROPEFFECT.DROPEFFECT_LINK) // paste shortcut
        {
            if (preferred.HasFlag(DROPEFFECT.DROPEFFECT_LINK))
            {
                var items = ShellItem.ArrayFromDataObject(dataObject);
                if (items.Count > 0)
                {
                    var folderPath = folder.SIGDN_FILESYSPATH;
                    if (folderPath is null)
                        return;

                    try
                    {
                        foreach (var item in items)
                        {
                            using var link = Link.FromNativeObject();
                            if (link == null)
                                continue;

                            var name = item.SIGDN_NORMALDISPLAY;
                            if (name == null)
                                continue;

                            name = IOUtilities.NameToValidFileName(Path.GetFileNameWithoutExtension(name));
                            if (name == null)
                                continue;

                            var filePath = Path.Combine(folderPath, name + ".lnk");
                            using var idl = item.GetIdList();
                            link.TargetIdList = idl;
                            link.Save(filePath);
                        }
                    }
                    finally
                    {
                        items.Dispose();
                    }
                }
            }
            return;
        }

        // else COPY or MOVE
        var hr = ShellN.Functions.SHCreateShellItemArrayFromDataObject(dataObject, typeof(IShellItemArray).GUID, out var unk);
        if (hr.IsError) // we're probably copying from machine to machine (only CFSTR_FILEDESCRIPTORW may be there and it's not supported by IFileOperation)
        {
            Application.TraceError($"SHCreateShellItemArrayFromDataObject failed: {hr}");
            var files = dataObj.Object.GetFileDescriptors();
            if (files.Count == 0)
                return;

            // copy to temp and then move from there
            var temp = Path.GetTempPath();
            var format = Clipboard.RegisterFormat(ShellN.Constants.CFSTR_FILECONTENTS);
            var list = new List<ShellItem>();
            try
            {
                for (var i = 0; i < files.Count; i++)
                {
                    var file = files[i];
                    var path = IOUtilities.NameToValidFileName(file.cFileName.ToString());
                    if (path == null)
                        continue;

                    dataObj.Object.TryGet(format, i, out Stream? stream);
                    if (stream == null)
                        continue;

                    var filePath = Path.Combine(temp, path);
                    using (stream)
                    {
                        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
                        stream.CopyTo(fs);
                    }

                    var item = ShellItem.FromParsingName(filePath, throwOnError: false);
                    if (item != null)
                    {
                        list.Add(item);
                    }

                    using var tempDataObject = ShellItem.CreateDataObject(list, disposeItems: false, throwOnError: false);
                    if (tempDataObject == null)
                    {
                        Application.TraceWarning("Failed to create temp drag data object.");
                        return;
                    }

                    unk = ComObject.GetOrCreateComInstance<IDataObject>(tempDataObject);
                    if (unk == 0)
                    {
                        Application.TraceWarning("Failed to get IUnknown from temp drag data object.");
                        return;
                    }

                    preferred = DROPEFFECT.DROPEFFECT_MOVE;
                }
            }
            finally
            {
                list.Dispose();
            }

            // continue with temp data object
        }

        using var array = ComObject.FromPointer<IShellItemArray>(unk); // for release
        using var fo = new FileOperation();
        fo.SetOperationFlags(FILEOPERATION_FLAGS.FOF_RENAMEONCOLLISION |
            FILEOPERATION_FLAGS.FOFX_NOSKIPJUNCTIONS |
            FILEOPERATION_FLAGS.FOFX_KEEPNEWERFILE |
            FILEOPERATION_FLAGS.FOFX_PRESERVEFILEEXTENSIONS);

        if (preferred.HasFlag(DROPEFFECT.DROPEFFECT_MOVE))
        {
            dataObject.SetPerformedDropEffect(DROPEFFECT.DROPEFFECT_MOVE);
            hr = fo.NativeObject.MoveItems(unk, folder.NativeObject);
            if (hr.IsError)
            {
                Application.TraceError($"MoveItems MOVE failed: {hr}");
                return;
            }

            hr = fo.PerformOperations(false);
            if (hr.IsError)
            {
                Application.TraceError($"PerformOperations MOVE failed: {hr}");
                return;
            }

            dataObject.SetPasteSucceededEffect(DROPEFFECT.DROPEFFECT_MOVE);
            return;
        }

        if (preferred.HasFlag(DROPEFFECT.DROPEFFECT_COPY))
        {
            dataObject.SetPerformedDropEffect(DROPEFFECT.DROPEFFECT_COPY);
            hr = fo.NativeObject.CopyItems(unk, folder.NativeObject);
            if (hr.IsError)
            {
                Application.TraceError($"CopyItems COPY failed: {hr}");
                return;
            }

            hr = fo.PerformOperations(false);
            if (hr.IsError)
            {
                Application.TraceError($"PerformOperations COPY failed: {hr}");
                return;
            }

            dataObject.SetPasteSucceededEffect(DROPEFFECT.DROPEFFECT_COPY);
            return;
        }
    }

    // drag & drop with files *out of* the webview is pooe because it only allows
    // one file ("DownloadUrl"), so we have to implement all this ourselves
    // the nice thing though is that we can now provide virtual files (pure Shell Items) and folder
    protected override async void OnMouseButtonDown(object? sender, MouseButtonEventArgs e)
    {
        if (e.Button == MouseButton.Left)
        {
            try
            {
                _draggingNames = await ExecuteScript($"getParsingNameAndSelectionAtPoint({e.Point.x}, {e.Point.y});", JsonSourceGenerationContext.Default.StringArray);
            }
            catch
            {
                // can happen if webview is not ready or being disposed
                _draggingNames = null;
            }

            if (_draggingNames != null && _draggingNames.Length == 0)
            {
                _draggingNames = null;
            }

            if (_draggingNames != null)
            {
                //Application.TraceVerbose($"Entries at point: {string.Join(System.Environment.NewLine, _draggingNames)}");
                _draggingPoint = e.Point;
                e.Handled = true;
            }
        }
    }

    protected override void OnMouseButtonUp(object? sender, MouseButtonEventArgs e)
    {
        if (e.Button == MouseButton.Left)
        {
            DirectN.Functions.ReleaseCapture();
            _draggingNames = null;
        }
    }

    protected override void OnMouseMove(object? sender, MouseEventArgs e)
    {
        if (_draggingNames != null)
        {
            var deltaLength = (e.Point.x - _draggingPoint.x) * (e.Point.x - _draggingPoint.x) + (e.Point.y - _draggingPoint.y) * (e.Point.y - _draggingPoint.y);
            if (deltaLength >= 100) // 10 pixels
            {
                var items = new Dictionary<string, ShellItem>(StringComparer.OrdinalIgnoreCase);
                foreach (var name in _draggingNames)
                {
                    var item = ShellItem.FromParsingName(name, null);
                    if (item != null)
                    {
                        items[name] = item;
                    }
                }
                _draggingNames = null;

                if (items.Count > 0)
                {
                    //DirectN.Functions.ReleaseCapture();
                    using var dragDataObject = ShellItem.CreateDataObject(items.Values);
                    if (dragDataObject == null)
                    {
                        Application.TraceWarning("Failed to create drag data object.");
                        _draggingNames = null;
                        return;
                    }

                    // this says "it's us who started the drag-drop"
                    // note d&d can work between two ShellBat instances (id is different)
                    dragDataObject.Object.Set(_clipboardDragFormatId);

                    ShellN.Functions.SHDoDragDrop(Handle, dragDataObject.Object, null, DROPEFFECT.DROPEFFECT_COPY | DROPEFFECT.DROPEFFECT_MOVE | DROPEFFECT.DROPEFFECT_LINK, out var effect);
                }
            }
        }
    }

    protected override void OnMouseWheel(object? sender, MouseWheelEventArgs e)
    {
        base.OnMouseWheel(sender, e);
        if (VIRTUAL_KEY.VK_CONTROL.IsPressed() && ShellBatInstance.Current.Settings.ViewBy == ViewBy.Images)
        {
            var size = ShellBatInstance.Current.Settings.ThumbnailsSize;
            var ratio = 1 + (e.Delta / 10f);
            size = Math.Clamp((int)(size * ratio), 10, 500);
            ExecuteScript($"setThumbnailsSize({size});");
            ShellBatInstance.Current.Settings.ThumbnailsSize = size;
            e.Handled = true;
        }
    }

    protected override LRESULT? WindowProc(HWND hwnd, uint msg, WPARAM wParam, LPARAM lParam)
    {
        //Application.TraceVerbose($"msg={MessageDecoder.Decode(hwnd, msg, wParam, lParam)}");
        // special case for context menu messages https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-icontextmenu3-handlemenumsg2
        if (msg == MessageDecoder.WM_INITMENUPOPUP ||
            msg == MessageDecoder.WM_MENUSELECT ||
            msg == MessageDecoder.WM_DRAWITEM ||
            msg == MessageDecoder.WM_MEASUREITEM ||
            msg == MessageDecoder.WM_MENUCHAR)
        {
            if (ShellItem.OnContextMenuWindowMessage(Handle, msg, wParam, lParam, out var result).IsSuccess)
                return result;
        }

        if (msg == MessageDecoder.WM_CLIPBOARDUPDATE)
        {
            OnClipboardChanged();
            return 0;
        }

        if (msg == MessageDecoder.WM_DEVICECHANGE)
        {
            //Application.TraceVerbose(MessageDecoder.Decode(hwnd, msg, wParam, lParam));
            if (CurrentEntry?.IsMyComputer == true)
            {
                RefreshEntries(null);
            }
        }

        return base.WindowProc(hwnd, msg, wParam, lParam);
    }

    protected virtual void OnClipboardChanged()
    {
        //Application.TraceVerbose("Clipboard changed");

        // don't clear if our cut format is still present
        using var dataObject = Clipboard.GetDataObject(false);
        if (dataObject != null)
        {
            foreach (var format in dataObject.EnumerateFormats(false))
            {
                var name = Clipboard.GetFormatName(format.cfFormat);
                //Application.TraceVerbose($"Clipboard format: cfFormat={format.cfFormat} '{name}', tymed={format.tymed}");
            }
        }

        if (dataObject != null && dataObject.EnumerateFormats(false).All(f => f.cfFormat != _clipboardCutFormatId))
        {
            Entry.ClearClipboardCuts();
            RefreshEntries(null);
        }
    }

    protected override bool OnMoved()
    {
        SavePlacement();
        return base.OnMoved();
    }

    protected override bool OnResized(WindowResizedType type, SIZE size)
    {
        if (type != WindowResizedType.Minimized)
        {
            SavePlacement();
        }

        return base.OnResized(type, size);
    }

    protected override Icon? LoadCreationIcon() => Program.LoadIcon();

    protected virtual void SavePlacement()
    {
        if (!_created)
            return;

        var placement = WINDOWPLACEMENT.GetPlacement(Handle);
        var pos = placement.rcNormalPosition;
        if (placement.length > 0 ||
            pos.Width <= 0 ||
            pos.Height <= 0)
        {
            if (!placement.Equals(ShellBatInstance.Current.Settings.MainPlacement))
            {
                ShellBatInstance.Current.Settings.MainPlacement = placement;
                ShellBatInstance.Current.SerializeSettings();
            }
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (Handle.Value != 0)
            {
                DirectN.Functions.RemoveClipboardFormatListener(Handle);
            }

            _terminals.Values.Dispose();
            _watcher?.Dispose();
            _watcher = null;
            Interlocked.Exchange(ref _pdfRenderer, null)?.Dispose();
        }
        base.Dispose(disposing);
    }

    public static PdfRenderer? GetPdfRenderer()
    {
        foreach (var window in ShellBatApplication.ShellBatWindows)
        {
            var renderer = window._pdfRenderer;
            if (renderer != null)
                return renderer;
        }
        return null;
    }

    private void LockProperties(ShellBatInstanceCommand? cmd)
    {
        if (cmd != null)
        {
            // show properties was open by this cmd
            _propertiesSourceProcessId = cmd.ProcessId;
            return;
        }

        // cancel
        _propertiesSourceProcessId = null;
    }

    internal bool IsPropertiesLockedBy(ShellBatInstanceCommand cmd) => _propertiesSourceProcessId == cmd.ProcessId;

    [System.Runtime.InteropServices.Marshalling.GeneratedComClass]
    private sealed partial class Site(ShellBatWindow window) :
        DirectN.IServiceProvider,
        IObjectWithSite,
        //ICustomQueryInterface,
        IHandlerActivationHost,
        DirectN.IOleWindow,
        //IWaitCursorManager,
        IDisposable
    {
        private nint _site;

        //public CustomQueryInterfaceResult GetInterface(ref Guid iid, out nint ppv)
        //{
        //    Application.TraceVerbose($"Site iid:{iid:B}");
        //    ppv = 0;
        //    return CustomQueryInterfaceResult.NotHandled;
        //}

        public HRESULT QueryService(in Guid guidService, in Guid riid, out nint ppvObject)
        {
            ppvObject = ComObject.GetOrCreateComInstance(this, riid, CreateComInterfaceFlags.None);
            //Application.TraceInfo("Site guidService:" + guidService.ToString("B") + " riid:" + riid.ToString("B") + " ppv:" + ppvObject);
            return ppvObject == 0 ? DirectN.Constants.E_NOINTERFACE : DirectN.Constants.S_OK;
        }

        public HRESULT GetSite(in Guid riid, out nint ppvSite)
        {
            if (_site != 0)
                return Marshal.QueryInterface(_site, riid, out ppvSite);

            ppvSite = 0;
            return DirectN.Constants.E_NOINTERFACE;
        }

        public HRESULT SetSite(nint pUnkSite)
        {
            Dispose();
            Marshal.AddRef(pUnkSite);
            _site = pUnkSite;
            //Application.TraceInfo($"{pUnkSite}");
            return DirectN.Constants.S_OK;
        }

        public HRESULT ContextSensitiveHelp(BOOL fEnterMode) => DirectN.Constants.E_NOTIMPL;
        public HRESULT GetWindow(out HWND phwnd)
        {
            phwnd = window.Handle;
            return DirectN.Constants.S_OK;
        }

        public void Dispose()
        {
            var site = Interlocked.Exchange(ref _site, 0);
            if (site != 0)
            {
                Marshal.Release(site);
            }
        }

        public HRESULT BeforeCoCreateInstance(in Guid clsidHandler, IShellItemArray itemsBeingActivated, IHandlerInfo handlerInfo)
        {
            //Application.TraceVerbose($"Site clsidHandler: {clsidHandler:B}");
            return 0;
        }

        public HRESULT BeforeCreateProcess(PWSTR applicationPath, PWSTR commandLine, IHandlerInfo handlerInfo)
        {
            //Application.TraceVerbose($"Site applicationPath: {applicationPath}");
            return 0;
        }

        //HRESULT IWaitCursorManager.Start(CURSORID id)
        //{
        //    Application.TraceVerbose($"Wait cursor start: {id}");
        //    return 0;
        //}

        //HRESULT IWaitCursorManager.Restore()
        //{
        //    Application.TraceVerbose($"Wait cursor restore");
        //    return 0;
        //}

        //HRESULT IWaitCursorManager.Stop(CURSORID id)
        //{
        //    Application.TraceVerbose($"Wait cursor stop: {id}");
        //    return 0;
        //}
    }
}
