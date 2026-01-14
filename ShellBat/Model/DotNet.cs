namespace ShellBat.Model;

// entry point to .NET from JavaScript
[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public sealed partial class DotNet(ShellBatWindow window) : DispatchObject
{
    public event EventHandler<WebEventArgs>? Event;
    public event EventHandler<NavigateEventArgs>? Navigate;
    public event EventHandler<WebErrorEventArgs>? Error;

#pragma warning disable CA1822 // Mark members as static

    public string? ServerUrl { get; internal set; }
    public bool IsAdministrator { get; } = GetIsAdministrator();

    private static bool GetIsAdministrator()
    {
        var type = SystemUtilities.GetTokenElevationType();
        return type == TokenElevationType.Full || type == TokenElevationType.Default;
    }

    public WebFolder? GetFolder(string? parsingName, bool fromHistory)
    {
        var entry = Entry.Get(null, parsingName, ShellItemParsingOptions.DontThrowOnError);
        Navigate?.Invoke(this, new NavigateEventArgs(entry, fromHistory));
        var folder = entry is not null ? new WebFolder(window, entry) : null;

        if (!window.CanChangeCursor)
        {
            window.CanChangeCursor = true;
            window.Cursor = Cursor.Arrow.Handle;
        }
        return folder;
    }

    public string? GetEntryEditName(string parsingName)
    {
        ArgumentNullException.ThrowIfNull(parsingName);
        var entry = Entry.Get(null, parsingName, ShellItemParsingOptions.DontThrowOnError);
        if (entry is null)
            return null;

        return entry.DisplayName;
    }

    public void ShowLoading(bool show, string? info) => window.ShowLoading(show, info);
    public void OpenConfigurationFolder() => WindowsExtensions.Open(Settings.ConfigurationDirectoryPath);
    public void CheckForUpdates() => _ = window.CheckForUpdates(false);
    public void SaveInstance() => window.SaveInstance();
    public void OpenWithExplorer(object?[]? selection) => _ = window.OpenWithExplorer(selection?.Select(o => o?.ToString()).WhereNotNull());
    public bool IsCurrentFavorite() => Settings.Current.GetFavoritesIndex(window.CurrentEntry) >= 0;
    public Task<bool> ToggleCurrentFavorite() => Settings.Current.ToggleFavorite(window.CurrentEntry);

    public bool RemoveHistoryEntries(string? rootParsingName)
    {
        if (History.Current.RemoveEntries(rootParsingName))
        {
            Settings.Current.SerializeToConfiguration();
            return true;
        }
        return false;
    }

    public WebPropertyGrid GetPropertyGrid(PropertyGridType type)
    {
        WebPropertyGrid grid;
        switch (type)
        {
            case PropertyGridType.Info:
                var info = new ShellBatSysInfo(window);
                grid = WebPropertyGrid.Reflect(info);

                grid.Options.IsReadOnly = true;
                grid.Options.Title = Res.SysInfo;
                grid.Options.Categories.Add(new WebPropertyGridCategory { Key = "System", Collapsed = true });
                break;

            case PropertyGridType.Settings:
                grid = WebPropertyGrid.Reflect(Settings.Current);
                grid.Options.Title = Res.Settings;
                break;

            case PropertyGridType.InstanceSettings:
                grid = WebPropertyGrid.Reflect(ShellBatInstance.Current.Settings);
                grid.Options.Title = Res.InstanceSettings;
                break;

            case PropertyGridType.KeyboardShortcuts:
                grid = WebCommand.BuildPropertyGrid();
                grid.Options.Title = Res.KeyboardShortcuts;
                break;

            default:
                throw new NotSupportedException();
        }

        if (!grid.Options.IsReadOnly)
        {
            grid.Saving += (s, e) =>
            {
                switch (type)
                {
                    case PropertyGridType.Settings:
                        if (grid.ApplyChanges((Settings)grid.InstanceObject!))
                        {
                            Settings.Current.SerializeToConfiguration();
                        }
                        break;

                    case PropertyGridType.InstanceSettings:
                        if (grid.ApplyChanges((ShellBatInstanceSettings)grid.InstanceObject!))
                        {
                            ShellBatInstance.Current.SerializeSettings();
                        }
                        break;

                    case PropertyGridType.KeyboardShortcuts:
                        foreach (var prop in grid.Instance.ChangedPropertyNames)
                        {
                            try
                            {
                                var json = grid.Instance.GetNullifiedString(prop);
                                if (json != null)
                                {
                                    var key = JsonSerializer.Deserialize(json, JsonSourceGenerationContext.Default.WebCommandKey);
                                    if (key != null && key.CommandName != null)
                                    {
                                        ShellBatCommand.SetCommandKey(key);
                                    }
                                }
                            }
                            catch
                            {
                                // continue
                            }
                        }
                        break;
                }
            };
        }

        grid.Options.GroupByCategory = true;
        grid.Options.SwalClassName = "fld-swal-50";
        grid.Options.BaseClassName = "fld-pg";
        return grid;
    }

    public string? ValidateCommandKey(object? commandKey)
    {
        if (commandKey is not DirectN.IDispatch disp)
            return null;

        var evt = disp.AsKeyEventArgs();
        if (evt == null)
            return null;

        var commandName = disp.GetProperty<string?>("commandName");
        if (commandName == null)
            return null;

        var key = ShellBatCommand.Validate(window.Handle, evt, commandName);
        return JsonSerializer.Serialize(key, JsonSourceGenerationContext.Default.WebCommandKey);
    }

    public WebPropertyGrid GetInstanceSettingsPropertyGrid()
    {
        var grid = WebPropertyGrid.Reflect(ShellBatInstance.Current.Settings);
        grid.Options.Title = Res.InstanceSettings;
        grid.Options.GroupByCategory = true;
        grid.Options.SwalClassName = "fld-swal-50";
        grid.Options.BaseClassName = "fld-pg";
        return grid;
    }

    public WebHistoryEntry? GetFirstHistoryEntry() => GetHistory().FirstOrDefault();
    public nint GetWindowHandle() => window.Handle.Value;
    public string GetSettings() => JsonSerializer.Serialize(Settings.Current, JsonSourceNoDefaultGenerationContext.Default.Settings);
    public string GetInstanceSettings()
    {
        var settings = ShellBatInstance.Current.Settings;
        return JsonSerializer.Serialize(settings, JsonSourceNoDefaultGenerationContext.Default.ShellBatInstanceSettings);
    }

    public void CreateNewFolder() => window.CreateNewFolder();

    public void GCCollect()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        window.ShowSuccess(new() { Text = Res.GCCompleted });
    }

    public void ClearCaches()
    {
        window.ClearCaches(CacheTypes.All);
        window.ShowSuccess(new() { Text = Res.CacheCleared });
    }

    public void DeleteHistory()
    {
        var count = History.Current.Clear();
        if (count > 0)
        {
            History.Current.SerializeToConfiguration();
        }
        window.ShowSuccess(new() { Text = string.Format(Res.HistoryDeleted, count) });
    }

    public void SaveTheme() => window.SaveTheme();
    public void LoadTheme(string? filePath) => window.LoadTheme(filePath);
    public void RefreshCurrentTheme() => window.RefreshCurrentTheme();

    public string GetAboutHtml()
    {
        var url = "https://github.com/smourier/ShellBat";
        var asm = Assembly.GetExecutingAssembly();
        return $"<p style='font-family:monospace'>{Program.AppId} V{asm.GetInformationalVersion()}<br/>"
            + $"{asm.GetCopyright()}<br/><br/>"
            + "<a center href=\"https://stackoverflow.com/users/403671/simon-mourier?tab=profile\"><img src=\"https://stackexchange.com/users/flair/174359.png\" width=\"208\" height=\"58\" alt=\"Profile for Simon Mourier at Stack Overflow, Q&amp;A for professional and enthusiast programmers\" title=\"profile for Simon Mourier at Stack Overflow, Q&amp;A for professional and enthusiast programmers\"></a>"
            + "<br/>"
            + $"{Res.SourceCode}:<br/><a href='{url}'>{url}</a></p>";
    }

    public void RegisterShellIntegrations()
    {
        ShellBatInstance.RegisterComObjects(Registry.CurrentUser);
        window.ShowSuccess(new() { Text = Res.ShellIntegrationsRegistered });
    }

    public void UnregisterShellIntegrations()
    {
        ShellBatInstance.UnregisterComObjects(Registry.CurrentUser);
        window.ShowSuccess(new() { Text = Res.ShellIntegrationsUnregistered });
    }

    public void StopSearch(string? queryId) => window.StopSearch(queryId);
    public void StartSearch(SearchType type, string queryId, string query, object? options) => window.StartSearch(type, queryId, query, options);

    public WebWindow? GetSearchWindow(SearchType type) => window.GetWindow(type.ToString(), null);
    public WebWindow? GetWindow(string id, string? parameters) => window.GetWindow(id, parameters);
    public void RunTerminal(string terminalKey, string? commandLine) => window.RunTerminal(terminalKey, new WebWindowShowOptions { CommandLine = commandLine });
    public WebTerminal? GetTerminal(string id, string key) => window.GetTerminal(id, key);
    public WebActions GetActions() => window.GetActions();
    public WebContextMenu? GetContextMenu(object? parsingNames)
    {
        if (parsingNames is not DirectN.IDispatch disp)
            return null;

        var list = disp.GetProperty<object[]?>("parsingNames");
        if (list == null || list.Length == 0)
        {
            var current = window.CurrentEntry;
            if (current is null)
                return null;

            return window.GetContextMenu([current.ParsingName]);
        }

        return window.GetContextMenu([.. list.OfType<string>().WhereNotNull()]);
    }

    public string? RenameEntry(string parsingName, string newName)
    {
        ArgumentNullException.ThrowIfNull(parsingName);
        ArgumentNullException.ThrowIfNull(newName);
        var entry = Entry.Get(null, parsingName, ShellItemParsingOptions.DontThrowOnError);
        if (entry is null)
            return DirectN.Constants.E_FAIL.GetException()?.Message;

        var hr = entry.Rename(newName, out var renamedEntry, _TRANSFER_SOURCE_FLAGS.TSF_NORMAL, false);
        if (hr.IsError)
            return hr.GetException()?.Message;

        window.RefreshEntries(renamedEntry?.ParsingName);
        return null;
    }

    public void ExecuteAction(string actionId, object? parsingNames)
    {
        if (parsingNames is not DirectN.IDispatch disp)
            return;

        var list = disp.GetProperty<object[]?>("parsingNames");
        if (list == null || list.Length == 0)
        {
            var current = window.CurrentEntry;
            if (current is null)
                return;

            window.ExecuteAction([current.ParsingName], actionId);
            return;
        }

        window.ExecuteAction(list.Select(l => l as string), actionId);
    }

    public IEnumerable<WebCommand> GetCommands() => WebCommand.GetCommands();
    public IEnumerable<WebTheme> GetAvailableThemes() => ShellBatTheme.AvailableThemes.Select(t => new WebTheme(t)).OrderBy(t => t.DisplayName);
    public IEnumerable<WebScreen> GetScreens() => WebScreen.All;
    public IEnumerable<WebShellBatInstance> GetInstances() => ShellBatInstance.Current.All.Where(i => !i.IsInHttpServerOnlyMode).Select(i => new WebShellBatInstance(i));
    public IEnumerable<WebHistoryEntry> GetFavorites() => Settings.Current.Favorites.Select(e => e.ToWebHistoryEntry());
    public IEnumerable<WebHistoryEntry> GetHistory() => History.Current.Entries.Select(e => e.ToWebHistoryEntry());

    public string GetResources(string? language)
    {
        CultureInfo culture;
        language = language.Nullify();
        if (language == null)
        {
            culture = CultureInfo.CurrentUICulture;
        }
        else
        {
            if (int.TryParse(language, out var lcid))
            {
                culture = new CultureInfo(lcid);
            }
            else
            {
                try
                {
                    culture = new CultureInfo(language);
                }
                catch
                {
                    culture = CultureInfo.CurrentUICulture;
                }
            }
        }

        var set = Res.ResourceManager.GetResourceSet(culture, true, true);
        if (set == null)
            return "{ }";

        var dic = new Dictionary<string, string?>();
        foreach (var kv in set.Cast<DictionaryEntry>())
        {
            dic[kv.Key.ToString()!] = kv.Value?.ToString();
        }

        return JsonSerializer.Serialize(dic, JsonSourceGenerationContext.Default.IDictionaryStringString);
    }

    public void SetSetting(string name, object? value)
    {
        ArgumentNullException.ThrowIfNull(name);
        var prop = typeof(Settings).GetProperty(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
        if (prop == null || !prop.CanWrite)
            throw new ShellBatException($"0004: Setting '{name}' was not found.");

        if (Equals(prop.GetValue(ShellBatInstance.Current.Settings), value))
            return;

        if (Conversions.TryChangeObjectType(value, prop.PropertyType, out var converted))
        {
            value = converted;
            prop.SetValue(Settings.Current, value);
            Settings.Current.SerializeToConfiguration();
            return;
        }
        throw new ShellBatException($"0005: Value '{value}' for setting '{name}' is not valid.");
    }

    public void SetInstanceSetting(string name, object? value)
    {
        ArgumentNullException.ThrowIfNull(name);
        var prop = typeof(ShellBatInstanceSettings).GetProperty(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
        if (prop == null || !prop.CanWrite)
            throw new ShellBatException($"0012: Instance setting '{name}' was not found.");

        if (Equals(prop.GetValue(ShellBatInstance.Current.Settings), value))
            return;

        if (Conversions.TryChangeObjectType(value, prop.PropertyType, out var converted))
        {
            value = converted;
            prop.SetValue(ShellBatInstance.Current.Settings, value);
            ShellBatInstance.Current.SerializeSettings();
            return;
        }
        throw new ShellBatException($"0013: Value '{value}' for instance setting '{name}' is not valid.");
    }

    public void OnError(string? message, string? source, int line, int col, string? stack) => Error?.Invoke(this, new WebErrorEventArgs(message, source, line, col, stack));

    public object? SendEvent(WebEventType type, object? obj)
    {
        var e = new WebEventArgs(type, obj as DirectN.IDispatch);
        Event?.Invoke(this, e);
        return e.Output;
    }

    protected override object? GetTaskResult(Task task)
    {
        if (task is Task<bool> boolTask)
            return boolTask.Result;

        return base.GetTaskResult(task);
    }
#pragma warning restore CA1822 // Mark members as static
}
