namespace ShellBat;

[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)]
public sealed partial class Settings : JsonBasedSettings<Settings>
{
    public const string GlobalFileName = "globalSettings.json";
    public const string Index = "Index.html";
    public const string CacheDirectoryName = "Cache";
    public static string ConfigurationDirectoryPath { get; }
    public static string WebView2UserDataPath { get; internal set; }
    public static string MonacoPath { get; }
    public static string WebRootPath { get; }
    public static string ThemesPath { get; }
    public static string HttpServerCachePath { get; internal set; }
    public static string ConfigurationFilePath { get; }
    public static string TempDirectoryPath { get; }
    public static Settings Current { get; }

    public const string InstancingCategoryName = "Instancing"; //  properties hidden if unspecified instance
    public const string DiagnosticsCategoryName = "Diagnostics";
    public const string AppearanceCategoryName = "Appearance";
    public const string BrowsingCategoryName = "Browsing";
    public const string WindowingCategoryName = "Windowing";
    public const string ViewersCategoryName = "Viewers";
    public const string TerminalsCategoryName = "Terminals";

    static Settings()
    {
        // configuration files are stored in %localappdata%
        ConfigurationDirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), typeof(Settings).Namespace!);

        // data is stored in user's Documents
        ConfigurationFilePath = Path.Combine(ConfigurationDirectoryPath, GlobalFileName);

        TempDirectoryPath = Path.Combine(Path.GetTempPath(), typeof(Settings).Namespace!);
        WebView2UserDataPath = Path.Combine(ConfigurationDirectoryPath, "WebView2");
        HttpServerCachePath = Path.Combine(WebView2UserDataPath, CacheDirectoryName);
        WebRootPath = Path.Combine(ConfigurationDirectoryPath, "WebRoot", Assembly.GetExecutingAssembly().GetInformationalVersion()!);
        ThemesPath = Path.Combine(ConfigurationDirectoryPath, ShellBatTheme.ThemesDirectoryName, Assembly.GetExecutingAssembly().GetInformationalVersion()!);
        MonacoPath = Path.Combine(WebRootPath, "Monaco");

        // build settings
        Settings? settings = null;
        if (IOUtilities.PathIsFile(ConfigurationFilePath))
        {
            settings = Deserialize(ConfigurationFilePath, JsonSourceNoDefaultGenerationContext.Default, true);
        }

        Current = settings ?? new Settings();
        Current.FilePath = ConfigurationFilePath;

        // delete various caches
        Task.Run(PreviewViewer.DeleteTempFiles);
    }

    [Browsable(false)]
    public string? LastHttpServerPorts { get => GetPropertyValue<string?>(); set => SetPropertyValue(value); }

    [Browsable(false)]
    public IList<string>? WebView2NativeImageExtensions { get => GetPropertyValue<IList<string>?>(); set => SetPropertyValue(value); }

    [Browsable(false)]
    public IList<int> AllowedThumbnailsSizes
    {
        get
        {
            var sizes = GetPropertyValue<IList<int>?>();
            if (sizes == null || sizes.Count == 0)
                return [10, 16, 20, 30, 32, 40, 48, 50, 60, 64, 70, 80, 90, 96, 100, 110, 125, 128, 150, 175, 200, 250, 256, 300, 400, 500, 512];

            return sizes;
        }
        set => SetPropertyValue(value);
    }

    [DefaultValue(true)]
    [LocalizedCategory(DiagnosticsCategoryName)]
    public bool SendErrorReport { get => GetPropertyValue(true); set => SetPropertyValue(value); }

    [DefaultValue(10)]
    [LocalizedCategory(InstancingCategoryName)]
    public int MaximumInstances { get => GetPropertyValue(10); set => SetPropertyValue(value); }

    [DefaultValue(5000)]
    [LocalizedCategory(InstancingCategoryName)]
    public int InstancesUpdateInterval { get => GetPropertyValue(5000); set => SetPropertyValue(value); }

    [LocalizedCategory(InstancingCategoryName)]
    [DefaultValue(30)]
    public int BackupsMaxDays { get => GetPropertyValue(30); set => SetPropertyValue(value); }

    [DefaultValue(100)]
    [LocalizedCategory(BrowsingCategoryName)]
    public int MaximumHistoryEntries { get => GetPropertyValue(100); set => SetPropertyValue(value); }

    [DefaultValue(1200u)]
    [LocalizedCategory(ViewersCategoryName)]
    public uint PdfSaveImageSize { get => GetPropertyValue(1200u); set => SetPropertyValue(value); }

    [DefaultValue(PdfViewer.SaveDefaultFormat)]
    [LocalizedCategory(ViewersCategoryName)]
    [WebPropertyGridProperty(EnumProvider = EnumProvider.ImageFormat)]
    public string? PdfSaveDefaultFormat { get => GetNullifiedValue() ?? PdfViewer.SaveDefaultFormat; set => SetPropertyValue(value); }

    [Browsable(false)]
    public IList<Favorite> Favorites { get => GetPropertyValue<IList<Favorite>>(null) ?? []; set => SetPropertyValue(value); }

    [Browsable(false)]
    public IDictionary<string, string> Shortcuts { get => GetPropertyValue<IDictionary<string, string>>(null) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); set => SetPropertyValue(value); }

    [DefaultValue(false)]
    [LocalizedCategory(TerminalsCategoryName)]
    public bool CloseExitedTerminalProcesses { get => GetPropertyValue(false); set => SetPropertyValue(value); }

    [Browsable(false)]
    public IEnumerable<TerminalEntry> Terminals { get => GetPropertyValue<IEnumerable<TerminalEntry>>(null) ?? []; set => SetPropertyValue(value); }

    [Browsable(false)]
    public IEnumerable<LanguageExtensionAlias> LanguageExtensionAliases { get => GetPropertyValue<IEnumerable<LanguageExtensionAlias>>(null) ?? []; set => SetPropertyValue(value); }

    public void SerializeToConfiguration() => Serialize(JsonSourceNoDefaultGenerationContext.Default);
    public void Backup() => Backup(TimeSpan.FromDays(Current.BackupsMaxDays));

    public IEnumerable<TerminalEntry> GetTerminals()
    {
        var terminals = Terminals.ToList();
        if (terminals.Count == 0)
        {
            terminals.Add(new TerminalEntry
            {
                Key = "pwsh",
                DisplayName = Res.Powershell,
                CommandLine = "pwsh.exe -NoExit -Command \"Set-Location -LiteralPath '{0}'\"",
                ChangeDirectoryCommand = "Set-Location -LiteralPath '{0}'",
            });
            terminals.Add(new TerminalEntry
            {
                Key = "powershell",
                DisplayName = Res.WindowsPowershell,
                CommandLine = "powershell.exe -NoExit -Command \"Set-Location -LiteralPath '{0}'\"",
                ChangeDirectoryCommand = "Set-Location -LiteralPath '{0}'",
            });
            terminals.Add(new TerminalEntry
            {
                Key = "cmd",
                DisplayName = Res.CommandPrompt,
                CommandLine = "cmd.exe /k \"cd /d {0}\"",
                SupportsShellBatSync = true,
                SupportsCmdNotFound = true,
                ChangeDirectoryCommand = "cd /d {0}",
            });
            terminals.Add(new TerminalEntry
            {
                Key = "bash",
                DisplayName = Res.Bash,
                IsWsl = true,
                CommandLine = "%windir%\\system32\\bash.exe",
                ChangeDirectoryCommand = "cd '{0}'",
            });
        }
        return terminals;
    }

    public bool RemoveDeletedFavorites()
    {
        var changed = false;
        foreach (var fav in Favorites.ToArray())
        {
            if (Entry.Get(null, fav.ParsingName, ShellItemParsingOptions.DontThrowOnError) == null && Favorites.Remove(fav))
            {
                changed = true;
            }
        }
        return changed;
    }

    public int GetFavoritesIndex(Entry? entry)
    {
        if (entry is null)
            return -1;

        return Favorites.IndexOf(f => f.ParsingName.EqualsIgnoreCase(entry.ParsingName));
    }

    public async Task<bool> ToggleFavorite(Entry? entry)
    {
        if (entry is null)
            return false;

        var index = GetFavoritesIndex(entry);
        if (index >= 0)
        {
            Favorites.RemoveAt(index);
        }
        else
        {
            var fav = new Favorite
            {
                ParsingName = entry.ParsingName,
                DisplayName = entry.DisplayName,
                IconPath = await ShellBatHttpLocalServer.GetImageCacheFilePath(UrlType.Icon, entry.ParsingName, 16)
            };
            Favorites.Add(fav);
        }

        SerializeToConfiguration();
        return true;
    }
}
