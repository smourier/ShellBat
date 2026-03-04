namespace ShellBat;

[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)]
public sealed partial class ShellBatInstanceSettings : JsonBasedSettings<ShellBatInstanceSettings>
{
    public const int DefaultThumbnailsSize = 96;
    public const int DefaultDockPadding = 10;
    public const int DefaultPaging = 10;

    internal RECT? GetInitialRect()
    {
        if (MainPlacement is null)
            return null;

        var pos = MainPlacement.Value.rcNormalPosition;
        // check pos is within any monitor
        if (WebScreen.All.Any(m => m.WorkingArea.IntersectsWith(pos)))
            return pos;

        return null;
    }

    [Browsable(false)]
    public WINDOWPLACEMENT? MainPlacement { get => GetPropertyValue<WINDOWPLACEMENT?>(); set => SetPropertyValue(value); }

    [Browsable(false)]
    public WINDOWPLACEMENT? FullScreenPrevPlacement { get => GetPropertyValue<WINDOWPLACEMENT?>(); set => SetPropertyValue(value); }

    [WebPropertyGridProperty(Windows11Only = true)]
    [LocalizedCategory(Settings.AppearanceCategoryName)]
    [DefaultValue(DWM_WINDOW_CORNER_PREFERENCE.DWMWCP_ROUND)]
    public DWM_WINDOW_CORNER_PREFERENCE WindowCorner
    {
        get => GetPropertyValue(DWM_WINDOW_CORNER_PREFERENCE.DWMWCP_ROUND);
        set
        {
            if (SetPropertyValue(value))
            {
                Program.MainWindow?.SetCorner(value);
            }
        }
    }

    [WebPropertyGridProperty(NeedsRestart = true)]
    [LocalizedCategory(Settings.AppearanceCategoryName)]
    [DefaultValue(COREWEBVIEW2_SCROLLBAR_STYLE.COREWEBVIEW2_SCROLLBAR_STYLE_DEFAULT)]
    public COREWEBVIEW2_SCROLLBAR_STYLE ScrollbarStyle { get => GetPropertyValue(COREWEBVIEW2_SCROLLBAR_STYLE.COREWEBVIEW2_SCROLLBAR_STYLE_DEFAULT); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.AppearanceCategoryName)]
    [ReadOnly(true)]
    public string? ThemeFilePath
    {
        get => GetNullifiedValue();
        set
        {
            value = value.Nullify();
            if (value != null)
            {
                // can be absolute or relative (to config, process) paths
                // if same directory as themes path (modulo app's version), convert to relative
                var dir = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetFullPath(value))); // one level up for version
                var tpDir = Path.GetDirectoryName(Settings.ThemesPath);
                if (dir.EqualsIgnoreCase(tpDir))
                {
                    value = Path.GetFileName(value);
                }
            }

            SetPropertyValue(value);
        }
    }

    [LocalizedCategory(Settings.DiagnosticsCategoryName)]
    [DefaultValue(TraceLevel.Verbose)]
    public TraceLevel LogLevel { get => GetPropertyValue(TraceLevel.Verbose); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.DiagnosticsCategoryName)]
    [DefaultValue(true)]
    public bool LogToFile { get => GetPropertyValue(true); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.DiagnosticsCategoryName)]
#if !DEBUG
    [Browsable(false)]
#endif
    public bool? ShowDevTools { get => GetPropertyValue<bool?>(null); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.BrowsingCategoryName)]
    [DefaultValue(true)]
    public bool NotifyOnFolderEvents { get => GetPropertyValue(true); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.BrowsingCategoryName)]
    [DefaultValue(ViewBy.Details)]
    public ViewBy ViewBy { get => GetPropertyValue(ViewBy.Details); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.BrowsingCategoryName)]
    [DefaultValue(SortBy.Name)]
    public SortBy SortBy { get => GetPropertyValue(SortBy.Name); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.BrowsingCategoryName)]
    [DefaultValue(SortDirection.Ascending)]
    public SortDirection SortDirection { get => GetPropertyValue(SortDirection.Ascending); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.BrowsingCategoryName)]
    [DefaultValue(EntryEnumerateOptions.None)]
    public EntryEnumerateOptions EntryEnumerateOptions { get => GetPropertyValue(EntryEnumerateOptions.None); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.BrowsingCategoryName)]
    [DefaultValue(ViewByDetailsOptions.None)]
    public ViewByDetailsOptions ViewByDetailsOptions { get => GetPropertyValue(ViewByDetailsOptions.None); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.BrowsingCategoryName)]
    [DefaultValue(ViewByImageOptions.DisplayTitle)]
    public ViewByImageOptions ViewByImageOptions { get => GetPropertyValue(ViewByImageOptions.DisplayTitle); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.BrowsingCategoryName)]
    [DefaultValue(DefaultThumbnailsSize)]
    public int ThumbnailsSize { get => GetPropertyValue(DefaultThumbnailsSize); set => SetPropertyValue(value); }

    [DefaultValue(DefaultPaging)]
    [Browsable(false)]
    public int Paging { get => GetPropertyValue(DefaultPaging); set => SetPropertyValue(value); } // used by js

    [LocalizedCategory(Settings.BrowsingCategoryName)]
    public string? DateFormat { get => GetPropertyValue<string?>(); set => SetPropertyValue(value); }

    [Browsable(false)]
    public string? InstanceName { get => GetNullifiedValue(); set => SetPropertyValue(value.Nullify()); }

    [Browsable(false)]
    public string? WebView2UserDataPath { get => GetNullifiedValue(); set => SetPropertyValue(value.Nullify()); }

    [LocalizedCategory(Settings.InstancingCategoryName)] // hidden if unspecified instance
    [DefaultValue(false)]
    public bool IsInstanceScopedHistory
    {
        get => GetPropertyValue(false);
        set
        {
            if (SetPropertyValue(value))
            {
                ShellBatInstance.Current?.SetHistory();
            }
        }
    }

    [Browsable(false)]
    public IList<ShellBatChildWindow>? ChildWindows { get => GetPropertyValue<IList<ShellBatChildWindow>>() ?? []; set => SetPropertyValue(value); }

    [DefaultValue(DefaultDockPadding)]
    [Browsable(false)]
    public int DockPadding { get => GetPropertyValue(DefaultDockPadding); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.ViewersCategoryName)]
    [DefaultValue(true)]
    public bool ImageFillWidth { get => GetPropertyValue(true); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.ViewersCategoryName)]
    [DefaultValue(false)]
    public bool ImageFillHeight { get => GetPropertyValue(false); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.ViewersCategoryName)]
    [DefaultValue(false)]
    public bool VideoAutoPlay { get => GetPropertyValue(false); set => SetPropertyValue(value); }

    [LocalizedCategory(Settings.ViewersCategoryName)]
    [DefaultValue(false)]
    public bool AudioAutoPlay { get => GetPropertyValue(false); set => SetPropertyValue(value); }
}
