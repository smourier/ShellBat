export class WebEventType {
    static Unspecified = 0;
    static CaptionSizeChanged = 1;
    static WindowNotOpen = 2;
    static VisualViewportChanged = 3;
    static Close = 4;
    static Minimize = 5;
    static MaximizeRestore = 6;
    static Navigate = 7;
    static ConsoleLog = 8;
    static OnSearchDoubleClicked = 9;
    static RemoveDeletedItemsFromFavorites = 10;
    static KeyDown = 11;
    static KeyUp = 12;
    static MoveHistoryBack = 13;
    static MoveHistoryForward = 14;
    static Log = 15;
    static CaptionButtonsSizeChanged = 16;
    static OpenNewInstanceAdministrator = 17;
    static EntryDoubleClicked = 18;
    static UpdateNow = 19;
    static ArrangeInstances = 20;
    static SetGlobalEvent = 21;
    static OpenLogsFolder = 22;
    static RemoveDeletedItemsFromHistory = 23;
    static ExportAsCsv = 24;
    static ExportExtensionsAsCsv = 25;
    static DisposeChildWindows = 26;
    static OpenNewInstanceOnScreen = 27;
    static TerminalKeyEvent = 28;
    static ReadyStateChange = 29;
    static OpenNewInstance = 30;
    static QuitAllInstances = 31;
    static SwitchToInstance = 32;
    static QuitInstance = 33;
}

export class MenuId {
    static AppView = 0;
    static AppHistory = 1;
    static AppFavorites = 2;
    static AppInfo = 3;
    static AppInstance = 4;
    static AppThumbnailSize = 5;
    static AppSort = 6;
    static AppActions = 7;

    static MaxId = 8;
}

export class PropertyGridType {
    static Info = 0;
    static Settings = 1;
    static InstanceSettings = 2;
    static KeyboardShortcuts = 3;
}

export class SearchType {
    static FindStrings = 0;
    static WindowsSearch = 1;
}

export class SortBy {
    static Unspecified = 0;
    static Name = 1;
    static Size = 2;
    static DateModified = 3;
    static Extension = 4;
}

export class ViewBy {
    static Unspecified = 0;
    static Details = 1;
    static Images = 2;
}

export class ViewByImageOptions {
    static None = 0x0;
    static DisplayTitle = 0x1;
    static RenderPdfThumbnails = 0x2;
    static SquareThumbnails = 0x4;
}

export class ViewByDetailsOptions {
    static None = 0x0;
    static ShowIcons = 0x1;
    static ShowThumbnails = 0x2;
}

export class SortDirection {
    static Unspecified = 0;
    static Ascending = 1;
    static Descending = 2;
}

export class WebEntryOptions {
    static None = 0x0;
    static IsFolder = 0x1;
    static IsNotWebView2NativeImage = 0x2;
    static IsCompressed = 0x4;
    static IsImage = 0x8;
    static IsHidden = 0x10;
    static IsSystem = 0x20;
    static IsPdf = 0x40;
    static IsCut = 0x80;
}

export class EntryEnumerateOptions {
    static None = 0x0;
    static IncludeHidden = 0x1;
    static IncludeSystem = 0x2;
    static ExcludeFolders = 0x4;
    static ExcludeFiles = 0x8;
    static ShowCompressedAsFile = 0x10;
    static DontInvokeDefaultCommand = 0x20;
    static ImagesOnly = 0x40;
    static DontUseCache = 0x80;
}

export class EntryViewType {
    static Details = 0x0;
    static Images = 0x1;
}

export class WebWindowStyles {
    static IsCloseable = 0x1;
    static IsDraggable = 0x2;
    static IsResizeable = 0x3;
}

export class EditorControlEventType {
    static Unknown = 0;
    static Load = 1;
    static ContentChanged = 2;
    static KeyUp = 3;
    static KeyDown = 4;
    static EditorCreated = 5;
    static EditorLostFocus = 6;
    static EditorGotFocus = 7;
    static CursorPositionChanged = 8;
    static CursorSelectionChanged = 9;
    static ModelLanguageChanged = 10;
    static ConfigurationChanged = 11;
    static OpenResource = 12;
    static Paste = 13;
}

export class BorderSnap {
    static None = 0;   
    static Left = 1;
    static Top = 2;
    static Right = 3;
    static Bottom = 4;
    static LeftTop = 5;
    static RightTop = 6;
    static LeftBottom = 7;
    static RightBottom = 8;
}
