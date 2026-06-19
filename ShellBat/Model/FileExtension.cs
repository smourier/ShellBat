namespace ShellBat.Model;

public class FileExtension : IEquatable<FileExtension>, IComparable<FileExtension>, IComparable
{
    public const string ShellBatExtensionName = ".shellBat";
    public const string ShellBatThemeExtensionName = ".shellBatTheme";

    private static readonly Lazy<ConcurrentDictionary<string, IReadOnlyList<string>>> _mimeTypesByExtension = new(GetMimeTypesByExtension);
    private static readonly ConcurrentDictionary<string, FileExtension> _fileExtensions = new(StringComparer.OrdinalIgnoreCase);

    // https://developer.mozilla.org/en-US/docs/Web/Media/Guides/Formats/Image_types
    public static string[] DefaultWebView2NativeImageExtensions { get; } = [
        ".apng",
        ".png",
        ".avif",
        ".gif",
        ".jpg",
        ".jpeg",
        ".jpe",
        ".jif",
        ".jfif",
        ".pjpeg",
        ".pjp",
        ".png",
        ".svg",
        ".webp",
        ".bmp",
        ".ico",
        ".cur"
    ];

    // https://stackoverflow.com/a/77656328/403671
    public static IReadOnlyList<string> CompressedWellKnownExtensions { get; } =
    [
        // windows built-in compressed types
        ".zip",
        ".7z",
        ".gz",
        ".bz2",
        ".tar",
        ".rar",
        ".tgz",
        ".tbz2",
        ".tzst",
        ".txz",
        ".zst",
        ".xz",
    ];

    public static IReadOnlySet<string> TextWellKnownExtensions { get; } = new HashSet<string>(
    [
        ".appxmanifest",
        ".asax",
        ".ascx",
        ".ashx",
        ".asmx",
        ".asp",
        ".axml",
        ".bas",
        ".bat",
        ".btproj",
        ".cbl",
        ".class",
        ".cmd",
        ".cob",
        ".cs",
        ".cshtml",
        ".css",
        ".cfg",
        ".config",
        ".cbproj",
        ".crproj",
        ".csproj",
        ".dproj",
        ".dbproj",
        ".dbschema",
        ".disco",
        ".deploymanifest",
        ".diagram",
        ".dotsettings",
        ".editorconfig",
        ".edmx",
        ".eml",
        ".frm",
        ".go",
        ".htm",
        ".html",
        ".iqy",
        ".inf",
        ".ini",
        ".isl",
        ".isproj",
        ".java",
        ".js",
        ".json",
        ".l4g",
        ".log",
        ".licx",
        ".master",
        ".manifest",
        ".nuspec",
        ".pas",
        ".package",
        ".pbxproj",
        ".plist",
        ".props",
        ".ps1",
        ".psd1",
        ".psm1",
        ".rdl",
        ".readme",
        ".reg",
        ".resx",
        ".rs",
        ".rtf",
        ".rzt",
        ".schemaview",
        ".sh",
        ".sitemap",
        ".sln",
        ".spdata",
        ".sql",
        ".sqlproj",
        ".sqlcmdvars",
        ".sqldeployment",
        ".sqlsettings",
        ".snippet",
        ".storyboard",
        ".svc",
        ".svcinfo",
        ".svcmap",
        ".svg",
        ".targets",
        ".tcl",
        ".tpl",
        ".tplxaml",
        ".vb",
        ".vbhtml",
        ".vbp",
        ".vbproj",
        ".vbs",
        ".vcproj",
        ".vcxproj",
        ".vdproj",
        ".webpart",
        ".wsdl",
        ".wxi",
        ".wxl",
        ".wxs",
        ".wixlib",
        ".vixproj",
        ".xml",
        ".xaml",
        ".xsd",
        ".xsl",
        ".xslt",
    ]);

    public static IReadOnlyList<string> CanBeDecompressedExtensions { get; } =
    [
        // known zip in disguised formats
        ".apk",
        ".aab",
        ".aar",
        ".appx",
        ".appxbundle",
        ".cbz",
        ".ctx",
        ".docx",
        ".epub",
        ".ipa",
        ".jar",
        ".msix",
        ".msixbundle",
        ".nupkg",
        ".odg",
        ".odp",
        ".ods",
        ".odt",
        ".oxps",
        ".pptx",
        ".svgz",
        ".unitypackage",
        ".xlsx",
        ".xlsm",
        ".xpi",
        ".vsdx",
        ".vsix",
        ".war",
        ".whl",
    ];

    public static FileExtension None { get; } = new(string.Empty);

    private readonly Lazy<PreviewHandlerInfo?> _previewHandlerInfo;
    private readonly Lazy<string?> _displayName;

    private sealed class PreviewHandlerInfo
    {
        public Guid Clsid;
        public string? ProcessCommandLine;
        public string? ProcessCommandLineDiscriminant;
    }

    private FileExtension(string name)
    {
        Name = name.ToLowerInvariant();
        IsExecutable = Name.EqualsIgnoreCase(".exe") || Name.EqualsIgnoreCase(".dll") || Name.EqualsIgnoreCase(".com") || Name.EqualsIgnoreCase(".scr");
        IsShortcut = Name.EqualsIgnoreCase(".lnk");
        IsPdf = Name == ".pdf";
        IsCompressed = CompressedWellKnownExtensions.Contains(Name);
        IsImage = GetIsImage();
        IsDiskImage = Name.EqualsIgnoreCase(".iso") ||
                Name.EqualsIgnoreCase(".img") ||
                Name.EqualsIgnoreCase(".vhd") ||
                Name.EqualsIgnoreCase(".vhdx");

        IsWebView2NativeImage = WebView2NativeImageExtensions.Contains(name);
        IsWicImage = WicImagingComponent.DecoderFileExtensions.Contains(name);

        ShellN.Functions.AssocGetPerceivedType(PWSTR.From(Name), out var perceived, out var flags, 0);
        PerceivedType = perceived;

        IsText = PerceivedType == PERCEIVED.PERCEIVED_TYPE_TEXT || TextWellKnownExtensions.Contains(Name);
        _previewHandlerInfo = new(GetPreviewHandlerInfo);
        _displayName = new(() => GetAssocString(ASSOCSTR.ASSOCSTR_FRIENDLYDOCNAME));
    }

    public string Name { get; }
    public string NameWithoutDot => Name.StartsWith('.') && Name.Length > 1 ? Name[1..] : Name;
    public string? DisplayName => _displayName.Value;
    public PERCEIVED PerceivedType { get; }
    public Guid? PreviewHandlerClsid => _previewHandlerInfo.Value?.Clsid;
    public string? PreviewHandlerProcessCommandLine => _previewHandlerInfo.Value?.ProcessCommandLine;
    public string? PreviewHandlerProcessCommandLineDiscriminant => _previewHandlerInfo.Value?.ProcessCommandLineDiscriminant;
    public bool IsNone => string.IsNullOrEmpty(Name);
    public bool IsWebView2NativeImage { get; private set; } // heic are typically not webview2 native images
    public bool IsWicImage { get; private set; } // svg are typically not wic images
    public bool IsCompressed { get; }
    public bool CanBeDecompressed => CanBeDecompressedExtensions.Contains(Name);
    public bool IsExecutable { get; }
    public bool IsShortcut { get; }
    public bool IsPdf { get; }
    public bool IsImage { get; }
    public bool IsDiskImage { get; }
    public bool IsText { get; }

    public override string ToString() => Name;
    public override int GetHashCode() => Name.GetHashCode();
    public override bool Equals(object? obj) => obj is FileExtension other && Equals(other);
    public bool Equals(FileExtension? other) => other is not null && Name.EqualsIgnoreCase(other.Name);
    int IComparable.CompareTo(object? obj) => CompareTo(obj as FileExtension);
    public int CompareTo(FileExtension? other)
    {
        ArgumentNullException.ThrowIfNull(other);
        return string.Compare(Name, other.Name);
    }

    public AssocHandler? GetAssocHandler(Func<AssocHandler, bool> includes)
    {
        foreach (var handler in GetAssocHandlers())
        {
            if (includes(handler))
                return handler;

            handler.Dispose();
        }
        return null;
    }

    public IReadOnlyList<AssocHandler> GetAssocHandlers() => AssocHandler.Enumerate(Name, ASSOC_FILTER.ASSOC_FILTER_RECOMMENDED);
    public HRESULT TryCreatePreviewHandler(out IComObject<IPreviewHandler>? handler)
    {
        if (PreviewHandlerClsid == null)
        {
            handler = null;
            return DirectN.Constants.S_OK;
        }

        return ComObject.TryCoCreate(PreviewHandlerClsid.Value, out handler, CLSCTX.CLSCTX_LOCAL_SERVER);
    }

    public static WicBitmapSource? GetFolderSmallIcon() => GetFileExtensionSmallIcon(string.Empty);
    public unsafe static WicBitmapSource? GetFileExtensionSmallIcon(string extension)
    {
        ArgumentNullException.ThrowIfNull(extension);

        var flags = SHGFI_FLAGS.SHGFI_ICON | SHGFI_FLAGS.SHGFI_USEFILEATTRIBUTES | SHGFI_FLAGS.SHGFI_SMALLICON;
        var atts = FILE_FLAGS_AND_ATTRIBUTES.FILE_ATTRIBUTE_NORMAL;
        if (extension.Length == 0 || extension.EqualsIgnoreCase(".folder") || extension.EqualsIgnoreCase("folder"))
        {
            extension = "folder";
            atts = FILE_FLAGS_AND_ATTRIBUTES.FILE_ATTRIBUTE_DIRECTORY;
            flags |= SHGFI_FLAGS.SHGFI_OPENICON;
        }

        WicBitmapSource? bitmap = null;
        var info = new SHFILEINFOW();
        ShellN.Functions.SHGetFileInfoW(PWSTR.From(extension), atts, (nint)(&info), (uint)sizeof(SHFILEINFOW), flags);
        if (info.hIcon != 0)
        {
            bitmap = WicBitmapSource.FromHICON(info.hIcon);
            if (bitmap != null)
            {
                if (bitmap.PixelFormat != WicPixelFormat.GUID_WICPixelFormat32bppPBGRA)
                {
                    bitmap.ConvertTo(WicPixelFormat.GUID_WICPixelFormat32bppPBGRA);
                }
            }

            DirectN.Functions.DestroyIcon(info.hIcon);
        }

        return bitmap;
    }

    [return: NotNullIfNotNull(nameof(extension))]
    public static FileExtension? GetByExtension(string? extension)
    {
        if (extension == null)
            return null;

        if (!_fileExtensions.TryGetValue(extension, out var fileExtension))
        {
            fileExtension = new FileExtension(extension);
            _fileExtensions[extension] = fileExtension;
        }
        return fileExtension;
    }

    public static IReadOnlyList<string> GetMimeTypesByExtension(string? extension)
    {
        if (extension == null)
            return [];

        _mimeTypesByExtension.Value.TryGetValue(extension, out var mimeTypes);
        return mimeTypes ?? [];
    }

    public static void ExportAsCsv(string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        IOUtilities.FileEnsureDirectory(filePath);

        var fileExtensions = new List<FileExtension>();
        foreach (var ext in Registry.ClassesRoot.GetSubKeyNames())
        {
            if (ext.Length < 2 || ext[0] != '.')
                continue;

            var fileExtension = GetByExtension(ext);
            if (fileExtension == null)
                continue;

            fileExtensions.Add(fileExtension);
        }

        using var writer = new StreamWriter(filePath, false, Encoding.Unicode);
        var i = 0;
        foreach (var ext in fileExtensions.OrderBy(e => e.Name, StringComparer.OrdinalIgnoreCase))
        {
            if (i == 0)
            {
                WriteHeader(writer);
                writer.WriteLine();
                i++;
            }

            ext.Write(writer);
            writer.WriteLine();
        }
    }

    public static string GetDisplayName(string extension)
    {
        ArgumentNullException.ThrowIfNull(extension);
        var ext = GetByExtension(extension);
        return ext.DisplayName ?? ext.NameWithoutDot;
    }

    private string? GetAssocString(ASSOCSTR str)
    {
        uint size = 0;
        ShellN.Functions.AssocQueryStringW(
            ASSOCF.ASSOCF_NONE,
            str,
            PWSTR.From(Name),
            PWSTR.Null,
            PWSTR.Null,
            ref size);
        if (size == 0)
            return null;

        using var p = new AllocPwstr((size + 1) * 2);
        if (ShellN.Functions.AssocQueryStringW(
            ASSOCF.ASSOCF_NONE,
            str,
            PWSTR.From(Name),
            PWSTR.Null,
            p,
            ref size).IsSuccess)
            return p.ToString();

        return null;
    }

    private bool GetIsImage()
    {
        if (PerceivedType == PERCEIVED.PERCEIVED_TYPE_IMAGE)
            return true;

        if (IsWicImage)
            return true;

        var mimeTypes = GetMimeTypesByExtension(Name);
        if (mimeTypes.Any(t => t.StartsWith("image/", StringComparison.OrdinalIgnoreCase)))
            return true;

        return false;
    }

    private static void WriteHeader(TextWriter writer)
    {
        if (writer == null)
            return;

        writer.WriteCsvCell(nameof(Name));
        writer.WriteCsvCell(nameof(PerceivedType));
        writer.WriteCsvCell(nameof(PreviewHandlerClsid));
        writer.WriteCsvCell(nameof(PreviewHandlerProcessCommandLine));
        writer.WriteCsvCell(nameof(IsWebView2NativeImage));
        writer.WriteCsvCell(nameof(IsCompressed));
        writer.WriteCsvCell(nameof(IsExecutable));
        writer.WriteCsvCell(nameof(IsShortcut));
        writer.WriteCsvCell(nameof(IsImage));
        writer.WriteCsvCell("MonacoLanguageId");
        writer.WriteCsvCell("MimeTypes", false);
    }

    private void Write(TextWriter writer)
    {
        if (writer == null)
            return;

        writer.WriteCsvCell(Name);
        writer.WriteCsvCell(PerceivedType != PERCEIVED.PERCEIVED_TYPE_UNKNOWN ? PerceivedType.ToString() : null);
        writer.WriteCsvCell(PreviewHandlerClsid?.ToString("B"));
        writer.WriteCsvCell(PreviewHandlerProcessCommandLine);
        writer.WriteCsvCell(IsWebView2NativeImage ? "TRUE" : null);
        writer.WriteCsvCell(IsCompressed ? "TRUE" : null);
        writer.WriteCsvCell(IsExecutable ? "TRUE" : null);
        writer.WriteCsvCell(IsShortcut ? "TRUE" : null);
        writer.WriteCsvCell(IsImage ? "TRUE" : null);
        writer.WriteCsvCell(MonacoExtensions.GetLanguageByExtension(Name));

        var mimeTypes = GetMimeTypesByExtension(Name);
        if (mimeTypes.Count > 0)
        {
            writer.WriteCsvCell(string.Join(";", mimeTypes), false);
        }
        else
        {
            writer.WriteCsvCell(null, false);
        }
    }

    private static ConcurrentDictionary<string, IReadOnlyList<string>> GetMimeTypesByExtension()
    {
        var dic = new ConcurrentDictionary<string, IReadOnlyList<string>>(StringComparer.OrdinalIgnoreCase);
        using var mimeKey = Registry.ClassesRoot.OpenSubKey(@"MIME\Database\Content Type", false);
        if (mimeKey != null)
        {
            foreach (var mimeType in mimeKey.GetSubKeyNames())
            {
                using var subKey = mimeKey.OpenSubKey(mimeType, false);
                var extension = (subKey?.GetValue("Extension") as string)?.Nullify();
                if (extension == null)
                    continue;

                if (!dic.TryGetValue(extension, out var existingMimeTypes))
                {
#pragma warning disable IDE0028 // Simplify collection initialization
                    existingMimeTypes = new List<string>();
#pragma warning restore IDE0028 // Simplify collection initialization
                    dic[extension] = existingMimeTypes;
                }

                if (existingMimeTypes.Contains(mimeType))
                    continue;

                ((List<string>)existingMimeTypes).Add(mimeType);
            }
        }
        return dic;
    }

    private static IReadOnlyList<string> WebView2NativeImageExtensions
    {
        get
        {
            var webView2NativeImageExtensions = Settings.Current.WebView2NativeImageExtensions;
            if (webView2NativeImageExtensions == null || webView2NativeImageExtensions.Count == 0)
                return DefaultWebView2NativeImageExtensions;

            return [.. webView2NativeImageExtensions];
        }
    }

    private PreviewHandlerInfo? GetPreviewHandlerInfo()
    {
        var clsid = GetPreviewHandlerClsid(Name);
        if (clsid == null)
            return null;

        var info = new PreviewHandlerInfo
        {
            Clsid = clsid.Value
        };

        // try to get the process command line from the registry
        // so we can kill it if needed (stuck)
        using var handlerKey = Registry.ClassesRoot.OpenSubKey(Path.Combine("CLSID", clsid.Value.ToString("B"), "LocalServer32"), false);
        if (handlerKey != null)
        {
            // note that the command line is already env variables-expanded
            var cmdLine = (handlerKey.GetValue(null) as string).Nullify();
            if (cmdLine != null)
            {
                if (cmdLine.StartsWith('"'))
                {
                    var endQuoteIndex = cmdLine.IndexOf('"', 1);
                    if (endQuoteIndex > 1)
                    {
                        cmdLine = cmdLine[1..endQuoteIndex].Nullify();
                    }
                }

                if (cmdLine != null)
                {
                    var exePos = cmdLine.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
                    if (exePos >= 0)
                    {
                        cmdLine = cmdLine[..(exePos + 4)].Nullify();
                    }

                    if (cmdLine != null)
                    {

                        if (cmdLine.Length > 2 && cmdLine[0] == '"' && cmdLine.EndsWith('"'))
                        {
                            cmdLine = cmdLine[1..^1].Nullify();
                        }

                        info.ProcessCommandLine = cmdLine;
                        info.ProcessCommandLineDiscriminant = " -Embedding";
                    }
                }
            }
        }

        if (info.ProcessCommandLine == null)
        {
            // hack, Acrobat Reader's preview handler CLSID does not have LocalServer32 key it's an inproc dll
            // that ends up loading acrobat.exe so we hardcode it here
            if (IsPdf)
            {
                info.ProcessCommandLine = "acrobat.exe";
                info.ProcessCommandLineDiscriminant = "pdfshell_prev";
            }
        }
        return info;
    }

    private static Guid? GetPreviewHandlerClsid(string extension)
    {
        if (string.IsNullOrEmpty(extension))
            return null;

        uint size = 256;
        using var clsid = new Pwstr(Marshal.AllocCoTaskMem((int)size));
        if (ShellN.Functions.AssocQueryStringW(
            ASSOCF.ASSOCF_INIT_DEFAULTTOSTAR,
            ASSOCSTR.ASSOCSTR_SHELLEXTENSION,
            PWSTR.From(extension),
            PWSTR.From(typeof(IPreviewHandler).GUID.ToString("B")),
            clsid, ref size).IsError)
            return null;

        if (Guid.TryParse(clsid.ToString(), out var guid))
            return guid;

        return null;
    }

    internal static List<string> GetMediaFileExtensions(string type)
    {
        var list = new List<string>();
        foreach (var ext in Registry.ClassesRoot.GetSubKeyNames())
        {
            if (!ext.StartsWith('.'))
                continue;

            using var key = Registry.ClassesRoot.OpenSubKey(ext);
            var contentType = key?.GetValue("Content Type") as string;
            if (!string.IsNullOrWhiteSpace(contentType) && contentType.StartsWith(type + "/", StringComparison.OrdinalIgnoreCase))
            {
                list.Add(ext);
            }
        }

        list.Sort(StringComparer.OrdinalIgnoreCase);
        return list;
    }
}
