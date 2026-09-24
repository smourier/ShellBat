namespace ShellBat.Monaco;

public static class MonacoExtensions
{
    public static bool LanguagesLoaded { get; private set; }

    private static readonly ConcurrentDictionary<string, LanguageExtensionPoint> _languagesById = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<string, IReadOnlyList<LanguageExtensionPoint>> _languagesByExtension = new(StringComparer.OrdinalIgnoreCase);
    private static bool _loadingLanguages;

    public static async Task LoadLanguages(IComObject<ICoreWebView2> webView)
    {
        ArgumentNullException.ThrowIfNull(webView);
        if (LanguagesLoaded)
            return;

        if (_loadingLanguages)
        {
            do
            {
                await Task.Delay(20);
            }
            while (_loadingLanguages);
            return;
        }

        _loadingLanguages = true;

        var json = await webView.ExecuteScriptAsJson("monaco.languages.getLanguages()");
        var languages = JsonSerializer.Deserialize(json!, JsonSourceGenerationContext.Default.LanguageExtensionPointArray);
        if (languages != null)
        {
            foreach (var language in languages)
            {
                if (language == null || language.Id == null)
                    continue;

                if (language.Id.StartsWith("freemarker2.")) // get rid of these
                    continue;

                _languagesById[language.Id] = language;
                if (language.Extensions != null)
                {
                    foreach (var ext in language.Extensions)
                    {
                        if (!_languagesByExtension.TryGetValue(ext, out var list))
                        {
                            var l = new List<LanguageExtensionPoint>();
                            list = l;
                            _languagesByExtension[ext] = list;
                        }
                        ((List<LanguageExtensionPoint>)list).Add(language);
                    }
                }
            }
        }

        if (!_languagesByExtension.IsEmpty)
        {
            // add some well-known languages that are not recognized by Monaco
            addExtensionLike(".idl", ".c", false);
            addExtensionLike(FileExtension.ShellBatExtensionName, ".xml", false);
            addExtensionLike(".appxmanifest", ".xml", false);
            addExtensionLike(".log", ".txt", false);
            addExtensionLike(".manifest", ".xml", false);
            addExtensionLike(".resx", ".xml", false);
            addExtensionLike(".user", ".xml", false);
            addExtensionLike(".sln", ".txt", false);
            addExtensionLike(".md", ".txt", false);

            foreach (var alias in Settings.Current.LanguageExtensionAliases ?? [])
            {
                if (string.IsNullOrWhiteSpace(alias.Extension) || string.IsNullOrWhiteSpace(alias.LanguageId))
                    continue;

                addExtensionLike(alias.Extension!, alias.LanguageId, alias.Force);
            }

            static void addExtensionLike(string ext, string likeExt, bool force)
            {
                if (!force && _languagesByExtension.ContainsKey(ext))
                    return;

                if (!_languagesByExtension.TryGetValue(likeExt, out var list) || list.Count == 0)
                    return;

                var first = list.FirstOrDefault(l => l.Extensions != null) ?? list[0];
                if (first.Extensions == null)
                {
                    first.Extensions = [ext];
                }
                else
                {
                    var exts = new List<string>(first.Extensions)
                    {
                        ext
                    };
                    first.Extensions = [.. exts];
                }
                _languagesByExtension[ext] = list;
            }
        }

        _loadingLanguages = false;
        LanguagesLoaded = true;
    }

    public static string? GetLanguageName(string id)
    {
        if (!LanguagesLoaded)
            throw new InvalidOperationException();

        if (id == null)
            return null;

        var languages = GetLanguages();
        if (languages == null)
            return null;

        languages.TryGetValue(id, out var lang);
        if (lang != null)
            return lang.Name;

        return null;
    }

    public static string? GetThemeName(this EditorTheme theme) => theme switch
    {
        EditorTheme.VisualStudio => "vs",
        EditorTheme.VisualStudioDark => "vs-dark",
        EditorTheme.HighContrastBlack => "hc-black",
        EditorTheme.HighContrastLight => "hc-light",
        _ => theme.ToString(),
    };

    public static IDictionary<string, IReadOnlyList<LanguageExtensionPoint>> GetLanguagesByExtension()
    {
        if (!LanguagesLoaded)
            throw new InvalidOperationException();

        return _languagesByExtension;
    }

    public static IDictionary<string, LanguageExtensionPoint>? GetLanguages()
    {
        if (!LanguagesLoaded)
            return null;

        return _languagesById;
    }

    public static string? GetLanguageByExtension(string ext, string? defaultLanguageId = null)
    {
        if (string.IsNullOrWhiteSpace(ext) || !LanguagesLoaded)
            return defaultLanguageId;

        if (_languagesByExtension.TryGetValue(ext, out var langs) && langs.Count > 0)
            return langs[0].Id ?? defaultLanguageId;

        using var key = Registry.ClassesRoot.OpenSubKey(ext, false);
        if (key != null)
        {
            var ct = key.GetValue("Content Type") as string;
            if (!string.IsNullOrWhiteSpace(ct))
            {
                using var mime = Registry.ClassesRoot.OpenSubKey(Path.Combine(@"MIME\Database\Content Type", ct), false);
                var mimeExt = (mime?.GetValue("Extension") as string)?.Nullify();
                if (mimeExt != null && _languagesByExtension.TryGetValue(mimeExt, out langs) && langs.Count > 0)
                    return langs[0].Id ?? defaultLanguageId;
            }
        }

        return defaultLanguageId;
    }

    public static bool IsUnknownLanguageExtension(string ext)
    {
        if (string.IsNullOrWhiteSpace(ext))
            return true;

        if (LanguagesLoaded)
        {
            var languages = GetLanguagesByExtension();
            if (languages.TryGetValue(ext, out var list) && list.Count > 0)
                return false;
        }

        using var key = Registry.ClassesRoot.OpenSubKey(ext, false);
        if (key == null)
            return true;

        var ct = key.GetValue("Content Type") as string;
        if (string.IsNullOrWhiteSpace(ct))
            return true;
        return false;
    }
}
