namespace ShellBat.Model.Viewers;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class TextViewer(Entry entry) : Viewer(entry)
{
    public static IReadOnlyList<string> WellKnownTextFileNames { get; } = (string[])
    [
        "license",
        "readme",
        "changelog",
        "contributing",
        "changes",
        "news",
        "history",
        "upgrade",
        "install",
        "release",
        "notice",
        "credits",
        "authors",
        "copying",
    ];

    private LinesStream? _stream;

    public override string? Icon => "fa-regular fa-file-lines";
    public override bool IsSupported
    {
        get
        {
            if (Entry.Extension.IsText)
                return true;

            if (WellKnownTextFileNames.Contains(Entry.DisplayName, StringComparer.OrdinalIgnoreCase))
                return true;

            return base.IsSupported;
        }
    }

    public virtual bool HighlightCrLf { get; set; }
    public virtual Wrap Wrap { get; set; }
    public bool IsLoaded => _stream?.IsLoaded ?? true;
    public int TotalLines => _stream?.Lines.Count ?? 0;
    public string? Encoding => _stream?.Encoding?.WebName;

    public override Task<WebPropertyGrid?> GetProperties()
    {
        var grid = new WebPropertyGrid();

        grid.Properties.Add(new WebPropertyGridProperty()
        {
            Key = nameof(LinesStream.Bom),
            DisplayName = "BOM",
            IsReadOnly = true,
        });

        grid.Properties.Add(new WebPropertyGridProperty()
        {
            Key = nameof(LinesStream.Encoding),
            DisplayName = "Encoding",
            IsReadOnly = true,
        });

        grid.Properties.Add(new WebPropertyGridProperty()
        {
            Key = nameof(HighlightCrLf),
            DisplayName = Res.HighlightCrLf,
            Editor = "Boolean",
        });

        grid.Properties.Add(new WebPropertyGridProperty()
        {
            Key = nameof(Wrap),
            DisplayName = Res.Wrap,
            Editor = "Enum",
            EnumValues = Enum.GetValues<Wrap>().ToDictionary(t => Res.ResourceManager.GetString($"Wrap{t}") ?? t.ToString(), t => (object?)t),
        });

        grid.Properties.Add(new WebPropertyGridProperty()
        {
            Key = nameof(Encoding),
            DisplayName = Res.Encoding,
            IsReadOnly = true,
        });

        grid.Instance.PropertyChanged += (s, e) =>
        {
            switch (e.PropertyName)
            {
                case nameof(HighlightCrLf):
                    HighlightCrLf = grid.Instance.GetTypedValue<bool>(nameof(HighlightCrLf));
                    break;

                case nameof(Wrap):
                    var newWrap = grid.Instance.GetTypedValue<Wrap>(nameof(Wrap));
                    if (newWrap != Wrap)
                    {
                        Interlocked.Exchange(ref _stream, null)?.Dispose();
                    }
                    Wrap = newWrap;
                    break;
            }
        };

        grid.Instance.SetValue(nameof(LinesStream.Bom), _stream?.Bom.ToString());
        grid.Instance.SetValue(nameof(LinesStream.Encoding), _stream?.Encoding?.WebName);
        grid.Instance.SetValue(nameof(Encoding), Encoding);
        grid.Instance.SetValue(nameof(Wrap), Wrap);

        grid.Options.GroupByCategory = false;
        grid.Options.PropertyNamePostfix = ":";
        grid.Options.BaseClassName = "viewer-Text-pg";
        return Task.FromResult<WebPropertyGrid?>(grid);
    }

    public virtual string GetText(int startLine, int countLines)
    {
        var expr = GetTexts(startLine, countLines).WhereNotNull();
        return string.Join(HighlightCrLf ? "<span class=crlf>\u2936</span><br/>" : "<br/>", expr.Select(s => s.Replace("<", "&lt;")));
    }

    public virtual IEnumerable<string>? GetTexts(int startLine, int countLines)
    {
        if (countLines <= 0)
        {
            countLines = 1;
        }

        var stream = Entry.OpenRead();
        if (stream == null)
            return null;

        _stream ??= new LinesStream(stream);
        if (_stream == null)
            return null;

        _stream.Wrap = Wrap;
        _stream.Load();

        return _stream.GetTexts(startLine, countLines);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            Interlocked.Exchange(ref _stream, null)?.Dispose();
        }
        base.Dispose(disposing);
    }
}
