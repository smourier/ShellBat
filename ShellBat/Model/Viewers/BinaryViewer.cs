namespace ShellBat.Model.Viewers;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class BinaryViewer(Entry entry) : Viewer(entry)
{
    private Stream? _stream;
    private MemoryStream? _memoryStream;
    private int _bytesPerRow = 16;

    public override string? Icon => "fa-solid fa-square-binary";
    public override int SortOrder => int.MaxValue;
    public override bool IsSupported => !Entry.IsFolder && Entry.Size > 0;

    public virtual bool Lowercase { get; set; } = false;
    public virtual bool HighlightCrLf { get; set; } = false;
    public int TotalLines => (int)((Entry.Size + _bytesPerRow - 1) / _bytesPerRow);
    public int BytesPerRow
    {
        get => _bytesPerRow;
        set
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(value, 8);
            _bytesPerRow = value;
        }
    }

    public override Task<WebPropertyGrid?> GetProperties()
    {
        var grid = new WebPropertyGrid();

        grid.Properties.Add(new WebPropertyGridProperty()
        {
            Key = nameof(Lowercase),
            DisplayName = Res.Lowercase,
            Editor = "Boolean",
        });

        grid.Properties.Add(new WebPropertyGridProperty()
        {
            Key = nameof(HighlightCrLf),
            DisplayName = Res.HighlightCrLf,
            Editor = "Boolean",
        });

        grid.Properties.Add(new WebPropertyGridProperty()
        {
            Key = nameof(BytesPerRow),
            DisplayName = Res.BytesPerRow,
            Editor = "Enum",
            EnumValues = new Dictionary<string, object?>()
            {
                { "8", 8 },
                { "16", 16 },
                { "24", 24 },
                { "32", 32 },
                { "64", 64 }, // Baby I’m a Star
                { "128", 128 }, // Let's Go Crazy
            }
        });

        grid.Instance.PropertyChanged += (s, e) =>
        {
            switch (e.PropertyName)
            {
                case nameof(HighlightCrLf):
                    HighlightCrLf = grid.Instance.GetTypedValue<bool>(nameof(HighlightCrLf));
                    break;

                case nameof(BytesPerRow):
                    BytesPerRow = grid.Instance.GetTypedValue<int>(nameof(BytesPerRow));
                    break;

                case nameof(Lowercase):
                    Lowercase = grid.Instance.GetTypedValue<bool>(nameof(Lowercase));
                    break;
            }
        };

        grid.Instance.SetValue(nameof(Lowercase), Lowercase);
        grid.Instance.SetValue(nameof(BytesPerRow), BytesPerRow);

        grid.Options.GroupByCategory = false;
        grid.Options.PropertyNamePostfix = ":";
        grid.Options.BaseClassName = "viewer-Binary-pg";
        return Task.FromResult<WebPropertyGrid?>(grid);
    }

    public virtual string? GetText(int startLine, int countLines)
    {
        //Application.TraceVerbose($"BinaryViewer.GetText(startLine={startLine}, countLines={countLines})");
        if (countLines <= 0)
        {
            countLines = 1;
        }

        _stream ??= Entry.OpenRead();
        if (_stream == null)
            return null;

        var position = startLine * _bytesPerRow;
        Stream readableStream;
        if (_stream.CanSeek)
        {
            _stream.Seek(position, SeekOrigin.Begin);
            readableStream = _stream;
        }
        else
        {
            _memoryStream ??= new MemoryStream();
            // read what's missing into memory stream
            var toRead = position - _memoryStream.Length + countLines * _bytesPerRow;
            if (toRead > 0)
            {
                var msBuffer = new byte[81920];
                while (toRead > 0)
                {
                    var read = _stream.Read(msBuffer, 0, (int)Math.Min(msBuffer.Length, toRead));
                    if (read == 0)
                        break;

                    _memoryStream.Write(msBuffer, 0, read);
                    toRead -= read;
                }
            }

            _memoryStream.Seek(position, SeekOrigin.Begin);
            readableStream = _memoryStream;
        }

        var sb = new StringBuilder();

        // 0         1         2         3         4         5         6         7
        // 012345678901234567890123456789012345678901234567890123456789012345678901234
        // Offset    00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F  0123456789ABCDEF");
        // --------  -----------------------------------------------  ----------------");
        const int spacing = 2;
        const int hexSpacing = 1;
        const int offsetLength = 8;
        var length = offsetLength + spacing + (2 + hexSpacing) * _bytesPerRow + spacing + _bytesPerRow;
        var buffer = new StringBuilder(length);
        Span<byte> bytes = stackalloc byte[_bytesPerRow];
        var upper = !Lowercase;
        var crlf = HighlightCrLf;

        for (var line = 0; line < countLines; line++)
        {
            var read = readableStream.Read(bytes);
            if (read == 0)
                break;

            var pos = position.ToString("X8");
            if (!upper)
            {
                pos = pos.ToLowerInvariant();
            }

            buffer.Append(pos);
            for (var sp = 0; sp < spacing; sp++)
            {
                buffer.Append(' ');
            }

            var displayed = Math.Min(read, _bytesPerRow);
            for (var i = 0; i < displayed; i++)
            {
                var b = bytes[i];

                if (crlf && (b == '\r' || b == '\n'))
                {
                    buffer.Append("<span class=crlf>");
                }

                buffer.Append((b >> 4).NibbleToChar(upper));
                buffer.Append(((int)b).NibbleToChar(upper));

                if (crlf && (b == '\r' || b == '\n'))
                {
                    buffer.Append("</span>");
                }
                buffer.Append(' ');
            }

            for (var sp = 0; sp < spacing; sp++)
            {
                buffer.Append(' ');
            }

            var left = _bytesPerRow - read;
            if (left > 0)
            {
                buffer.Append(' ', left * (2 + hexSpacing));
            }

            for (var i = 0; i < displayed; i++)
            {
                var b = bytes[i];
                if (b == '<')
                {
                    buffer.Append("&lt;");
                    continue;
                }

                if (b > 31 && b < 128)
                {
                    buffer.Append((char)b);
                    continue;
                }

                if (crlf && (b == '\r' || b == '\n'))
                {
                    buffer.Append("<span class=crlf>\u2936</span>");
                }
                else
                {
                    buffer.Append('.');
                }
            }

            sb.AppendLine(buffer.ToString());
            position += read;
            buffer.Clear();
        }
        return sb.ToString();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            Interlocked.Exchange(ref _stream, null)?.Dispose();
            Interlocked.Exchange(ref _memoryStream, null)?.Dispose();
        }
        base.Dispose(disposing);
    }
}
