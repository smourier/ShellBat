namespace ShellBat.Utilities;

public abstract class JsonBasedSettings<T> : INotifyPropertyChanged, IDisposable, ICloneable where T : new()
{
    public const string BackupDirectoryExtension = ".bak";

    private readonly ConcurrentDictionary<string, object?> _values = new(StringComparer.OrdinalIgnoreCase);

    public event PropertyChangedEventHandler? PropertyChanged;

    private Timer? _serializeTimer;

    [JsonIgnore]
    [Browsable(false)]
    public virtual string? FilePath { get; set; }

    public static T? Deserialize(string filePath, JsonSerializerContext resolver, bool returnDefaultIfError = false)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        ArgumentNullException.ThrowIfNull(resolver);

        if (!IOUtilities.PathIsFile(filePath))
            return new T();

        try
        {
            return FileSystemUtilities.WrapSharingViolations(() =>
            {
                using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return Deserialize(stream, resolver, returnDefaultIfError);
            });
        }
        catch
        {
            if (returnDefaultIfError)
                return default;

            return new T();
        }
    }

    public static T? Deserialize(Stream stream, JsonSerializerContext resolver, bool returnDefaultIfError = false)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(resolver);

        try
        {
            return JsonSerializer.Deserialize(stream, (JsonTypeInfo<T>)resolver.GetTypeInfo(typeof(T))!);
        }
        catch
        {
            if (returnDefaultIfError)
                return default;

            return new T();
        }
    }

    public virtual void Serialize(Stream stream, JsonSerializerContext resolver)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(resolver);
        JsonSerializer.Serialize(stream, this, (JsonTypeInfo<T>)resolver.GetTypeInfo(typeof(T))!);
    }

    public virtual void Serialize(JsonSerializerContext resolver, int deferTime = 300)
    {
        ArgumentNullException.ThrowIfNull(resolver);
        if (FilePath == null)
            throw new InvalidOperationException($"Cannot serialize settings because {nameof(FilePath)} is not set.");

        if (deferTime <= 0)
        {
            serialize();
            return;
        }

        var timer = _serializeTimer;
        if (timer != null)
        {
            timer.Change(deferTime, Timeout.Infinite);
            return;
        }

        _serializeTimer = new Timer(_ =>
        {
            //Application.TraceVerbose($"Serializing instance settings to '{FilePath}'");
            serialize();
            Interlocked.Exchange(ref _serializeTimer, null)?.SafeDispose();
        }, null, deferTime, Timeout.Infinite);

        void serialize()
        {
            var temp = Path.GetTempFileName();
            using (var stream = File.OpenWrite(temp))
            {
                Serialize(stream, resolver);
            }

            FileSystemUtilities.WrapSharingViolations(() =>
            {
                IOUtilities.FileEnsureDirectory(FilePath);
                IOUtilities.FileOverwrite(temp, FilePath);
            });
        }
    }

    protected void OnPropertyChanged(string name) => OnPropertyChanged(this, new PropertyChangedEventArgs(name));
    protected virtual void OnPropertyChanged(object sender, PropertyChangedEventArgs e) => PropertyChanged?.Invoke(sender, e);

    protected string? GetNullifiedValue([CallerMemberName] string? propertyName = null)
    {
        ArgumentNullException.ThrowIfNull(propertyName);
        _values.TryGetValue(propertyName, out var value);
        if (value == null)
            return null;

        return string.Format(CultureInfo.InvariantCulture, "{0}", value).Nullify();
    }

    protected Tv? GetPropertyValue<Tv>(Tv? defaultValue = default, [CallerMemberName] string? propertyName = null)
    {
        if (!TryGetPropertyValue(out var value, propertyName))
            return defaultValue;

        if (!Conversions.TryChangeType<Tv>(value, out var convertedValue))
            return defaultValue;

        return convertedValue;
    }

    protected virtual bool TryGetPropertyValue(out object? value, [CallerMemberName] string? propertyName = null)
    {
        ArgumentNullException.ThrowIfNull(propertyName);
        return _values.TryGetValue(propertyName, out value);
    }

    protected virtual bool SetPropertyValue(object? value, [CallerMemberName] string? propertyName = null)
    {
        ArgumentNullException.ThrowIfNull(propertyName);
        var changed = true;
        _values.AddOrUpdate(propertyName, value, (k, o) =>
        {
            changed = !Equals(value, o);
            return value;
        });

        if (changed)
        {
            OnPropertyChanged(propertyName);
        }
        return changed;
    }

    public virtual void CopyFrom(JsonBasedSettings<T> other)
    {
        ArgumentNullException.ThrowIfNull(other);
        var frozen = other._values.ToArray();
        foreach (var kv in frozen)
        {
            SetPropertyValue(kv.Value, kv.Key);
        }
    }

    object ICloneable.Clone() => Clone()!;
    protected virtual T CreateNew() => new();
    public virtual T Clone()
    {
        var clone = CreateNew();
        ((JsonBasedSettings<T>)(object)clone!).CopyFrom(this);
        return clone;
    }

    public virtual void Backup(TimeSpan maxDuration)
    {
        if (FilePath == null)
            throw new InvalidOperationException($"Cannot backup because {nameof(FilePath)} is not set.");

        try
        {
            BackupPrivate(FilePath, maxDuration);
        }
        catch
        {
            // do nothing, will work next time hopefully...
        }
    }

    private static void BackupPrivate(string filePath, TimeSpan maxDuration)
    {
        if (maxDuration.TotalDays < 0)
            return;

        if (!IOUtilities.PathIsFile(filePath))
            return;

        var fileName = Path.GetFileNameWithoutExtension(filePath);

        // *warning* if you change format here, change parsing code below
        var bakPath = Path.Combine(
            Path.Combine(Path.GetDirectoryName(filePath)!, fileName + BackupDirectoryExtension),
            string.Format(CultureInfo.InvariantCulture, "{0:yyyy}_{0:MM}_{0:dd}.{1}.{2}.json", DateTime.Now, Environment.TickCount, fileName));

        IOUtilities.FileEnsureDirectory(bakPath);
        IOUtilities.FileOverwrite(filePath, bakPath);
        var dir = Path.GetDirectoryName(bakPath)!;
        foreach (var file in Directory.GetFiles(dir))
        {
            if (string.Compare(file, bakPath, StringComparison.OrdinalIgnoreCase) == 0)
                continue;

            var name = Path.GetFileNameWithoutExtension(file);
            var tick = name.IndexOf('.');
            if (tick < 0)
                continue;

            var dates = name[..tick].Split('.');
            if (dates.Length < 1)
                continue;

            var date = dates[0].Replace('_', '/');
            if (!DateTime.TryParse(date, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                continue;

            if (DateTime.Now - dt > maxDuration)
            {
                IOUtilities.FileDelete(file, false);
            }
        }
    }

    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    protected virtual void Dispose(bool disposing)
    {
        var serializeTimer = Interlocked.Exchange(ref _serializeTimer, null);
        if (serializeTimer != null)
        {
            serializeTimer.SafeDispose();
            //Application.TraceVerbose($"Serializing instance settings to '{FilePath}' on dispose");
            Serialize(JsonSourceGenerationContext.Default);
        }
    }
}
