namespace ShellBat;

[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)]
public sealed partial class History : JsonBasedSettings<History>
{
    public const string GlobalFileName = "globalHistory.json";

    public static History Current { get; internal set; } = null!;
    private static string _filePath = null!;

    internal static void SetHistory(string filePath)
    {
        if (Current != null && _filePath.EqualsIgnoreCase(filePath))
            return;

        _filePath = filePath;
        History? history = null;
        if (IOUtilities.PathIsFile(filePath))
        {
            history = Deserialize(filePath, JsonSourceGenerationContext.Default, true);
        }

        Current = history ?? new History();
        Current.FilePath = filePath;
    }

    private readonly Navigation _navigation = new();

    [Browsable(false)]
    public IEnumerable<HistoryEntry> Entries { get => _navigation.Entries; set => _navigation.Entries = value; }

    [JsonIgnore]
    [Browsable(false)]
    public HistoryEntry? PreviousEntry => _navigation.Previous;

    [JsonIgnore]
    [Browsable(false)]
    public HistoryEntry? NextEntry => _navigation.Next;

    [JsonIgnore]
    [Browsable(false)]
    public HistoryEntry? CurrentEntry => _navigation.Current;

    [JsonIgnore]
    [Browsable(false)]
    public int EntriesCount => _navigation.Count;

    public void SerializeToConfiguration() => Serialize(JsonSourceGenerationContext.Default);
    public void Backup() => Backup(TimeSpan.FromDays(Settings.Current.BackupsMaxDays));

    public async Task AddEntry(Entry? entry)
    {
        if (entry is null)
            return;

        await _navigation.Visit(entry);
        SerializeToConfiguration();
    }

    public int Clear() => _navigation.Clear();
    public bool RemoveDeletedEntries() => _navigation.RemoveDeletedEntries();
    public bool RemoveEntries(string? rootParsingName)
    {
        if (rootParsingName == null)
            return false;

        var changed = false;
        foreach (var entry in _navigation.Entries.ToList())
        {
            if (entry.ParsingName.EqualsIgnoreCase(rootParsingName) ||
                entry.ParsingName.StartsWith(rootParsingName + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
                entry.ParsingName.StartsWith(rootParsingName + Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            {
                if (_navigation.Remove(entry))
                {
                    changed = true;
                }
            }
        }
        return changed;
    }

    public bool RemoveEntry(HistoryEntry? entry)
    {
        if (entry == null)
            return false;

        return _navigation.Remove(entry);
    }

    public HistoryEntry? MoveBack(int step = 1)
    {
        if (step < 1)
            return null;

        return _navigation.MoveBack(step);
    }

    public HistoryEntry? MoveForward(int step = 1)
    {
        if (step < 1)
            return null;

        return _navigation.MoveForward(step);
    }

    private sealed class Navigation
    {
        private readonly List<HistoryEntry> _entries = [];
        private int _currentIndex = -1;

        public int Count => _entries.Count;
        public HistoryEntry? Current => _currentIndex < 0 ? null : _entries[_currentIndex];
        public HistoryEntry? Previous => _currentIndex >= _entries.Count - 1 ? null : _entries[_currentIndex + 1];
        public HistoryEntry? Next => _currentIndex <= 0 ? null : _entries[_currentIndex - 1];

        public bool Remove(HistoryEntry entry) => _entries.Remove(entry);
        public async Task Visit(Entry entry)
        {
            var index = _entries.FindIndex(e => e.ParsingName.EqualsIgnoreCase(entry.ParsingName));
            if (index < 0)
            {
                if (entry.IsDesktop)
                    return;

                var he = new HistoryEntry
                {
                    ParsingName = entry.ParsingName,
                    DisplayName = entry.FullDisplayName,
                    IconPath = await ShellBatHttpLocalServer.GetImageCacheFilePath(UrlType.Icon, entry.ParsingName, 16)
                };
                _entries.Add(he);
            }
            else
            {
                // if already exists, move to that entry
                var existing = _entries[index];
                existing.LastVisitedTime = DateTime.Now;
                existing.DisplayName = entry.FullDisplayName;
#pragma warning disable IDE0074 // Use compound assignment
                if (existing.IconPath == null)
                {
                    existing.IconPath = await ShellBatHttpLocalServer.GetImageCacheFilePath(UrlType.Icon, entry.ParsingName, 16);
                }
#pragma warning restore IDE0074 // Use compound assignment
            }

            _entries.Sort();

            // maximum 100 entries
            var max = Settings.Current.MaximumHistoryEntries.Clamp(10, 1000);
            if (_entries.Count > max)
            {
                _entries.RemoveRange(max, _entries.Count - max);
            }
            _currentIndex = 0;
        }

        public HistoryEntry? MoveBack(int step)
        {
            if (_entries.Count == 0)
                return null;

            var newIndex = Math.Min(_entries.Count - 1, _currentIndex + step);
            if (newIndex == _currentIndex)
                return null;

            _currentIndex = newIndex;
            return _entries[_currentIndex];
        }

        public HistoryEntry? MoveForward(int step)
        {
            if (_entries.Count == 0)
                return null;

            var newIndex = Math.Max(0, _currentIndex - step);
            if (newIndex == _currentIndex)
                return null;

            _currentIndex = newIndex;
            return _entries[_currentIndex];
        }

        public IEnumerable<HistoryEntry> Entries
        {
            get => _entries;
            set
            {
                _entries.Clear();
                if (value != null)
                {
                    _entries.AddRange(value);
                }

                _entries.Sort();
                _currentIndex = _entries.Count > 0 ? 0 : -1;
            }
        }

        public int Clear()
        {
            var count = _entries.Count;
            _entries.Clear();
            _currentIndex = -1;
            return count;
        }

        public bool RemoveDeletedEntries()
        {
            var changed = false;
            foreach (var entry in _entries.ToArray())
            {
                if (Entry.Get(null, entry.ParsingName, ShellItemParsingOptions.DontThrowOnError) == null && _entries.Remove(entry))
                {
                    changed = true;
                }
            }
            return changed;
        }
    }
}
