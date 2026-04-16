namespace ShellBat.Model;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class WebPropertiesWindow : WebWindow
{
    public const string WindowId = "Properties";

    internal static string GetNewId() => $"{WindowId}_{Guid.NewGuid():N}";

    public WebPropertiesWindow(ShellBatWindow window, Entry entry, string id = WindowId)
        : base(window, id, entry?.ParsingName)
    {
        ArgumentNullException.ThrowIfNull(entry);
        Entry = entry;
        Options.Icon = "fa-solid fa-file";
        Options.Title = entry.DisplayName;
        Viewers.AddRange(Viewer.GetSupportedViewers(entry).OrderBy(v => v.SortOrder));
    }

    public Entry Entry { get; }
    public string ParsingName => Entry.ParsingName;

    public override bool IsInstanceEqual(object? other) => ComObject.WithComInstance(other, unk =>
    {
        if (ComObject.ComWrappers.GetOrCreateObjectForComInstance(unk, CreateObjectFlags.Unwrap) is not WebPropertiesWindow obj)
            return false;

        return obj.ParsingName.EqualsIgnoreCase(ParsingName);
    }, throwIfZero: false);
}


