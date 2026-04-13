namespace ShellN.Extensions.Utilities;

public class FileOpenDialog(nint site = 0, bool enableEvents = false)
    : FileDialog(DirectN.Extensions.Com.ComObject.CoCreate<IFileOpenDialog>(Constants.FileOpenDialog)!, site, enableEvents)
{
    public static FILEOPENDIALOGOPTIONS DefaultOptions { get; } = FILEOPENDIALOGOPTIONS.FOS_FILEMUSTEXIST | FILEOPENDIALOGOPTIONS.FOS_PATHMUSTEXIST;

    public virtual IReadOnlyList<ShellItem> GetResults()
    {
        ((IFileOpenDialog)NativeObject).GetResults(out var array);
        using var items = new ComObject<IShellItemArray>(array);
        return [.. items.Enumerate()];
    }

    public virtual IReadOnlyList<ShellItem> GetSelectedItems()
    {
        ((IFileOpenDialog)NativeObject).GetSelectedItems(out var array);
        using var items = new ComObject<IShellItemArray>(array);
        return [.. items.Enumerate()];
    }
}
