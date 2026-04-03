namespace ShellN.Extensions.Utilities;

public abstract class FileDialog : InterlockedComObject<IFileDialog>
{
    public event EventHandler<FileDialogEventArgs>? Event;
    private FileDialogEvents? _events;
    private uint _eventsCookie;

    protected FileDialog(IComObject<IFileDialog> fileDialog, nint site = 0, bool enableEvents = false)
        : base(fileDialog)
    {
        if (site != 0 && NativeObject is IObjectWithSite ows)
        {
            ows.SetSite(site).ThrowOnError();
        }

        if (enableEvents)
        {
            _events = new FileDialogEvents();
            _events.Event += (s, e) => OnEvent(this, e);
            NativeObject.Advise(_events, out _eventsCookie).ThrowOnError();
        }
    }

    protected virtual void OnEvent(object sender, FileDialogEventArgs e) => Event?.Invoke(sender, e);

    // type format is "Description|*.ext1;*.ext2"
    public virtual void SetFileTypes(IEnumerable<string> types, bool throwOnError = true)
    {
        if (types == null)
            return;

        static string getName(string type)
        {
            var parts = type.Split('|');
            return parts.Length > 1 ? parts[0] : type;
        }

        static string getSpec(string type)
        {
            var parts = type.Split('|');
            return parts.Length > 1 ? parts[1] : type;
        }

        var fileTypes = types.Select(t => new COMDLG_FILTERSPEC
        {
            pszName = new(Marshal.StringToCoTaskMemUni(getName(t))),
            pszSpec = new(Marshal.StringToCoTaskMemUni(getSpec(t)))
        }).ToArray();
        if (fileTypes.Length == 0)
            return;

        try
        {
            NativeObject.SetFileTypes(fileTypes.Length(), fileTypes).ThrowOnError(throwOnError);
        }
        finally
        {
            foreach (var ft in fileTypes)
            {
                Marshal.FreeCoTaskMem(ft.pszName.Value);
                Marshal.FreeCoTaskMem(ft.pszSpec.Value);
            }
        }
    }

    public virtual void SetOptions(FILEOPENDIALOGOPTIONS options) => NativeObject.SetOptions(options);
    public virtual void SetFileTypeIndex(uint index) => NativeObject.SetFileTypeIndex(index);
    public virtual int GetFileTypeIndex() { if (NativeObject.GetFileTypeIndex(out var index).IsError) return -1; return (int)index; }
    public virtual void SetFileName(string name) => NativeObject.SetFileName(PWSTR.From(name));
    public virtual void SetFileNameLabel(string name) => NativeObject.SetFileNameLabel(PWSTR.From(name));
    public virtual void SetDefaultExtension(string extension) => NativeObject.SetDefaultExtension(PWSTR.From(extension));
    public virtual void SetTitle(string title) => NativeObject.SetTitle(PWSTR.From(title));
    public virtual void SetFilter(IShellItemFilter? filter) => NativeObject.SetFilter(filter!).ThrowOnError();
    public virtual bool Show(HWND owner) => NativeObject.Show(owner).IsOk;

    public virtual string? GetFileName()
    {
        NativeObject.GetFileName(out var pszName);
        return pszName.ToStringAndDispose();
    }

    public virtual ShellFolder? GetFolder()
    {
        NativeObject.GetFolder(out var itemObj);
        var item = ShellItem.FromObject(itemObj);
        if (item is ShellFolder folder)
            return folder;

        item?.Dispose();
        return null;
    }

    public virtual ShellItem? GetResult()
    {
        NativeObject.GetResult(out var item);
        return ShellItem.FromObject(item);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            var cookie = Interlocked.Exchange(ref _eventsCookie, 0);
            if (cookie != 0)
            {
                NativeObject.Unadvise(cookie);
            }

            if (RawNativeObject is IObjectWithSite ows)
            {
                ows.SetSite(0);
            }
        }
        base.Dispose(disposing);
    }
}
