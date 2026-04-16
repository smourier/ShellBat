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

    public virtual void SetDefaultFolder(IComObject<IShellItem> folder) => NativeObject.SetDefaultFolder(folder?.Object!).ThrowOnError();
    public virtual void SetDefaultFolder(ShellFolder folder) => NativeObject.SetDefaultFolder(folder?.NativeObject!).ThrowOnError();
    public virtual void SetDefaultFolder(IShellItem folder) => NativeObject.SetDefaultFolder(folder).ThrowOnError();

    public virtual void SetNavigationRoot(ShellItem root) => ((IFileDialog2)NativeObject).SetNavigationRoot(root?.NativeObject!).ThrowOnError();
    public virtual void SetNavigationRoot(IComObject<IShellItem> root) => ((IFileDialog2)NativeObject).SetNavigationRoot(root?.Object!).ThrowOnError();
    public virtual void SetNavigationRoot(IShellItem root) => ((IFileDialog2)NativeObject).SetNavigationRoot(root).ThrowOnError();

    public virtual void SetClientGuid(in Guid guid) => NativeObject.SetClientGuid(guid).ThrowOnError();
    public virtual void ClearClientData() => NativeObject.ClearClientData().ThrowOnError();

    public virtual void AddPlace(IComObject<IShellItem> item, FDAP fdap) => NativeObject.AddPlace(item?.Object!, fdap).ThrowOnError();
    public virtual void AddPlace(ShellItem item, FDAP fdap) => NativeObject.AddPlace(item?.NativeObject!, fdap).ThrowOnError();
    public virtual void AddPlace(IShellItem item, FDAP fdap) => NativeObject.AddPlace(item, fdap).ThrowOnError();

    public virtual void SetOptions(FILEOPENDIALOGOPTIONS options) => NativeObject.SetOptions(options).ThrowOnError();
    public virtual void SetFileTypeIndex(uint index) => NativeObject.SetFileTypeIndex(index).ThrowOnError();
    public virtual int GetFileTypeIndex() { if (NativeObject.GetFileTypeIndex(out var index).IsError) return -1; return (int)index; }
    public virtual void SetFileName(string name) => NativeObject.SetFileName(PWSTR.From(name)).ThrowOnError();
    public virtual void SetFileNameLabel(string name) => NativeObject.SetFileNameLabel(PWSTR.From(name)).ThrowOnError();
    public virtual void SetDefaultExtension(string extension) => NativeObject.SetDefaultExtension(PWSTR.From(extension)).ThrowOnError();
    public virtual void SetTitle(string title) => NativeObject.SetTitle(PWSTR.From(title)).ThrowOnError();
    public virtual void SetFilter(IShellItemFilter? filter) => NativeObject.SetFilter(filter!).ThrowOnError();
    public virtual bool Show(HWND owner) => NativeObject.Show(owner).IsOk;

    public virtual void SetCancelButtonLabel(string label) => ((IFileDialog2)NativeObject).SetCancelButtonLabel(PWSTR.From(label)).ThrowOnError();
    public virtual void SetOkButtonLabel(string label) => ((IFileDialog2)NativeObject).SetOkButtonLabel(PWSTR.From(label)).ThrowOnError();

    public virtual string? GetFileName()
    {
        NativeObject.GetFileName(out var pszName);
        return pszName.ToStringAndDispose();
    }

    public virtual ShellItem? GetCurrentSelection()
    {
        NativeObject.GetCurrentSelection(out var item);
        if (item == null)
            return null;

        return ShellItem.FromObject(item);
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

    public virtual void Close(HRESULT hr) => NativeObject.Close(hr).ThrowOnError();
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
