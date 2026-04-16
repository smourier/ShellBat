namespace ShellN.Extensions;

public sealed class ShellFolder : ShellItem
{
    public static ShellFolder Desktop => _desktop.Value;
    private static readonly Lazy<ShellFolder> _desktop = new(() => { Functions.SHGetDesktopFolder(out var pf).ThrowOnError(); return (ShellFolder)FromObject(pf, owned: false)!; }, true);

    public static ShellFolder MyComputer => _myComputer.Value;
    private static readonly Lazy<ShellFolder> _myComputer = new(() => (ShellFolder)FromParsingName($"::{Constants.CLSID_MyComputer:B}", owned: false)!, true);

    internal ShellFolder(IShellItem2 comObject, bool owned)
        : this(new ComObject<IShellItem2>(comObject), owned)
    {
    }

    internal ShellFolder(IComObject<IShellItem2> comObject, bool owned)
        : base(comObject, owned)
    {
        if (!IsFolder)
            throw new ArgumentException(null, nameof(comObject));
    }

    public bool IsDesktop => ClassId == Constants.ShellDesktop;
    public bool IsRecycleBin => ClassId == Constants.CLSID_RecycleBin;
    public bool IsApplications => ClassId == Constants.CLSID_Applications;
    public bool IsMyComputer => ClassId == Constants.CLSID_MyComputer;
    public override FileAttributes? FileAttributes => System.IO.FileAttributes.Directory | base.FileAttributes;
    public Guid? ClassId
    {
        get
        {
            using var sf = BindToHandler<IShellFolder>(Constants.BHID_SFObject);
            var persist = sf.As<IPersist>();
            if (persist == null)
                return null;

            persist.Object.GetClassID(out var clsid);
            return clsid;
        }
    }

    public HRESULT RenameItem(ShellItem item, string newName, _TRANSFER_SOURCE_FLAGS flags = _TRANSFER_SOURCE_FLAGS.TSF_NORMAL, bool throwOnError = true)
        => RenameItem(item?.NativeObject!, newName, flags, throwOnError);

    public HRESULT RenameItem(IShellItem item, string newName, _TRANSFER_SOURCE_FLAGS flags = _TRANSFER_SOURCE_FLAGS.TSF_NORMAL, bool throwOnError = true)
    {
        var hr = RenameItem(item, newName, out var renamedItem, flags, throwOnError);
        renamedItem?.Dispose();
        return hr;
    }

    public HRESULT RenameItem(ShellItem item, string newName, out ShellItem? renamedItem, _TRANSFER_SOURCE_FLAGS flags = _TRANSFER_SOURCE_FLAGS.TSF_NORMAL, bool throwOnError = true)
        => RenameItem(item?.NativeObject!, newName, out renamedItem, flags, throwOnError);

    public HRESULT RenameItem(IShellItem item, string newName, out ShellItem? renamedItem, _TRANSFER_SOURCE_FLAGS flags = _TRANSFER_SOURCE_FLAGS.TSF_NORMAL, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(item);
        renamedItem = null;
        using var handler = BindToHandler<ITransferSource>(Constants.BHID_Transfer);
        if (handler == null)
            return DirectN.Constants.E_NOINTERFACE;

        var hr = handler.Object.RenameItem(item, PWSTR.From(newName), (uint)flags, out var newItemObj).ThrowOnError(throwOnError);
        if (hr.IsError)
            return hr;

        var newItem = new ComObject<IShellItem>(newItemObj);
        if (newItem == null)
            return DirectN.Constants.E_NOINTERFACE;

        renamedItem = FromObject(newItem.Object);
        return DirectN.Constants.S_OK;
    }

    // another way but less controllable
    //using var app = DirectN.Extensions.Com.ComObject.CoCreate<DirectN.IDispatch>("Shell.Application");
    //using var ns = app.CallMethodComObject<DirectN.IDispatch>("Namespace", [SIGDN_DESKTOPABSOLUTEPARSING]);
    //ns.CallObjectMethod("NewFolder", [name]);
    //
    // also note: _TRANSFER_SOURCE_FLAGS.TSF_RENAME_EXIST creates a "name (Copy)" which is not what we want
    public ShellItem? CreateNewFolder(string name, bool makeUniqueName = true)
        => CreateNewItem(name, System.IO.FileAttributes.Directory, 0, _TRANSFER_SOURCE_FLAGS.TSF_NORMAL, makeUniqueName);

    public ShellItem? CreateNewItem(
        string name,
        FileAttributes attributes = System.IO.FileAttributes.Normal,
        ulong size = 0,
        _TRANSFER_SOURCE_FLAGS flags = _TRANSFER_SOURCE_FLAGS.TSF_NORMAL,
        bool makeUniqueName = true)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException(null, nameof(name));

        using var handler = BindToHandler<ITransferDestination>(Constants.BHID_Transfer);
        if (handler == null)
            return null;

        return ShellExtensions.MakeUniqueWithExtension(name, (tentativeName) =>
        {
            var hr = handler.Object.CreateItem(
                PWSTR.From(tentativeName),
                (uint)attributes,
                size,
                (uint)flags,
                typeof(IShellItem).GUID,
                out var newItemPtr,
                typeof(IShellItemResources).GUID,
                out var ppvResources);
            if (hr.IsSuccess)
            {
                using var resources = DirectN.Extensions.Com.ComObject.FromPointer<IShellItemResources>(ppvResources); // we don't use it but need to release
                var newItem = DirectN.Extensions.Com.ComObject.FromPointer<IShellItem>(newItemPtr);
                return (FromObject(newItem?.Object), MakeUniqueResult.Success);
            }

            if (hr != DirectN.Constants.ERROR_ALREADY_EXISTS)
                return (null, MakeUniqueResult.Error);

            if (!makeUniqueName)
                return (null, MakeUniqueResult.Error);

            return (null, MakeUniqueResult.Continue);
        });
    }

    public IEnumerable<ShellItem> EnumerateChildrenWhere(Func<ShellItem, bool> includes, _SHCONTF? flags = null, bool owned = true)
    {
        ArgumentNullException.ThrowIfNull(includes);
        foreach (var child in EnumerateChildren(flags, owned))
        {
            if (includes(child))
                yield return child;

            child.Dispose();
        }
    }

    public IEnumerable<ShellItem> EnumerateAllChildrenWhere(Func<ShellItem, bool> includes, _SHCONTF? flags = null, bool owned = true)
    {
        ArgumentNullException.ThrowIfNull(includes);
        foreach (var child in EnumerateAllChildren(flags, owned))
        {
            if (includes(child))
                yield return child;

            child.Dispose();
        }
    }

    public IEnumerable<ShellItem> EnumerateAllChildren(_SHCONTF? flags = null, bool owned = true)
    {
        var folders = new List<ShellFolder>();

        _SHCONTF? fl = flags;
        if (fl != null)
        {
            fl |= _SHCONTF.SHCONTF_FOLDERS;
        }

        foreach (var child in EnumerateChildren(fl, owned))
        {
            if (child is ShellFolder folder)
            {
                folders.Add(folder);
            }
            else
            {
                if (flags == null || flags.Value.HasFlag(_SHCONTF.SHCONTF_NONFOLDERS))
                    yield return child;
            }
        }

        foreach (var folder in folders)
        {
            foreach (var child in folder.EnumerateAllChildren(flags, owned))
            {
                yield return child;
            }

            if (flags == null || flags.Value.HasFlag(_SHCONTF.SHCONTF_FOLDERS))
                yield return folder;
        }
    }

    public IEnumerable<ShellItem> EnumerateChildren(_SHCONTF? flags = null, bool owned = true)
    {
        IComObject<IBindCtx>? context = null;
        if (flags.HasValue)
        {
            context = IBindCtxExtensions.CreateBindCtx() ?? throw new InvalidOperationException();
            context.Object.AddToBindCtx(Constants.STR_ENUM_ITEMS_FLAGS, flags.Value).ThrowOnError();
        }

        using var items = BindToHandler<IEnumShellItems>(Constants.BHID_EnumItems, context);
        if (items == null)
        {
            context?.Dispose();
            yield break;
        }

        try
        {
            var index = 0;
            var shellItems = new nint[1];
            do
            {
                var hr = items.Object.Next(1, shellItems, 0);
                if (!hr.IsOk)
                    break;

                var child = FromObject(shellItems[0], owned: owned);
                if (child == null)
                    continue;

                if (child != null)
                {
                    child.ChildIndex = index++;
                    yield return child;
                }
            }
            while (true);
        }
        finally
        {
            context?.Dispose();
        }
    }
}
