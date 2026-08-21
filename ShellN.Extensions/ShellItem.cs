namespace ShellN.Extensions;

public partial class ShellItem : InterlockedComObject<IShellItem2>, IItemWithAbsoluteIdList
{
    // note Attributes doesn't read *all* flags by default, only the capability and storage masks (most essential and not too slow to get)
    public static SFGAO_FLAGS AskedForFlags { get; set; } = SFGAO_FLAGS.SFGAO_CAPABILITYMASK | SFGAO_FLAGS.SFGAO_STORAGECAPMASK;

    private const string _contextMenuPropName = "__ShellNContextMenu";

    public static ShellItem? FromObject(object? obj, bool noFolder = false, bool owned = true)
    {
        if (obj == null)
            return null;

        if (obj is nint unk)
        {
            var si = DirectN.Extensions.Com.ComObject.FromPointer<IShellItem2>(unk);
            if (si != null)
                return FromObject(si.Object, noFolder, owned);

            return null;
        }

        if (obj is ShellItem item)
            return item;

        if (obj is IComObject<IDataObject> dataObject)
            return FromDataObject(dataObject, DATAOBJ_GET_ITEM_FLAGS.DOGIF_DEFAULT, noFolder, owned);

        if (obj is IShellFolder folder)
        {
            var CLSID_ShellItem = new Guid("9ac9fbe1-e0a2-4ad6-b4ee-e212013ea917");
            var withFolder = ComObject<IInitializeWithFolder>.CoCreate(CLSID_ShellItem, throwOnError: false);
            if (withFolder != null)
            {
                if (withFolder.Object.Initialize(folder).IsSuccess)
                    return FromObject(withFolder.Object, noFolder, owned);
            }
            return null;
        }

        if (obj is IShellItem2 shellItem)
        {
            if (noFolder)
                return new ShellItem(shellItem, owned);

            shellItem.GetAttributes(SFGAO_FLAGS.SFGAO_FOLDER, out var atts);
            if (!atts.HasFlag(SFGAO_FLAGS.SFGAO_FOLDER))
                return new ShellItem(shellItem, owned);

            return new ShellFolder(shellItem, owned);
        }

        return null;
    }

    public static ShellItem? FromItemIdList(ItemIdList? idList, bool noFolder = false, bool owned = true, bool throwOnError = true)
    {
        if (idList is null || idList.Pointer == 0 || idList.Size <= 2)
            return null;

        return FromPidl(idList.Pointer, noFolder, owned, throwOnError);
    }

    public static ShellItem? FromPidl(nint pidl, bool noFolder = false, bool owned = true, bool throwOnError = true)
    {
        if (pidl == 0)
            return null;

        Functions.SHCreateItemFromIDList(pidl, typeof(IShellItem).GUID, out var obj).ThrowOnError(throwOnError);
        return FromObject(obj, noFolder, owned);
    }

    public static ShellItem? FromParsingName(string name, IBindCtx? context = null, bool noFolder = false, bool owned = true, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(name);
        Functions.SHCreateItemFromParsingName(PWSTR.From(name), context, typeof(IShellItem2).GUID, out var ppv).ThrowOnError(throwOnError);
        return FromObject(ppv, noFolder, owned);
    }

    public static ShellItem? FromSplitParsingName(string name, IBindCtx? context = null, bool noFolder = false, bool owned = true, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(name);

        var segments = SplitPath(name).ToArray();
        if (segments.Length == 0)
            return null;

        if (segments.Length == 1)
            return FromParsingName(name, context, noFolder, owned, throwOnError);

        ShellFolder? folder = null;
        // get folders first
        for (var i = 0; i < segments.Length - 1; i++)
        {
            var segment = segments[i];
            ShellItem? item;
            if (i == 0)
            {
                item = FromParsingName(segment, context, false, owned, throwOnError);
            }
            else
            {
                item = folder!.FindChild(s => s.SIGDN_PARENTRELATIVEPARSING.EqualsIgnoreCase(segment));
            }

            if (item is not ShellFolder f)
            {
                item?.Dispose();
                folder?.Dispose();
                return null;
            }

            folder = f;
        }
        if (folder == null)
            return null;

        var lastSegment = segments[^1];
        return folder.FindChild(i => i.SIGDN_PARENTRELATIVEPARSING.EqualsIgnoreCase(lastSegment));
    }

    public static ShellItem? FromDataObject(IComObject<IDataObject>? dataObject, DATAOBJ_GET_ITEM_FLAGS flags = DATAOBJ_GET_ITEM_FLAGS.DOGIF_DEFAULT, bool noFolder = false, bool owned = true, bool throwOnError = true)
        => FromDataObject(dataObject?.Object, flags, noFolder, owned, throwOnError);

    public static ShellItem? FromDataObject(IDataObject? dataObject, DATAOBJ_GET_ITEM_FLAGS flags = DATAOBJ_GET_ITEM_FLAGS.DOGIF_DEFAULT, bool noFolder = false, bool owned = true, bool throwOnError = true)
    {
        if (dataObject == null)
            return null;

        Functions.SHGetItemFromDataObject(dataObject, flags, typeof(IShellItem).GUID, out var ppv).ThrowOnError(throwOnError);
        return FromObject(ppv, noFolder, owned);
    }

    public static IReadOnlyList<ShellItem> ArrayFromDataObject(IComObject<IDataObject>? dataObject, bool noFolder = false, bool owned = true, bool throwOnError = true)
        => ArrayFromDataObject(dataObject?.Object, noFolder, owned, throwOnError);

    public static IReadOnlyList<ShellItem> ArrayFromDataObject(IDataObject? dataObject, bool noFolder = false, bool owned = true, bool throwOnError = true)
    {
        if (dataObject == null)
            return [];

        Functions.SHCreateShellItemArrayFromDataObject(dataObject, typeof(IShellItemArray).GUID, out var ppv).ThrowOnError(throwOnError);
        using var array = DirectN.Extensions.Com.ComObject.FromPointer<IShellItemArray>(ppv);
        if (array == null)
            return [];

        return [.. array.Enumerate(noFolder, owned, throwOnError)];
    }

    public static void OpenDesktop() => OpenFoldersAndSelectItems([string.Empty]);
    public static void OpenFoldersAndSelectItems(IEnumerable<string> parsingNames, OFASI flags = OFASI.OFASI_NONE, bool throwOnError = false)
    {
        if (parsingNames == null)
            return;

        var items = new List<ShellItem>();
        foreach (var name in parsingNames)
        {
            var item = FromParsingName(name, throwOnError: throwOnError);
            if (item != null)
            {
                items.Add(item);
            }
        }
        OpenFoldersAndSelectItems(items, flags, disposeItems: true);
    }

    public static void OpenFoldersAndSelectItems(IEnumerable<ShellItem> items, OFASI flags = OFASI.OFASI_NONE, bool disposeItems = true, bool throwOnError = false)
    {
        if (items == null)
            return;

        var disposables = new List<IDisposable>();
        if (disposeItems)
        {
            disposables.AddRange(items);
        }

        var itemsByPidls = new Dictionary<ItemIdList, List<ItemIdList>>(ItemIdListEqualityComparer.Instance);

        foreach (var item in items)
        {
            var parentIdl = item.GetParentAndItemIdList(out var itemIdl, throwOnError);
            if (parentIdl is null || itemIdl is null)
                continue;

            if (!itemsByPidls.TryGetValue(parentIdl, out var list))
            {
                list = [];
                itemsByPidls.Add(parentIdl, list);
            }

            disposables.Add(parentIdl);
            disposables.Add(itemIdl);
            list.Add(itemIdl);
        }

        try
        {
            unsafe
            {
                foreach (var kv in itemsByPidls)
                {
                    var pidls = kv.Value.Select(b => b.Pointer).ToArray();
                    Functions.SHOpenFolderAndSelectItems(kv.Key.Pointer, pidls.Length(), (nint)Unsafe.AsPointer(ref MemoryMarshal.GetArrayDataReference(pidls)), (uint)flags).ThrowOnError(throwOnError);
                }
            }
        }
        finally
        {
            disposables.Dispose();
        }
    }

    public static IComObject<IDataObject>? CreateEmptyDataObject(bool throwOnError = true) => CreateDataObject(null, disposeItems: false, throwOnError: throwOnError);
    public static unsafe IComObject<IDataObject>? CreateDataObject(IEnumerable<ShellItem>? items, bool disposeItems = true, bool throwOnError = true)
    {
        var disposables = new List<IDisposable>();
        if (disposeItems && items != null)
        {
            disposables.AddRange(items);
        }

        var idls = new List<ItemIdList>();
        if (items != null)
        {
            foreach (var item in items)
            {
                var idl = item.GetIdList(false);
                if (idl is null)
                    continue;

                disposables.Add(idl);
                idls.Add(idl);
            }
        }

        try
        {
            var pidls = idls.Select(i => i.Pointer).ToArray();
            var terminator = 0;
            var hr = Functions.SHCreateDataObject((nint)(&terminator), pidls.Length(), (nint)Unsafe.AsPointer(ref MemoryMarshal.GetArrayDataReference(pidls)), null, typeof(IDataObject).GUID, out var unk).ThrowOnError(throwOnError);
            if (unk == 0)
                return null;

            return DirectN.Extensions.Com.ComObject.FromPointer<IDataObject>(unk);
        }
        finally
        {
            disposables.Dispose();
        }
    }

    private Lazy<PropertyStore?> _fastPropertiesPropertyStore;

    internal ShellItem(IShellItem2 comObject, bool owned)
        : this(new ComObject<IShellItem2>(comObject), owned)
    {
    }

    internal ShellItem(IComObject<IShellItem2> comObject, bool owned)
        : base(comObject)
    {
        Owned = owned;
        ResetCachedPropertyStore(GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY);

        SIGDN_FILESYSPATH = GetDisplayName(SIGDN.SIGDN_FILESYSPATH, throwOnError: false);
        if (SIGDN_FILESYSPATH != null)
        {
            DriveNumber = GetDriveNumber(SIGDN_FILESYSPATH);
            if (DriveNumber >= 0)
            {
                const uint DRIVE_REMOTE = 4;
                var type = Functions.GetDriveTypeW(PWSTR.From(SIGDN_FILESYSPATH));
                IsNetwork = type == DRIVE_REMOTE;
                if (IsNetwork)
                {
                    var root = DriveRoot;
                    if (root != null && root.Length > 0)
                    {
                        // some "connected" drives are in fact disconnected drives (?)
                        // IsNetDrive & WNetGetConnection don't help here, they can say the same,
                        // depending on some things we can't control
                        IsSlowNetworkDrive = NetworkConnection.GetConnectedOrRememberedDisks()
                            .Any(d => d.LocalName?.Length > 0 && d.LocalName[0].Equals(root[0])); // all uppercase drive letters

                        if (IsSlowNetworkDrive)
                        {
                            // avoid kicking 20secs timeout on slow network drives
                            // presume it's like a normal, say, C: drive
                            Attributes = SFGAO_FLAGS.SFGAO_CANCOPY |
                                SFGAO_FLAGS.SFGAO_CANLINK |
                                SFGAO_FLAGS.SFGAO_STORAGE |
                                SFGAO_FLAGS.SFGAO_CANRENAME |
                                SFGAO_FLAGS.SFGAO_HASPROPSHEET |
                                SFGAO_FLAGS.SFGAO_DROPTARGET |
                                SFGAO_FLAGS.SFGAO_STORAGEANCESTOR |
                                SFGAO_FLAGS.SFGAO_FILESYSANCESTOR |
                                SFGAO_FLAGS.SFGAO_FOLDER |
                                SFGAO_FLAGS.SFGAO_FILESYSTEM;
                        }
                    }
                }
            }
        }
        else
        {
            DriveNumber = -1;
        }

        if (!IsSlowNetworkDrive)
        {
            NativeObject.GetAttributes(AskedForFlags, out var flags);
            Attributes = flags;
        }
    }

    public bool Owned { get; }
    public int ChildIndex { get; internal set; }
    public SFGAO_FLAGS Attributes { get; }
    public bool IsNetwork { get; }
    public bool IsSlowNetworkDrive { get; }
    public int DriveNumber { get; }
    public bool IsDrive => DriveRoot != null && DriveRoot.EqualsIgnoreCase(SIGDN_FILESYSPATH);
    public string? DriveRoot => DriveNumber >= 0 ? $"{(char)(DriveNumber + 'A')}:\\" : null;
    public bool IsFolder => Attributes.HasFlag(SFGAO_FLAGS.SFGAO_FOLDER);
    public bool IsHidden => Attributes.HasFlag(SFGAO_FLAGS.SFGAO_HIDDEN);
    public bool IsReadOnly => Attributes.HasFlag(SFGAO_FLAGS.SFGAO_READONLY);
    public string? Type { get { if (!TryGetPropertyValue(PROPERTYKEYS.SystemItemType, out string? value)) return null; return value; } }
    public string? Extension { get { if (!TryGetPropertyValue(PROPERTYKEYS.SystemFileExtension, out string? value)) return null; return value; } }
    public string? SIGDN_NORMALDISPLAY => GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY, throwOnError: false);
    public string? SIGDN_FILESYSPATH { get; }
    public string? SIGDN_DESKTOPABSOLUTEPARSING => GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEPARSING, throwOnError: false);
    public string? SIGDN_DESKTOPABSOLUTEEDITING => GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEEDITING, throwOnError: false);
    public string? SIGDN_URL => GetDisplayName(SIGDN.SIGDN_URL, throwOnError: false);
    public string? SIGDN_PARENTRELATIVEPARSING => GetDisplayName(SIGDN.SIGDN_PARENTRELATIVEPARSING, throwOnError: false);
    public DateTime? DateModified { get { if (!TryGetPropertyValue(PROPERTYKEYS.SystemDateModified, out DateTime value)) return null; return value; } }
    public DateTime? DateCreated { get { if (!TryGetPropertyValue(PROPERTYKEYS.SystemDateCreated, out DateTime value)) return null; return value; } }
    public DateTime? DateAccessed { get { if (!TryGetPropertyValue(PROPERTYKEYS.SystemDateAccessed, out DateTime value)) return null; return value; } }
    public PERCEIVED? Perceived { get { if (!TryGetPropertyValue(PROPERTYKEYS.SystemPerceivedType, GETPROPERTYSTOREFLAGS.GPS_DEFAULT, out PERCEIVED value)) return null; return value; } }
    public long? Size { get { if (!TryGetPropertyValue(PROPERTYKEYS.SystemSize, out long value)) return null; return value; } }
    public virtual FileAttributes? FileAttributes { get { if (!TryGetPropertyValue(PROPERTYKEYS.SystemFileAttributes, out FileAttributes value)) return null; return value; } }
    public IEnumerable<KeyValuePair<PROPERTYKEY, object?>> Properties => EnumerateProperties(throwOnError: false);

    public int EncryptionFlags
    {
        get
        {
            using var ps = GetPropertyStore(GETPROPERTYSTOREFLAGS.GPS_DEFAULT, false);
            if (ps == null)
                return 0;

            ps.TryGetValue<int>(PROPERTYKEYS.SystemVolumeBitLockerProtection, out var value);
            return value;
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (!Owned)
            return;

        base.Dispose(disposing);
    }

    private static int GetDriveNumber(string? path)
    {
        if (path != null && path.Length >= 2 && char.IsAsciiLetter(path[0]) && path[1] == Path.VolumeSeparatorChar)
            return char.ToUpperInvariant(path[0]) - 'A';

        return -1;
    }

    public IComObject<IDataObject>? CreateDataObject(bool throwOnError = true) => CreateDataObject([this], false, throwOnError);
    public ShellFolder? GetParent(bool owned = true)
    {
        NativeObject.GetParent(out var item);
        return FromObject(item, owned: owned) as ShellFolder;
    }

    public unsafe ItemIdList? GetParentIdList(bool throwOnError = false)
    {
        nint parentPidl = 0;
        var hr = ((IParentAndItem)NativeObject).GetParentAndItem((nint)(&parentPidl), 0, 0).ThrowOnError(throwOnError);
        if (hr.IsError)
            return null;

        return new ItemIdList(parentPidl, true);
    }

    public unsafe ItemIdList? GetParentAndItemIdList(out ItemIdList? itemRelativeIdList, bool throwOnError = false)
    {
        itemRelativeIdList = null;
        nint parentPidl = 0;
        nint childPidl = 0;
        var hr = ((IParentAndItem)NativeObject).GetParentAndItem((nint)(&parentPidl), 0, (nint)(&childPidl)).ThrowOnError(throwOnError);
        if (hr.IsError)
            return null;

        if (childPidl != 0)
        {
            itemRelativeIdList = new ItemIdList(childPidl, true);
        }
        return parentPidl != 0 ? new ItemIdList(parentPidl, true) : null;
    }

    public unsafe ItemIdList? GetItemRelativeIdList(bool throwOnError = false)
    {
        nint childPidl = 0;
        var hr = ((IParentAndItem)NativeObject).GetParentAndItem(0, 0, (nint)(&childPidl)).ThrowOnError(throwOnError);
        if (hr.IsError)
            return null;

        return childPidl != 0 ? new ItemIdList(childPidl, true) : null;
    }

    [MemberNotNull(nameof(_fastPropertiesPropertyStore))]
    public virtual void ResetCachedPropertyStore(GETPROPERTYSTOREFLAGS flags = GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY)
    {
        if (flags != GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY)
            throw new NotSupportedException();

        _fastPropertiesPropertyStore = new Lazy<PropertyStore?>(GetFastPropertyStore);
    }

    private PropertyStore? GetFastPropertyStore()
    {
        NativeObject.GetPropertyStore(GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY, typeof(IPropertyStore).GUID, out var ppv);
        if (ppv == 0)
            return null;

        return new PropertyStore(DirectN.Extensions.Com.ComObject.FromPointer<IPropertyStore>(ppv)!);
    }

    public PropertyStore? GetPropertyStore(GETPROPERTYSTOREFLAGS flags = GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY, bool throwOnError = true)
    {
        if (flags == GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY)
            return _fastPropertiesPropertyStore.Value;

        NativeObject.GetPropertyStore(flags, typeof(IPropertyStore).GUID, out var ppv).ThrowOnError(throwOnError);
        if (ppv == 0)
            return null;

        return new PropertyStore(DirectN.Extensions.Com.ComObject.FromPointer<IPropertyStore>(ppv)!);
    }

    ItemIdList? IItemWithAbsoluteIdList.AbsoluteIdList => GetIdList();
    ItemId? IItemWithId.Id => GetIdList()?.LastId;

    public override string ToString() => SIGDN_NORMALDISPLAY ?? string.Empty;

    public HRESULT TryOpenStream(out Stream? stream, IBindCtx? ctx = null)
    {
        stream = null;
        var hr = NativeObject.BindToHandler(ctx, Constants.BHID_Stream, typeof(DirectN.IStream).GUID, out var unk);
        if (hr.IsError)
            return hr;

        var strm = DirectN.Extensions.Com.ComObject.FromPointer<DirectN.IStream>(unk);
        if (strm == null)
            return DirectN.Constants.E_FAIL;

        stream = new StreamOnIStream(strm.Object, true);
        return DirectN.Constants.S_OK;
    }

    public Stream? OpenStream(IComObject<IBindCtx>? ctx = null) => OpenStream(ctx?.Object);
    public Stream? OpenStream(IBindCtx? ctx = null)
    {
        NativeObject.BindToHandler(ctx, Constants.BHID_Stream, typeof(DirectN.IStream).GUID, out var unk);
        var stream = DirectN.Extensions.Com.ComObject.FromPointer<DirectN.IStream>(unk);
        if (stream == null)
            return null;

        return new StreamOnIStream(stream.Object, true);
    }

    public Stream? OpenReadStream() => OpenStream((IBindCtx?)null);
    public Stream? OpenWriteStream()
    {
        using var ctx = IBindCtxExtensions.CreateBindCtx(STGM.STGM_WRITE)!;
        return OpenStream(ctx);
    }

    public nint GetStorageItem(bool throwOnError = true)
    {
        nint item = 0;
        DirectN.Extensions.Com.ComObject.WithComInstanceOfType<IShellItem>(NativeObject, unk =>
        {
            var options = new STORAGEITEM_FROM_SHELLITEM_CREATE_OPTIONS();
            Functions.CreateStorageItemFromShellItem(unk, options, typeof(IUnknown).GUID, out var ptr).ThrowOnError(throwOnError);
            item = ptr;
        });
        return item;
    }

    public string? GetDisplayName(SIGDN type, bool throwOnError = true)
    {
        NativeObject.GetDisplayName(type, out var name).ThrowOnError(throwOnError);
        return name.ToString();
    }

    // note this has limits, only ShowContextMenu can show all items
    public MenuItem GetContextMenu(object? site = null, IComObject<IBindCtx>? context = null, CMF flags = CMF.CMF_NORMAL) => GetContextMenu<MenuItem>(site, context, flags) ?? new MenuItem();
    public virtual T? GetContextMenu<T>(object? site = null, IComObject<IBindCtx>? context = null, CMF flags = CMF.CMF_NORMAL) where T : MenuItem, new()
    {
        using var cm = BindToHandler<IContextMenu>(Constants.BHID_SFUIObject);
        if (cm == null)
            return null;

        var cm2 = cm.As<IContextMenu2>();
        if (cm2 == null)
            return null;

        if (site != null && cm.Object is IObjectWithSite ows)
        {
            DirectN.Extensions.Com.ComObject.WithComInstance(site, ows.SetSite, true, false); // don't check any error
        }

        var menu = Functions.CreateMenu();
        if (menu.Value == 0)
            return null;

        try
        {
            HRESULT hr;
            var cm3 = cm2.As<IContextMenu3>();
            if (cm3 != null)
            {
                unsafe
                {
                    LRESULT result = 0;
                    hr = cm3.Object.HandleMenuMsg2(MessageDecoder.WM_INITMENUPOPUP, menu.Value, 0, (nint)(&result));
                    if (hr.IsError)
                    {
                        var hr2 = cm2.Object.HandleMenuMsg(MessageDecoder.WM_INITMENUPOPUP, menu.Value, 0);
                    }
                }
            }
            else
            {
                hr = cm2.Object.HandleMenuMsg(MessageDecoder.WM_INITMENUPOPUP, menu.Value, 0);
            }

            if (cm.Object.QueryContextMenu(menu, 0, 0, Constants.FCIDM_SHVIEWLAST, (uint)flags).IsError)
                return null;

            var item = new T();
            item.Extract(menu, cm2);
            return item;
        }
        finally
        {
            Functions.DestroyMenu(menu);
        }
    }

    public HRESULT InvokeDefaultCommand(bool throwOnError = true) => InvokeDefaultCommand(HWND.Null, throwOnError);
    public virtual HRESULT InvokeDefaultCommand(HWND hwnd, bool throwOnError = true)
    {
        if (ComObject.Object is not IParentAndItem pai)
            return DirectN.Constants.E_NOINTERFACE.ThrowOnError(throwOnError);

        nint psf;
        nint child;
        unsafe
        {
            var hr = pai.GetParentAndItem(0, (nint)(&psf), (nint)(&child)).ThrowOnError(throwOnError);
            if (hr.IsError)
                return hr;
        }

        using var shellFolder = DirectN.Extensions.Com.ComObject.FromPointer<IShellFolder>(psf);
        if (shellFolder == null || child == 0)
            return DirectN.Constants.E_NOINTERFACE.ThrowOnError(throwOnError);

        return Functions.SHInvokeDefaultCommand(hwnd, shellFolder.Object, child);
    }

    public HRESULT InvokeCommand(string? verb, bool throwOnError = true) => InvokeCommand(verb, HWND.Null, throwOnError);
    public virtual HRESULT InvokeCommand(string? verb, HWND hwnd, bool throwOnError = true)
    {
        if (ComObject.Object is not IParentAndItem pai)
            return DirectN.Constants.E_NOINTERFACE.ThrowOnError(throwOnError);

        nint psf;
        nint child;
        unsafe
        {
            var hr = pai.GetParentAndItem(0, (nint)(&psf), (nint)(&child)).ThrowOnError(throwOnError);
            if (hr.IsError)
                return hr;
        }

        using var shellFolder = DirectN.Extensions.Com.ComObject.FromPointer<IShellFolder>(psf);
        if (shellFolder == null || child == 0)
            return DirectN.Constants.E_NOINTERFACE.ThrowOnError(throwOnError);

        return Functions.SHInvokeCommand(hwnd, shellFolder.Object, child, PWSTR.From(verb));
    }

    public virtual HRESULT ShowContextMenu(
        object? site = null,
        IComObject<IBindCtx>? context = null,
        CMF flags = CMF.CMF_NORMAL,
        Func<IContextMenu, HWND, uint, HRESULT>? invoke = null)
    {
        using var pidl = GetIdList(false);
        if (pidl is null)
            return DirectN.Constants.E_FAIL;

        using var parent = GetParentIdList();
        if (parent is null)
            return DirectN.Constants.E_FAIL;

        return ShowContextMenu(parent, [pidl], site, context, flags, invoke);
    }

    public static HRESULT ShowContextMenu(
        ItemIdList folderIdList,
        IEnumerable<ItemIdList> list,
        object? site = null,
        IComObject<IBindCtx>? context = null,
        CMF flags = CMF.CMF_NORMAL,
        Func<IContextMenu, HWND, uint, HRESULT>? invoke = null)
    {
        ArgumentNullException.ThrowIfNull(folderIdList);
        ArgumentNullException.ThrowIfNull(list);
        var array = list.ToArray();

        DirectN.Functions.GetCursorPos(out var cursor);

        HWND hwnd = 0;
        if (site is DirectN.IOleWindow window)
        {
            window.GetWindow(out hwnd);
        }
        else if (site is IOleWindow window2)
        {
            window2.GetWindow(out hwnd);
        }

        using var folder = FromPidl(folderIdList.Pointer);
        if (folder == null)
            return DirectN.Constants.E_FAIL;

        using var shellFolder = folder.BindToHandler<IShellFolder2>(Constants.BHID_SFObject, context);
        if (shellFolder == null)
            return DirectN.Constants.E_NOINTERFACE;

        // note this code is a bit special because Explorer is not consistent
        // sometimes it needs the relative pidls to be valid (ie: are children of folder)
        // sometimes it not only need the relative pidls to be valid but it wants them (ex: folders under My PC)
        // so we do it in two passes
        // 1) try with all relative pidls. If it succeeds, ok (case of folders under My PC and others)
        // 2) if it fails, try with only the relative pidls that are children of folder
        var relativePidls = new List<nint>();
        for (var i = 0; i < array.Length; i++)
        {
            relativePidls.Add(array[i].RelativePointer);
        }

        nint unk;
        if (relativePidls.Count == 0)
        {
            shellFolder.Object.CreateViewObject(hwnd, typeof(IContextMenu).GUID, out unk);
        }
        else
        {
            if (shellFolder.Object.GetUIObjectOf(hwnd, relativePidls.Length(), [.. relativePidls], typeof(IContextMenu).GUID, 0, out unk).IsError)
            {
                relativePidls.Clear();
                for (var i = 0; i < array.Length; i++)
                {
                    // no need to free VARIANT for Size
                    if (shellFolder.Object.GetDetailsEx(array[i].RelativePointer, PropertyKeys.System.Size, out var name).IsError)
                        continue;

                    relativePidls.Add(array[i].RelativePointer);
                }

                if (relativePidls.Count == 0)
                {
                    shellFolder.Object.CreateViewObject(hwnd, typeof(IContextMenu).GUID, out unk);
                }
                else
                {
                    shellFolder.Object.GetUIObjectOf(hwnd, relativePidls.Length(), [.. relativePidls], typeof(IContextMenu).GUID, 0, out unk);
                }
            }
        }

        using var cm = DirectN.Extensions.Com.ComObject.FromPointer<IContextMenu>(unk);
        if (cm == null)
            return DirectN.Constants.E_NOINTERFACE;

        var cm2 = cm.As<IContextMenu2>();
        if (cm2 == null)
            return DirectN.Constants.E_NOINTERFACE;

        DirectN.Extensions.Com.ComObject.WithComInstance(cm2, unk =>
        {
            DirectN.Functions.SetPropW(hwnd, PWSTR.From(_contextMenuPropName), new HANDLE(unk));
        });

        if (site != null && cm.Object is IObjectWithSite ows)
        {
            DirectN.Extensions.Com.ComObject.WithComInstance(site, ows.SetSite, true, false); // don't check any error
        }

        var menu = Functions.CreatePopupMenu();
        if (menu.Value == 0)
            return DirectN.Constants.E_FAIL;

        const uint firstCommandId = 1;
        try
        {
            var hr = cm.Object.QueryContextMenu(menu, 0, firstCommandId, Constants.FCIDM_SHVIEWLAST, (uint)flags);
            if (hr.IsError)
                return hr;

            var id = Functions.TrackPopupMenu(menu, TRACK_POPUP_MENU_FLAGS.TPM_RETURNCMD, cursor.x, cursor.y, 0, hwnd, 0);
            if (id != 0)
            {
                if (invoke != null)
                    return invoke(cm.Object, hwnd, (uint)(id - (int)firstCommandId));

                return Invoke(cm.Object, hwnd, (uint)(id - (int)firstCommandId));
            }
            return DirectN.Constants.S_OK;
        }
        finally
        {
            DirectN.Extensions.Com.ComObject.WithComInstance(cm2, unk =>
            {
                DirectN.Functions.RemovePropW(hwnd, PWSTR.From(_contextMenuPropName));
            });
            Functions.DestroyMenu(menu);
        }
    }

    public static HRESULT OnContextMenuWindowMessage(HWND hwnd, uint msg, WPARAM wParam, LPARAM lParam, out LRESULT result)
    {
        result = 0;
        var unk = DirectN.Functions.GetPropW(hwnd, PWSTR.From(_contextMenuPropName));
        if (unk != 0)
        {
            // don't release the object we don't own it
            using var cm2 = DirectN.Extensions.Com.ComObject.FromPointer<IContextMenu2>(unk, releaseOnDispose: false);
            var cm3 = cm2.As<IContextMenu3>();
            if (cm3 != null)
            {
                unsafe
                {
                    LRESULT res = 0;
                    var hr = cm3.Object.HandleMenuMsg2(msg, wParam, lParam, (nint)(&res));
                    result = res;
                    if (hr.IsSuccess)
                        return hr;
                }
            }

            if (cm2 != null)
            {
                var hr = cm2.Object.HandleMenuMsg(msg, wParam, lParam);
                return hr;
            }
        }
        return DirectN.Constants.E_NOINTERFACE;
    }

    public static HRESULT InvokeContextMenuVerb(IEnumerable<ItemIdList> list, string verb, object? site = null, IComObject<IBindCtx>? context = null, CMF flags = CMF.CMF_NORMAL)
    {
        ArgumentNullException.ThrowIfNull(list);
        ArgumentNullException.ThrowIfNull(verb);
        var array = list.ToArray();
        if (array.Length == 0)
            return DirectN.Constants.E_INVALIDARG;

        DirectN.Functions.GetCursorPos(out var cursor);

        HWND hwnd = 0;
        if (site is DirectN.IOleWindow window)
        {
            window.GetWindow(out hwnd);
        }
        else if (site is IOleWindow window2)
        {
            window2.GetWindow(out hwnd);
        }

        // note this code presumes all items are in the same folder, which is usually the case
        // the first item is used to get the folder
        // if the items are in different folders, the context menu result is undefined
        using var first = FromPidl(array[0].Pointer, throwOnError: false);
        if (first == null)
            return DirectN.Constants.E_FAIL;

        using var firstFolder = first.GetParent();
        if (firstFolder == null)
            return DirectN.Constants.E_FAIL;

        var relativePidls = new nint[array.Length];
        for (var i = 0; i < array.Length; i++)
        {
            relativePidls[i] = array[i].RelativePointer;
        }

        using var shellFolder = firstFolder.BindToHandler<IShellFolder>(Constants.BHID_SFObject, context);
        if (shellFolder == null)
            return DirectN.Constants.E_FAIL;

        shellFolder.Object.GetUIObjectOf(hwnd, relativePidls.Length(), relativePidls, typeof(IContextMenu).GUID, 0, out var unk);
        using var cm = DirectN.Extensions.Com.ComObject.FromPointer<IContextMenu>(unk);
        if (cm == null)
            return DirectN.Constants.E_NOINTERFACE;

        var cm2 = cm.As<IContextMenu2>();
        if (cm2 == null)
            return DirectN.Constants.E_NOINTERFACE;

        DirectN.Extensions.Com.ComObject.WithComInstance(cm2, unk =>
        {
            DirectN.Functions.SetPropW(hwnd, PWSTR.From(_contextMenuPropName), new HANDLE(unk));
        });

        if (site != null && cm.Object is IObjectWithSite ows)
        {
            DirectN.Extensions.Com.ComObject.WithComInstance(site, ows.SetSite, true, false); // don't check any error
        }

        var menu = Functions.CreatePopupMenu();
        if (menu.Value == 0)
            return DirectN.Constants.E_OUTOFMEMORY;

        const uint firstCommandId = 1;
        try
        {
            if (cm.Object.QueryContextMenu(menu, 0, firstCommandId, Constants.FCIDM_SHVIEWLAST, (uint)flags).IsError)
                return DirectN.Constants.E_FAIL;

            return Invoke(cm.Object, hwnd, verb, 0);
        }
        finally
        {
            DirectN.Extensions.Com.ComObject.WithComInstance(cm2, unk =>
            {
                DirectN.Functions.RemovePropW(hwnd, PWSTR.From(_contextMenuPropName));
            });
            Functions.DestroyMenu(menu);
        }
    }

    public static HRESULT Invoke(IContextMenu cm, HWND hwnd, uint id) => Invoke(cm, hwnd, null, id);
    public static HRESULT Invoke(IContextMenu cm, HWND hwnd, string verb)
    {
        ArgumentNullException.ThrowIfNull(verb);
        return Invoke(cm, hwnd, verb, 0);
    }

    private static unsafe HRESULT Invoke(IContextMenu cm, HWND hwnd, string? verb, uint id)
    {
        ArgumentNullException.ThrowIfNull(cm);

        var mask = CMIC.CMIC_MASK_UNICODE | CMIC.CMIC_MASK_PTINVOKE;
        var info = new CMINVOKECOMMANDINFOEX
        {
            cbSize = (uint)sizeof(CMINVOKECOMMANDINFOEX),
            hwnd = hwnd,
            nShow = (int)SHOW_WINDOW_CMD.SW_SHOWNORMAL,
        };

        DirectN.Functions.GetCursorPos(out info.ptInvoke);

        if (DirectN.Functions.GetKeyState((int)VIRTUAL_KEY.VK_CONTROL) < 0)
        {
            mask |= CMIC.CMIC_MASK_CONTROL_DOWN;
        }

        if (DirectN.Functions.GetKeyState((int)VIRTUAL_KEY.VK_SHIFT) < 0)
        {
            mask |= CMIC.CMIC_MASK_SHIFT_DOWN;
        }

        info.fMask = (uint)mask;

        if (verb != null)
        {
            info.lpVerb = PSTR.From(verb);
            info.lpVerbW = PWSTR.From(verb);
        }
        else
        {
            info.lpVerb = new PSTR((int)id);
            info.lpVerbW = new PWSTR((int)id);
        }

        return cm.InvokeCommand(Unsafe.AsRef<CMINVOKECOMMANDINFO>((CMINVOKECOMMANDINFO*)&info));
    }

    public byte[]? GetIdListAsByteArray(bool throwOnError = true)
    {
        using var idl = GetIdList(throwOnError);
        return idl?.ToByteArray();
    }

    public ItemIdList? GetIdList(bool throwOnError = true) => DirectN.Extensions.Com.ComObject.WithComInstance(Disposable, unk =>
    {
        Functions.SHGetIDListFromObject(unk, out var pidl).ThrowOnError(throwOnError);
        if (pidl == 0)
            return null;

        return new ItemIdList(pidl, true);
    });

    public T? GetPropertyValue<T>(PROPERTYKEY key, T? defaultValue = default, GETPROPERTYSTOREFLAGS flags = GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY)
    {
        if (!TryGetPropertyValue(key, flags, out T? obj))
            return defaultValue;

        return obj;
    }

    public bool TryGetPropertyValue<T>(PROPERTYKEY key, out T? value) => TryGetPropertyValue(key, GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY, out value);
    public bool TryGetPropertyValue<T>(PROPERTYKEY key, GETPROPERTYSTOREFLAGS flags, out T? value)
    {
        if (!TryGetPropertyValue(key, flags, out object? obj))
        {
            value = default;
            return false;
        }
        return Conversions.TryChangeType(obj, out value);
    }

    public bool TryGetPropertyValue(PROPERTYKEY key, out object? value) => TryGetPropertyValue(key, GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY, out value);
    public bool TryGetPropertyValue(PROPERTYKEY key, GETPROPERTYSTOREFLAGS flags, out object? value)
    {
        if (flags == GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY)
        {
            var ps2 = GetPropertyStore(flags, throwOnError: false);
            if (ps2 == null)
            {
                value = null;
                return false;
            }

            return ps2.TryGetValue(key, out value);
        }

        using var ps = GetPropertyStore(flags, throwOnError: false);
        if (ps == null)
        {
            value = null;
            return false;
        }

        return ps.TryGetValue(key, out value);
    }

    public IEnumerable<KeyValuePair<PROPERTYKEY, object?>> EnumerateProperties(GETPROPERTYSTOREFLAGS flags = GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY, bool throwOnError = true)
    {
        var ps = GetPropertyStore(flags, throwOnError);
        if (ps == null)
            yield break;

        foreach (var kv in ps)
        {
            yield return kv;
        }
    }

    public HRESULT Rename(string newName, _TRANSFER_SOURCE_FLAGS flags = _TRANSFER_SOURCE_FLAGS.TSF_NORMAL, bool throwOnError = true)
    {
        using var folder = GetParent();
        if (folder == null)
            return DirectN.Constants.E_FAIL;

        return folder.RenameItem(this, newName, flags, throwOnError);
    }

    public HRESULT Rename(string newName, out ShellItem? renamedItem, _TRANSFER_SOURCE_FLAGS flags = _TRANSFER_SOURCE_FLAGS.TSF_NORMAL, bool throwOnError = true)
    {
        using var folder = GetParent();
        if (folder == null)
        {
            renamedItem = null;
            return DirectN.Constants.E_FAIL;
        }

        return folder.RenameItem(this, newName, out renamedItem, flags, throwOnError);
    }

    // note this can be used with T = IPortableDevice, bhid = BHID_SFObject, when used on WPD namespace extensions (Apple iPhone, etc.)
    public IComObject<T>? BindToHandler<T>(Guid bhid, IComObject<IBindCtx>? context = null)
    {
        NativeObject.BindToHandler(context?.Object!, bhid, typeof(T).GUID, out var unk);
        return DirectN.Extensions.Com.ComObject.FromPointer<T>(unk);
    }

    public unsafe HICON GetIconHandle(SHIL shil)
    {
        if (ComObject.Object is not IParentAndItem pai)
            return 0;

        nint psf;
        nint child;
        if (pai.GetParentAndItem(0, (nint)(&psf), (nint)(&child)).IsError)
            return 0;

        using var shellFolder = DirectN.Extensions.Com.ComObject.FromPointer<IShellFolder>(psf);
        if (shellFolder == null || child == 0)
            return 0;

        int indexSel;
        var index = Functions.SHMapPIDLToSystemImageListIndex(shellFolder.Object, child, (nint)(&indexSel));
        Marshal.FreeCoTaskMem(child);

        var list = ImageList.GetImageList(shil, false);
        if (list == null)
            return 0;

        list.Object.GetIcon(index, 0, out var hicon);
        return hicon;
    }

    public IComObject<IWICBitmap>? GetIconAsBitmap(SHIL shil)
    {
        var hicon = GetIconHandle(shil);
        if (hicon == 0)
            return null;

        var bitmap = WicImagingFactory.CreateBitmapFromHICON(hicon);
        DirectN.Functions.DestroyIcon(hicon);
        return bitmap;
    }

    public HBITMAP GetImageHandle(SIZE size, SIIGBF flags = SIIGBF.SIIGBF_RESIZETOFIT)
    {
        var factory = ComObject.As<IShellItemImageFactory>();
        if (factory == null)
            return 0;

        // note this can return E_PENDING if the image is not ready yet (and hbmp == 0)
        // if this happens a lot (big file, takes too much time, etc), you should call GetImageHandleAsync instead
        factory.Object.GetImage(size, flags, out var hbmp);
        return hbmp;
    }

    public IComObject<IWICBitmap>? GetImageAsBitmap(SIZE size, SIIGBF flags = SIIGBF.SIIGBF_RESIZETOFIT, WICBitmapAlphaChannelOption options = WICBitmapAlphaChannelOption.WICBitmapUseAlpha)
    {
        var hbmp = GetImageHandle(size, flags);
        if (hbmp == 0)
            return null;

        var bitmap = WicImagingFactory.CreateBitmapFromHBITMAP(hbmp, options);
        DirectN.Functions.DeleteObject(new HGDIOBJ { Value = hbmp });
        return bitmap;
    }

    public async Task<HBITMAP> GetImageHandleAsync(SIZE size, SIIGBF flags = SIIGBF.SIIGBF_RESIZETOFIT)
    {
        var factory = ComObject.As<IShellItemImageFactory>();
        if (factory == null)
            return 0;

        var hr = factory.Object.GetImage(size, flags, out var hbmp);
        // see here https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-iextractimage-getlocation#ieiflag_async
        // E_PENDING means the image is not ready yet, and we must call it on another thread
        if (hbmp == 0 && hr == DirectN.Constants.E_PENDING)
        {
            hbmp = await Task.Run(() => GetImageHandle(size, flags));
        }
        return hbmp;
    }

    public async Task<IComObject<IWICBitmap>?> GetImageAsBitmapAsync(SIZE size, SIIGBF flags = SIIGBF.SIIGBF_RESIZETOFIT, WICBitmapAlphaChannelOption options = WICBitmapAlphaChannelOption.WICBitmapUseAlpha)
    {
        var hbmp = await GetImageHandleAsync(size, flags);
        if (hbmp == 0)
            return null;

        var bitmap = WicImagingFactory.CreateBitmapFromHBITMAP(hbmp, options);
        DirectN.Functions.DeleteObject(new HGDIOBJ { Value = hbmp });
        return bitmap;
    }

    public static IEnumerable<string> SplitPath(string? path)
    {
        if (string.IsNullOrEmpty(path))
            yield break;

        var current = new StringBuilder();
        var i = 0;
        while (i < path.Length)
        {
            var c = path[i];
            if (c != '\\')
            {
                current.Append(c);
                i++;
                continue;
            }

            // we hit a '\'. If the next char is also '\', this '\' is the separator between the previous
            // segment and an opaque "\\x\y" segment that starts right after it.
            var startsOpaque = i + 1 < path.Length && path[i + 1] == '\\';

            // flush current segment (the '\' is a separator, not part of any segment).
            yield return current.ToString();
            current.Clear();
            i++; // consume the separator '\'

            if (startsOpaque)
            {
                // parse the opaque "\\server\share" block INTO current, but do not flush it here.
                // The next '\' (if any) will be handled by the outer loop as a regular separator, and the final flush at the end of the
                // method handles the case where the string ends with the opaque block.

                // consume the "\\" opaque marker.
                while (i < path.Length && path[i] == '\\')
                {
                    current.Append(path[i]);
                    i++;
                }

                // consume "server"/"host" (non-backslash chars).
                while (i < path.Length && path[i] != '\\')
                {
                    current.Append(path[i]);
                    i++;
                }

                // consume the single internal '\'.
                if (i < path.Length && path[i] == '\\')
                {
                    current.Append(path[i]);
                    i++;
                }

                // consume "share"/"device" (non-backslash chars).
                while (i < path.Length && path[i] != '\\')
                {
                    current.Append(path[i]);
                    i++;
                }
            }
        }

        // Flush the trailing segment (possibly empty if path ended with '\').
        yield return current.ToString();
    }
}
