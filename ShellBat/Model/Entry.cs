namespace ShellBat.Model;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class Entry : DispatchObject, IEquatable<Entry>
{
    public static Entry Desktop { get; } = new Entry(null, string.Empty, ShellFolder.Desktop);
    public static Entry MyComputer { get; } = new Entry(null, ShellFolder.MyComputer.SIGDN_DESKTOPABSOLUTEPARSING!, ShellFolder.MyComputer);
    public static Entry UserProfile { get; } = Get(null, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))!;

    private static readonly ConcurrentDictionary<string, object?> _clipboardCuts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Entry? _existingParent;
    private readonly Lazy<Entry?> _parent;
    private readonly string? _formattedLastWriteTime;

    private Entry(
        Entry? parent,
        string parsingName,
        string displayName,
        string fullDisplayName,
        FileExtension extension)
    {
        _existingParent = parent;
        _parent = new Lazy<Entry?>(GetParent);
        ParsingName = parsingName;
        DisplayName = displayName;
        FullDisplayName = fullDisplayName;
        Extension = extension;
    }

    private Entry(Entry? parent, string parsingName, ShellItem item)
    {
        if (parsingName.EndsWith(Path.DirectorySeparatorChar))
        {
            parsingName = parsingName[..^1];
        }

        _existingParent = parent;
        if (item is ShellFolder folder)
        {
            IsDesktop = folder.IsDesktop == true;
            IsRecycleBin = folder.IsRecycleBin == true;
            IsMyComputer = folder.IsMyComputer == true;
            IsApplications = folder.IsApplications == true;
            ParsingName = IsDesktop ? string.Empty : parsingName;
        }
        else
        {
            ParsingName = parsingName;
        }

        Attributes = item.Attributes;
        FileSystemPath = item.SIGDN_FILESYSPATH;
        DisplayName = item.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY, true)!;
        FullDisplayName = item.GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEEDITING, true)!;
        LastWriteTime = item.DateModified;
        CreationTime = item.DateCreated;
        LastAccessTime = item.DateAccessed;
        Size = item.Size ?? 0;
        FileAttributes = item.FileAttributes ?? (IsFolder ? FileAttributes.Directory : FileAttributes.Normal);
        var extension = item.Extension.Nullify() ?? Path.GetExtension(ParsingName).Nullify();
        Extension = FileExtension.GetByExtension(extension) ?? FileExtension.None;
        _parent = new Lazy<Entry?>(GetParent);

        if (!item.IsSlowNetworkDrive && item.IsDrive)
        {
            Drive = DriveInfo.GetDrives().FirstOrDefault(d => d.Name.EqualsIgnoreCase(FileSystemPath));
            if (Drive != null)
            {
                try
                {
                    var ratio = (double)Drive.AvailableFreeSpace / Drive.TotalSize;
                    _formattedLastWriteTime = $"[{Drive.DriveFormat}] {Res.Free} {ratio:P0} ({ShellBatExtensions.FormatByteSize(Drive.AvailableFreeSpace)} / {ShellBatExtensions.FormatByteSize(Drive.TotalSize)})";
                }
                catch
                {
                    // continue
                }
            }

            VirtualDiskImagePath = IOUtilities.GetVirtualDiskImagePath(ParsingName, false);
        }

        if (_formattedLastWriteTime == null)
        {
            var df = ShellBatInstance.Current.Settings.DateFormat ?? "yyyy/MM/dd HH:mm:ss";
            if (string.IsNullOrWhiteSpace(df) || df.EqualsIgnoreCase("default"))
            {
                _formattedLastWriteTime = LastWriteTime?.ToString();
            }
            else
            {
                _formattedLastWriteTime = LastWriteTime?.ToString(df);
            }
        }
    }

    public bool IsDesktop { get; }
    public bool IsRecycleBin { get; }
    public bool IsMyComputer { get; }
    public bool IsApplications { get; }
    public DriveInfo? Drive { get; }
    public string? VirtualDiskImagePath { get; }
    public bool IsZipEntry => this is ZipEntry;
    public bool IsFolder => FileAttributes.HasFlag(FileAttributes.Directory) || Attributes.HasFlag(SFGAO_FLAGS.SFGAO_FOLDER);
    public bool IsHidden => FileAttributes.HasFlag(FileAttributes.Hidden) || Attributes.HasFlag(SFGAO_FLAGS.SFGAO_HIDDEN);
    public bool IsSystem => FileAttributes.HasFlag(FileAttributes.System) || Attributes.HasFlag(SFGAO_FLAGS.SFGAO_SYSTEM);
    public bool CanRename => Attributes.HasFlag(SFGAO_FLAGS.SFGAO_CANRENAME);
    public bool CanDelete => Attributes.HasFlag(SFGAO_FLAGS.SFGAO_CANDELETE);
    public bool CanCopy => Attributes.HasFlag(SFGAO_FLAGS.SFGAO_CANCOPY);
    public bool CanMove => Attributes.HasFlag(SFGAO_FLAGS.SFGAO_CANMOVE);
    public bool CanLink => Attributes.HasFlag(SFGAO_FLAGS.SFGAO_CANLINK);
    public bool IsShortcut => Extension.IsShortcut;
    public bool CanBeOpenedAsFolder => IsFolder || Extension.IsCompressed || Extension.CanBeDecompressed;
    public virtual bool IsNewFolderSupported => Attributes.HasFlag(SFGAO_FLAGS.SFGAO_STORAGEANCESTOR);
    public virtual string ParsingName { get; }
    public virtual string DisplayName { get; } = null!;
    public virtual string FullDisplayName { get; } = null!;
    public virtual FileExtension Extension { get; } = null!;
    public virtual string? FileSystemPath { get; }
    public virtual long Size { get; protected set; }
    public virtual DateTime? LastWriteTime { get; protected set; }
    public virtual DateTime? CreationTime { get; }
    public virtual DateTime? LastAccessTime { get; }
    public virtual FileAttributes FileAttributes { get; }
    public virtual SFGAO_FLAGS Attributes { get; protected set; }
    public Entry? Parent => _parent.Value;

    public string EditName
    {
        get
        {
            var name = FullDisplayName;
            if (name.Contains('\\'))
                return name;

            if (FileSystemPath != null)
                return FileSystemPath;

            return ParsingName;
        }
    }

    // maybe too hardcoded ?
    public string? WslFileSystemPath
    {
        get
        {
            var drive = IOUtilities.GetDrive(FileSystemPath);
            if (drive == '\0')
                return null;

            return "/mnt/" + char.ToLowerInvariant(drive) + FileSystemPath![2..].Replace(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
    }

    public string FormattedSize => Size > 0 || !IsFolder ? Size.ToString() : Res._Dir_;
    public string? FormattedLastWriteTime => _formattedLastWriteTime;
    public string? FormattedCreationTime => CreationTime?.ToString();
    public string? FormattedLastAccessTime => LastAccessTime?.ToString();

    public virtual bool IsConsoleApplication()
    {
        if (Extension.Name == ".bat" || Extension.Name == ".cmd")
            return true;

        if (Extension.Name == ".exe")
        {
            using var stream = OpenRead();
            if (stream == null)
                return false;

            var pi = PEInfo.FromStream(stream);
            if (pi?.ImageSubSystem == IMAGE_SUBSYSTEM.IMAGE_SUBSYSTEM_WINDOWS_CUI)
                return true;
        }
        return false;
    }

    public virtual HRESULT Rename(string newName, _TRANSFER_SOURCE_FLAGS flags = _TRANSFER_SOURCE_FLAGS.TSF_NORMAL, bool throwOnError = true) =>
        Rename(newName, out _, flags, throwOnError);

    public virtual HRESULT Rename(string newName, out Entry? renamedEntry, _TRANSFER_SOURCE_FLAGS flags = _TRANSFER_SOURCE_FLAGS.TSF_NORMAL, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(newName);
        renamedEntry = null;
        using var item = GetItem(throwOnError ? ShellItemParsingOptions.None : ShellItemParsingOptions.DontThrowOnError);
        if (item == null)
            return DirectN.Constants.E_FAIL;

        // get parent before we rename it
        var parent = Parent;

        var hr = item.Rename(newName, out var renamedItem, flags, throwOnError);
        if (hr.IsError)
            return hr;

        var parsingName = renamedItem?.SIGDN_DESKTOPABSOLUTEPARSING;
        if (parsingName == null)
            return DirectN.Constants.E_FAIL;

        renamedEntry = new Entry(parent, parsingName, renamedItem!);
        return DirectN.Constants.S_OK;
    }

    public virtual Entry? CreateNewFolder(string? name = null, bool makeUniqueName = true)
    {
        using var folder = GetAsFolder(ShellItemParsingOptions.DontThrowOnError);
        if (folder == null)
            return null;

        name = name.Nullify() ?? Res.NewFolderName;
        using var newFolder = folder.CreateNewFolder(name, makeUniqueName);
        if (newFolder == null)
            return null;

        var parsingName = newFolder.SIGDN_DESKTOPABSOLUTEPARSING;
        if (parsingName == null)
            return null;

        return new Entry(this, parsingName, newFolder);
    }

    public virtual Entry? CreateNewItem(
        string name,
        FileAttributes attributes = FileAttributes.Normal,
        ulong size = 0,
        _TRANSFER_SOURCE_FLAGS flags = _TRANSFER_SOURCE_FLAGS.TSF_NORMAL,
        bool makeUniqueName = true)
    {
        ArgumentNullException.ThrowIfNull(name);
        using var folder = GetAsFolder(ShellItemParsingOptions.DontThrowOnError);
        if (folder == null)
            return null;

        using var newItem = folder.CreateNewItem(name, attributes, size, flags, makeUniqueName);
        if (newItem == null)
            return null;

        var parsingName = newItem.SIGDN_DESKTOPABSOLUTEPARSING;
        if (parsingName == null)
            return null;

        return new Entry(this, parsingName, newItem);
    }

    public virtual Entry? GetShortcutTargetEntry()
    {
        var path = FileSystemPath;
        if (path == null)
            return null;

        using var link = Link.Load(path);
        if (link == null)
            return null;

        var targetPath = link.TargetPath;
        if (targetPath != null)
            return Get(null, targetPath, ShellItemParsingOptions.DontThrowOnError);

        if (link.TargetIdList is not null)
            return Get(null, link.TargetIdList);

        return null;
    }

    public IComObject<IDataObject>? CreateDataObject(bool throwOnError = true) => CreateDataObject([this], throwOnError);

    public ItemIdList? GetIdList(bool throwOnError = true)
    {
        using var item = GetItem(ParsingName, throwOnError ? ShellItemParsingOptions.None : ShellItemParsingOptions.DontThrowOnError);
        if (item == null)
            return null;

        return item.GetIdList(throwOnError);
    }

    public WebEntryOptions GetWebOptions(EntryEnumerateOptions enumerateOptions)
    {
        var options = WebEntryOptions.None;
        if (IsFolder && (!Extension.IsCompressed || !enumerateOptions.HasFlag(EntryEnumerateOptions.ShowCompressedAsFile)))
        {
            options |= WebEntryOptions.IsFolder;
        }

        if (IsHidden)
        {
            options |= WebEntryOptions.IsHidden;
        }

        if (IsSystem)
        {
            options |= WebEntryOptions.IsSystem;
        }

        if (Extension.IsCompressed)
        {
            options |= WebEntryOptions.IsCompressed;
        }

        if (Extension.IsCompressed && !enumerateOptions.HasFlag(EntryEnumerateOptions.ShowCompressedAsFile))
        {
            options |= WebEntryOptions.IsFolder;
        }

        if (!Extension.IsWebView2NativeImage)
        {
            options |= WebEntryOptions.IsNotWebView2NativeImage;
        }

        if (Extension.IsImage)
        {
            options |= WebEntryOptions.IsImage;
        }

        if (Extension.IsPdf)
        {
            options |= WebEntryOptions.IsPdf;
        }

        if (IsEntryInClipboardCut(this))
        {
            options |= WebEntryOptions.IsCut;
        }
        return options;
    }

    protected virtual Entry? GetParent()
    {
        if (_existingParent is not null)
            return _existingParent;

        using var item = GetItem(ShellItemParsingOptions.DontThrowOnError);
        if (item is null)
            return null;

        if (item is ShellFolder folder && folder.IsDesktop)
            return null;

        using var itemParent = item.GetParent();
        if (itemParent == null)
            return null;

        var parsingName = itemParent.SIGDN_DESKTOPABSOLUTEPARSING;
        if (parsingName == null)
        {
            Application.TraceError($"Could not get parsing name for parent of '{ParsingName}'");
            return null;
        }
        return new Entry(null, parsingName, itemParent);
    }

    private IEnumerable<Entry> EnumerateDesktopChildren(ShellFolder folder, EntryEnumerateOptions options)
    {
        var list = new List<Entry>();

        // add shell:appsFolder entry
        using var apps = ShellItem.FromParsingName(ShellN.Constants.ShellAppsFolder, throwOnError: false);
        if (apps != null)
        {
            var parsingName = apps.SIGDN_DESKTOPABSOLUTEPARSING;
            if (parsingName != null)
            {
                list.Add(new Entry(this, parsingName, apps));
            }
        }

        var flags = GetFlags(options);
        foreach (var child in folder.EnumerateChildren(flags))
        {
            try
            {
                var parsingName = child.SIGDN_DESKTOPABSOLUTEPARSING;
                if (parsingName == null)
                    continue;

                list.Add(new Entry(this, parsingName, child));
            }
            finally
            {
                child.Dispose();
            }
        }

        // for some mysterious reason, desktop has duplicate entries like "Downloads" (!?)
        return list.Distinct();
    }

    public virtual HRESULT CopyTo(IEnumerable<ShellItem> items, POINTL? pt = null, bool throwOnError = false)
    {
        using var targetItem = GetAsFolder(throwOnError ? ShellItemParsingOptions.None : ShellItemParsingOptions.DontThrowOnError);
        if (targetItem is null)
            return DirectN.Constants.E_FAIL;

        // these folders need special handling (this is so dumb why don't they support ITransfer things so IFileOperation works?)
        if (targetItem.ClassId == ShellN.Constants.CLSID_ZipFolder)
        {
            using var dropTarget = targetItem.BindToHandler<IDropTarget>(ShellN.Constants.BHID_Transfer);
            if (dropTarget is not null)
            {
                var dataObject = ShellItem.CreateDataObject(items, disposeItems: false, throwOnError: throwOnError);
                if (dataObject is not null)
                {
                    var effect = DROPEFFECT.DROPEFFECT_COPY;
                    dataObject.Object.SetPreferredDropEffect(effect);
                    DirectN.Functions.GetCursorPos(out var cursor);
                    var pt2 = pt ?? new POINTL { x = cursor.x, y = cursor.y };
                    var hrDrop = ShellN.Functions.SHSimulateDrop(dropTarget.Object, dataObject.Object, MODIFIERKEYS_FLAGS.MK_CONTROL | MODIFIERKEYS_FLAGS.MK_LBUTTON, ref pt2, ref effect);
                    if (hrDrop.IsSuccess)
                        return hrDrop;
                }
            }
            // else continue (and probably fail)
        }

        using var fo = new FileOperation();
        var hr = fo.SetOperationFlags(FILEOPERATION_FLAGS.FOFX_NOSKIPJUNCTIONS);
        if (hr.IsError)
            return hr;

        hr = fo.CopyItems(items.Select(i => i.NativeObject), targetItem.NativeObject, throwOnError);
        if (hr.IsError)
            return hr;

        return fo.PerformOperations(throwOnError);
    }

    public virtual HRESULT MoveTo(IEnumerable<ShellItem> items, POINTL? pt = null, bool throwOnError = false)
    {
        using var targetItem = GetAsFolder(throwOnError ? ShellItemParsingOptions.None : ShellItemParsingOptions.DontThrowOnError);
        if (targetItem is null)
            return DirectN.Constants.E_FAIL;

        // these folders need special handling (this is stupid...)
        var isZipFolder = targetItem.ClassId == ShellN.Constants.CLSID_ZipFolder;

        var movableItems = new List<ShellItem>();
        var copyableItems = new List<ShellItem>();
        if (isZipFolder)
        {
            // if zip folder, all items must be copied and deleted afterwards
            copyableItems.AddRange(items);
        }
        else
        {
            foreach (var item in items)
            {
                if (!item.Attributes.HasFlag(SFGAO_FLAGS.SFGAO_CANMOVE))
                {
                    movableItems.Add(item);
                    continue;
                }

                if (item.Attributes.HasFlag(SFGAO_FLAGS.SFGAO_CANCOPY) && item.Attributes.HasFlag(SFGAO_FLAGS.SFGAO_CANDELETE))
                {
                    copyableItems.Add(item);
                    continue;
                }
            }
        }

        if (copyableItems.Count > 0)
        {
            var hrCopy = CopyTo(copyableItems, pt, throwOnError);
            if (hrCopy.IsError)
                return hrCopy;

            using var fo = new FileOperation();
            var hr = fo.SetOperationFlags(FILEOPERATION_FLAGS.FOFX_RECYCLEONDELETE, throwOnError);
            if (hr.IsError)
                return hr;

            hr = fo.DeleteItems(items.Select(i => i.NativeObject), throwOnError);
            if (hr.IsError)
                return hr;

            hr = fo.PerformOperations(throwOnError);
            if (hr.IsError)
                return hr;
        }

        if (movableItems.Count > 0)
        {
            using var fo = new FileOperation();
            var hr = fo.SetOperationFlags(FILEOPERATION_FLAGS.FOFX_NOSKIPJUNCTIONS, throwOnError);
            if (hr.IsError)
                return hr;

            hr = fo.MoveItems(items.Select(i => i.NativeObject), targetItem.NativeObject, throwOnError);
            if (hr.IsError)
                return hr;

            hr = fo.PerformOperations(throwOnError);
            if (hr.IsError)
                return hr;
        }
        return DirectN.Constants.S_OK;
    }

    // don't go in recycle bin
    public virtual HRESULT Delete(bool throwOnError = false)
    {
        var parent = Parent;
        if (parent is null)
            return DirectN.Constants.E_FAIL;

        using var item = GetItem(throwOnError ? ShellItemParsingOptions.None : ShellItemParsingOptions.DontThrowOnError);
        if (item is null)
            return DirectN.Constants.E_FAIL;

        using var folder = parent.GetAsFolder(throwOnError ? ShellItemParsingOptions.None : ShellItemParsingOptions.DontThrowOnError);
        if (folder is null)
            return DirectN.Constants.E_FAIL;

        using var handler = folder.BindToHandler<ITransferSource>(ShellN.Constants.BHID_Transfer);
        if (handler is null)
            return DirectN.Constants.E_FAIL;

        return handler.Object.RemoveItem(item.NativeObject, 0);
    }

    public virtual IEnumerable<Entry> EnumerateChildren(EntryEnumerateOptions options = EntryEnumerateOptions.None)
    {
        if (!IsFolder)
        {
            if (Extension.CanBeDecompressed && !options.HasFlag(EntryEnumerateOptions.ShowCompressedAsFile))
            {
                foreach (var child in ZipEntry.EnumerateCompressedChildren(this, options))
                {
                    yield return child;
                }
            }
            yield break;
        }

        using var folder = GetAsFolder(ShellItemParsingOptions.DontThrowOnError);
        if (folder == null)
            yield break;

        if (folder.IsDesktop)
        {
            foreach (var child in EnumerateDesktopChildren(folder, options))
            {
                yield return child;
            }
            yield break;
        }

        var invokeDefaultCommand = !options.HasFlag(EntryEnumerateOptions.DontInvokeDefaultCommand);
        if (invokeDefaultCommand && IsInvokeDefaultCommandRequired(folder))
        {
            folder.InvokeDefaultCommand();
        }

        var flags = GetFlags(options);
        foreach (var child in folder.EnumerateChildren(flags))
        {
            try
            {
                var parsingName = child.SIGDN_DESKTOPABSOLUTEPARSING;
                if (parsingName == null)
                {
                    Application.TraceError($"Could not get parsing name for child of '{ParsingName}'");
                    continue;
                }

                Entry entry;
                // requires special handling because the parsing name is relative to that folder (!)
                if (folder.IsApplications)
                {
                    entry = new Entry(this, Path.Combine(ShellN.Constants.ShellAppsFolder, parsingName), child);
                }
                else
                {
                    entry = new Entry(this, parsingName, child);
                }

                yield return entry;
            }
            finally
            {
                child.Dispose();
            }
        }
    }

    public virtual Stream? OpenRead()
    {
        try
        {
            if (FileSystemPath != null && IOUtilities.PathIsFile(FileSystemPath))
                return new FileStream(FileSystemPath, FileMode.Open, FileAccess.Read, FileShare.Read);

            using var item = GetItem(ShellItemParsingOptions.DontThrowOnError);
            if (item == null)
                return null;

            item.ComObject.Object.BindToHandler(null, ShellN.Constants.BHID_Stream, typeof(DirectN.IStream).GUID, out var unk);
            var stream = ComObject.FromPointer<DirectN.IStream>(unk);
            if (stream == null)
                return null;

            return new StreamOnIStream(stream.Object, true);
        }
        catch
        {
            return null;
        }
    }

    public virtual IReadOnlyList<Entry> GetBreadcrumbChildren() => Parent?.EnumerateChildren(EntryEnumerateOptions.ExcludeFiles).OrderBy(c => c.DisplayName).ToArray() ?? [];
    public virtual IReadOnlyList<Entry> GetBreadcrumbs()
    {
        var crumbs = new List<Entry>();
        var current = this;
        while (current is not null)
        {
            crumbs.Add(current);
            current = current.Parent;
        }
        crumbs.Reverse();
        return crumbs;
    }

    public override int GetHashCode() => ParsingName.GetHashCode();
    public override bool Equals(object? obj) => Equals(obj as Entry);
    public bool Equals(Entry? other) => other is not null && ParsingName.EqualsIgnoreCase(other.ParsingName);
    public static bool operator ==(Entry? left, Entry? right) => left?.Equals(right) ?? right is null;
    public static bool operator !=(Entry? left, Entry? right) => !(left == right);

    public override string ToString() => ParsingName;

    public static bool IsEntryInClipboardCut(Entry? entry) => entry is not null && _clipboardCuts.ContainsKey(entry.ParsingName);
    public static int ClearClipboardCuts()
    {
        var count = _clipboardCuts.Count;
        _clipboardCuts.Clear();
        return count;
    }

    public static void AddToClipboardCut(Entry entry)
    {
        ArgumentNullException.ThrowIfNull(entry);
        _clipboardCuts[entry.ParsingName] = null;
    }

    public static Entry? Get(Entry? parent, ItemIdList? idl)
    {
        if (idl is null || idl.IsTerminator)
            return null;

        using var item = ShellItem.FromItemIdList(idl, throwOnError: false);
        if (item == null)
            return null;

        return Get(parent, item);
    }

    public static Entry? Get(Entry? parent, ShellItem? item)
    {
        if (item == null)
            return null;

        var parsingName = item.SIGDN_DESKTOPABSOLUTEPARSING;
        if (parsingName == null)
        {
            Application.TraceError($"Could not get parsing name");
            return null;
        }

        return new Entry(parent, parsingName, item);
    }

    public static Entry? Get(Entry? parent, IShellItem? item)
    {
        using var item2 = ShellItem.FromObject(item);
        if (item2 == null)
            return null;

        return Get(parent, item2);
    }

    public static Entry? Get(Entry? parent, string? parsingName, ShellItemParsingOptions options = ShellItemParsingOptions.None, IComObject<IBindCtx>? context = null)
    {
        if (string.IsNullOrWhiteSpace(parsingName))
            return Desktop;

        using var item = GetItem(parsingName, options, context);
        if (item == null)
            return GetZipEntry(parsingName, options, context);

        return item != null ? new Entry(parent, parsingName, item) : null;
    }

    private static Entry? GetZipEntry(string parsingName, ShellItemParsingOptions options, IComObject<IBindCtx>? context)
    {
        var segments = parsingName.Split(['\\']);
        var index = segments.IndexOf(s => FileExtension.CanBeDecompressedExtensions.Contains(Path.GetExtension(s), StringComparer.OrdinalIgnoreCase));
        if (index >= 0 && index < segments.Length - 1)
        {
            // we have a compressed file in the path, so we cannot resolve it directly
            var pathToCompressed = string.Join(Path.DirectorySeparatorChar, segments[..(index + 1)]);
            var compressedEntry = Get(null, pathToCompressed, options | ShellItemParsingOptions.DontThrowOnError, context);
            if (compressedEntry is not null)
            {
                var relativePath = string.Join(Path.DirectorySeparatorChar, segments[(index + 1)..]);
                return ZipEntry.GetChild(compressedEntry, relativePath);
            }
        }

        return null;
    }

    public static IComObject<IDataObject>? CreateDataObject(IEnumerable<Entry> entries, bool throwOnError = true)
    {
        var items = entries.Select(e => e.GetItem(throwOnError ? ShellItemParsingOptions.None : ShellItemParsingOptions.DontThrowOnError)).WhereNotNull();
        return ShellItem.CreateDataObject(items, disposeItems: true, throwOnError: throwOnError);
    }

    public static int CompareExtensions(Entry x, Entry y)
    {
        ArgumentNullException.ThrowIfNull(x);
        ArgumentNullException.ThrowIfNull(y);

        if (x.IsFolder && y.IsFolder)
            return ShellN.Functions.StrCmpLogicalW(PWSTR.From(x.DisplayName), PWSTR.From(y.DisplayName));

        return x.Extension.CompareTo(y.Extension);
    }

    public static async Task<WicBitmapSource?> GetImage(string parsingName, int size, GetImageOptions options)
    {
        var item = GetItem(parsingName, ShellItemParsingOptions.DontThrowOnError);
        if (item is null)
        {
            var zipEntry = GetZipEntry(parsingName, ShellItemParsingOptions.None, null);
            if (zipEntry is not null)
            {
                if (zipEntry.Extension.IsImage)
                {
                    using var stream = zipEntry.OpenRead();
                    if (stream is not null)
                    {
                        var bitmap = WicBitmapSource.Load(stream);
                        if (bitmap.Width != size || bitmap.Height != size)
                        {
                            bitmap.Scale(size, null, WICBitmapInterpolationMode.WICBitmapInterpolationModeCubic);
                        }
                        return bitmap;
                    }
                }
                else
                {
                    // create a non-existing entry (use extension to get icon)
                    var attributes = zipEntry.IsFolder ? FileAttributes.Directory : FileAttributes.Normal;
                    var context = IBindCtxExtensions.CreateBindCtx(zipEntry.DisplayName, attributes: attributes);
                    item = GetItem(zipEntry.ParsingName, ShellItemParsingOptions.None, context);
                }
            }
        }

        try
        {
            return await GetImage(item, size, options);
        }
        finally
        {
            item?.Dispose();
        }
    }

    private static async Task<WicBitmapSource?> GetImage(ShellItem? item, int size, GetImageOptions options)
    {
        if (item == null)
            return null;

        var isImage = item.IsImage(out _, out var isWicImage);
        if (isImage && isWicImage)
        {
            item.ComObject.Object.BindToHandler(null, ShellN.Constants.BHID_Stream, typeof(DirectN.IStream).GUID, out var unk);
            if (unk == 0)
            {
                // can happen for items in .zip archive for example
                ShellBatInstance.LogWarning($"Could not get IStream for '{item.SIGDN_DESKTOPABSOLUTEPARSING}'");
                return null;
            }

            var stream = ComObject.FromPointer<DirectN.IStream>(unk);
            if (stream != null)
            {
                var managedStream = new StreamOnIStream(stream.Object, true);
                try
                {
                    var bitmap = WicBitmapSource.Load(managedStream);

                    // FUTURE: implement GetImageOptions
                    if (bitmap.Width != size || bitmap.Height != size)
                    {
                        bitmap.Scale(size, null, WICBitmapInterpolationMode.WICBitmapInterpolationModeCubic);
                    }
                    return bitmap;
                }
                catch
                {
                    // continue
                }
            }
        }

        var flags = SIIGBF.SIIGBF_RESIZETOFIT;
        if (options.HasFlag(GetImageOptions.CropToSquare))
        {
            flags |= SIIGBF.SIIGBF_CROPTOSQUARE;
        }

        if (options.HasFlag(GetImageOptions.CacheOnly))
        {
            flags |= SIIGBF.SIIGBF_INCACHEONLY;
        }

        if (options.HasFlag(GetImageOptions.IconOnly))
        {
            flags |= SIIGBF.SIIGBF_ICONONLY;
        }

        using var image = await item.GetImageAsBitmapAsync(new SIZE(size, size), flags);
        if (image == null)
        {
            ShellBatInstance.LogWarning($"Could not get image for '{item.SIGDN_DESKTOPABSOLUTEPARSING}' flags {flags}");
            return null;
        }

        var bmp = new WicBitmapSource(image);
        if (bmp.PixelFormat != WicPixelFormat.GUID_WICPixelFormat32bppPBGRA)
        {
            bmp.ConvertTo(WicPixelFormat.GUID_WICPixelFormat32bppPBGRA);
        }
        return bmp;
    }

    private static string? NormalizeParsingName(string? parsingName)
    {
        if (string.IsNullOrWhiteSpace(parsingName))
            return null;

        if (parsingName.Length == 40 && parsingName[0] == ':' &&
            parsingName[1] == ':' && parsingName[2] == '{' &&
            parsingName[39] == '}' && Guid.TryParse(parsingName[3..39], out var guid)) // ::{4234d49b-0245-4df3-b780-3893943456e1}
        {
            if (guid == ShellN.Constants.CLSID_Applications)
                return ShellN.Constants.ShellAppsFolder;
        }

        return parsingName;
    }

    public virtual void ExportAsCsv(string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        IOUtilities.FileEnsureDirectory(filePath);

        var settings = ShellBatInstance.Current.Settings;

        using var writer = new StreamWriter(filePath, false, Encoding.Unicode);
        var i = 0;
        foreach (var child in EnumerateChildren(settings.EntryEnumerateOptions).Order(new SortByComparer(settings.SortBy, settings.SortDirection, settings.EntryEnumerateOptions)))
        {
            if (i == 0)
            {
                WriteHeader(writer);
                writer.WriteLine();
                i++;
            }

            child.Write(writer);
            writer.WriteLine();
        }
    }

    private void WriteHeader(TextWriter writer)
    {
        writer.WriteCsvCell(nameof(ParsingName));
        writer.WriteCsvCell(nameof(DisplayName));
        writer.WriteCsvCell(nameof(Extension.Name));
        writer.WriteCsvCell(nameof(FileSystemPath));
        writer.WriteCsvCell(nameof(Size));
        writer.WriteCsvCell(nameof(LastWriteTime));
        writer.WriteCsvCell(nameof(CreationTime));
        writer.WriteCsvCell(nameof(LastAccessTime));
        writer.WriteCsvCell(nameof(FileAttributes));
        writer.WriteCsvCell(nameof(Attributes), false);
    }

    private void Write(TextWriter writer)
    {
        writer.WriteCsvCell(ParsingName);
        writer.WriteCsvCell(DisplayName);
        writer.WriteCsvCell(Extension.Name);
        writer.WriteCsvCell(FileSystemPath);
        writer.WriteCsvCell(Size.ToString());
        writer.WriteCsvCell(LastWriteTime.ToString());
        writer.WriteCsvCell(CreationTime.ToString());
        writer.WriteCsvCell(LastAccessTime.ToString());
        writer.WriteCsvCell(FileAttributes.ToString());
        writer.WriteCsvCell(Attributes.ToString(), false);
    }

    private ShellItem? GetItem(ShellItemParsingOptions options = ShellItemParsingOptions.None, IComObject<IBindCtx>? ctx = null, bool owned = true) => GetItem(ParsingName, options, ctx, owned);
    private ShellFolder? GetAsFolder(ShellItemParsingOptions options = ShellItemParsingOptions.None, IComObject<IBindCtx>? ctx = null, bool owned = true) => GetItem(options, ctx, owned) as ShellFolder;
    public static ShellItem? GetItem(string? parsingName, ShellItemParsingOptions options = ShellItemParsingOptions.None, IComObject<IBindCtx>? context = null, bool owned = true)
    {
        parsingName = NormalizeParsingName(parsingName);
        if (string.IsNullOrWhiteSpace(parsingName))
            return ShellFolder.Desktop;

        var hr = ShellN.Functions.SHCreateItemFromParsingName(PWSTR.From(parsingName), context?.Object, typeof(IShellItem2).GUID, out var ppv);
        if (hr.IsSuccess)
            return ShellItem.FromObject(ppv);

        if (options.HasFlag(ShellItemParsingOptions.TrySplitting) || parsingName.StartsWith("::"))
        {
            var item = GetItemWithSplit(parsingName, context, owned);
            if (item != null)
                return item;
        }

        var throwOnError = !options.HasFlag(ShellItemParsingOptions.DontThrowOnError);
        var invokeDefaultCommand = !options.HasFlag(ShellItemParsingOptions.DontInvokeDefaultCommand);

        // for some reasons, with encrypted drives, the parsing name might fail with this error (not with path/file not found)
        if (invokeDefaultCommand && hr.Value == DirectN.Constants.E_INVALIDARG)
        {
            try
            {
                var root = Path.GetPathRoot(parsingName);
                if (root != null)
                {
                    using var folder = ShellItem.FromParsingName(root, owned: owned, throwOnError: false);
                    if (IsInvokeDefaultCommandRequired(folder))
                    {
                        folder?.InvokeDefaultCommand();
                        return null; // we invoked the default command, so we don't need to return anything
                    }
                }
            }
            catch
            {
                // continue
            }
        }

        hr.ThrowOnError(throwOnError);
        return null;
    }

    private static ShellItem? GetItemWithSplit(string parsingName, IComObject<IBindCtx>? context = null, bool owned = true)
    {
        // note we cannot split the whole thing using \ because some parsing names use \\?\ (like portable devices, etc.)
        var indexOfSep = parsingName.IndexOf(Path.DirectorySeparatorChar);
        if (indexOfSep < 0)
            return null;

        var firstPart = parsingName[..indexOfSep];
        ShellN.Functions.SHCreateItemFromParsingName(PWSTR.From(firstPart), context?.Object, typeof(IShellItem2).GUID, out var ppv);
        if (ppv == 0)
            return null;

        var item = ShellItem.FromObject(ppv, owned: owned);
        if (item == null)
            return null;

        // first one is necessary a folder otherwise it would have been picked up by GetItem
        if (item is not ShellFolder root)
        {
            item.Dispose();
            return null;
        }

        var currentFolder = root;
        do
        {
            var leftPart = parsingName[(indexOfSep + 1)..];

            // equality first
            var child = currentFolder.EnumerateChildrenWhere(i => i.SIGDN_PARENTRELATIVEPARSING == leftPart, owned: owned).FirstOrDefault();
            if (child != null)
            {
                currentFolder.Dispose();
                return child;
            }

            // otherwise go deeper
            child = currentFolder.EnumerateChildrenWhere(i => leftPart.StartsWith(i.SIGDN_PARENTRELATIVEPARSING + Path.DirectorySeparatorChar), owned: owned).FirstOrDefault();
            if (child == null || child is not ShellFolder childFolder)
            {
                currentFolder.Dispose();
                return null;
            }

            var relativeName = child.SIGDN_PARENTRELATIVEPARSING;
            if (relativeName == null)
            {
                child.Dispose();
                currentFolder.Dispose();
                return null; // should not happen
            }

            indexOfSep += relativeName.Length + 1;
            currentFolder?.Dispose();
            currentFolder = childFolder;
        }
        while (true);
    }

    private static bool IsInvokeDefaultCommandRequired(ShellItem? item)
    {
        if (item == null)
            return false;

        if (item.IsFolder)
            return item.EncryptionFlags == 6; // bitlocker encrypted drive

        return false;
    }

    private static _SHCONTF GetFlags(EntryEnumerateOptions options = EntryEnumerateOptions.None)
    {
        var flags = _SHCONTF.SHCONTF_FOLDERS | _SHCONTF.SHCONTF_NONFOLDERS;
        if (options.HasFlag(EntryEnumerateOptions.IncludeHidden))
        {
            flags |= _SHCONTF.SHCONTF_INCLUDEHIDDEN;
        }

        if (options.HasFlag(EntryEnumerateOptions.IncludeSystem))
        {
            flags |= _SHCONTF.SHCONTF_INCLUDESUPERHIDDEN;
        }

        if (options.HasFlag(EntryEnumerateOptions.ExcludeFolders))
        {
            flags &= ~_SHCONTF.SHCONTF_FOLDERS;
        }

        if (options.HasFlag(EntryEnumerateOptions.ExcludeFiles))
        {
            flags &= ~_SHCONTF.SHCONTF_NONFOLDERS;
        }
        return flags;
    }

    [System.Runtime.InteropServices.Marshalling.GeneratedComClass]
    private sealed partial class ZipEntry : Entry
    {
        private readonly List<ZipEntry> _children = [];
        private readonly ZipArchiveEntry? _archiveEntry;

        private ZipEntry(
            Entry parent,
            ZipArchiveEntry? archiveEntry,
            string parsingName,
            string displayName,
            FileExtension extension
            ) : base(parent, parsingName, displayName, parsingName, extension)
        {
            _archiveEntry = archiveEntry;
            if (archiveEntry != null)
            {
                Size = archiveEntry.Length;
                LastWriteTime = archiveEntry.LastWriteTime.DateTime;
            }
            else
            {
                Attributes = SFGAO_FLAGS.SFGAO_FOLDER;
            }
        }

        public override bool IsNewFolderSupported => false;

        public override Stream? OpenRead()
        {
            if (_archiveEntry == null)
                return null;

            return _archiveEntry.Open();
        }

        public override IEnumerable<Entry> EnumerateChildren(EntryEnumerateOptions options = EntryEnumerateOptions.None)
        {
            foreach (var child in _children)
            {
                if (!child.Extension.IsImage && options.HasFlag(EntryEnumerateOptions.ImagesOnly))
                    continue;

                if (options.HasFlag(EntryEnumerateOptions.ExcludeFolders) && child._children.Count > 0)
                    continue;

                if (options.HasFlag(EntryEnumerateOptions.ExcludeFiles) && child._children.Count == 0)
                    continue;

                // FUTURE what about EntryEnumerateOptions.ShowCompressedAsFile?
                yield return child;
            }
        }

        public static Entry? GetChild(Entry entry, string relativePath)
        {
            var archive = MemoryZipArchive.GetArchive(entry);
            if (archive == null)
                return null;

            archive.Entries.TryGetValue(relativePath, out var zipEntry);
            return zipEntry;
        }

        public static IEnumerable<Entry> EnumerateCompressedChildren(Entry entry, EntryEnumerateOptions options = EntryEnumerateOptions.None)
        {
            var archive = MemoryZipArchive.GetArchive(entry);
            if (archive == null)
                yield break;

            foreach (var child in archive.Root.EnumerateChildren(options))
            {
                yield return child;
            }
        }

        private sealed partial class MemoryZipArchive(Entry entry) : IDisposable
        {
            private static readonly ConcurrentDictionary<string, MemoryZipArchive?> _archives = new(StringComparer.OrdinalIgnoreCase);

            public ZipArchive Archive { get; private set; } = null!;
            public DateTime CreationTime { get; } = DateTime.Now;
            public ZipEntry Root { get; } = new ZipEntry(entry, null, string.Empty, string.Empty, FileExtension.None);
            public IDictionary<string, ZipEntry> Entries { get; } = new Dictionary<string, ZipEntry>(StringComparer.OrdinalIgnoreCase);

            public void Dispose() => Archive.Dispose();

            public static MemoryZipArchive? GetArchive(Entry entry)
            {
                var archive = _archives.GetOrAdd(entry.ParsingName, p =>
                {
                    if (!entry.Extension.CanBeDecompressed)
                        return null;

                    // https://stackoverflow.com/a/2448637/403671
                    if (entry.Size < 22)
                        return null;

                    var stream = entry.OpenRead();
                    if (stream == null)
                        return null;

                    var archive = new MemoryZipArchive(entry) { Archive = new ZipArchive(stream, ZipArchiveMode.Read, false) };
                    archive.Entries[string.Empty] = archive.Root;

                    foreach (var archiveEntry in archive.Archive.Entries)
                    {
                        var extension = FileExtension.GetByExtension(Path.GetExtension(archiveEntry.Name)) ?? FileExtension.None;
                        var fn = archiveEntry.FullName.Replace('/', '\\');
                        var dir = Path.GetDirectoryName(fn) ?? string.Empty;
                        var parent = ensureEntry(dir);

                        var ze = new ZipEntry(parent, archiveEntry, Path.Combine(entry.ParsingName, fn), archiveEntry.Name, extension);
                        if (parent is ZipEntry zipEntry)
                        {
                            zipEntry._children.Add(ze);
                        }
                        else
                        {
                            archive.Root._children.Add(ze);
                        }
                        archive.Entries[fn] = ze;
                    }

                    Entry ensureEntry(string path)
                    {
                        if (path == string.Empty)
                            return entry;

                        if (archive.Entries.TryGetValue(path, out var zipEntry))
                            return zipEntry;

                        var parent = ensureEntry(Path.GetDirectoryName(path) ?? string.Empty);
                        var extension = FileExtension.GetByExtension(Path.GetExtension(path)) ?? FileExtension.None;
                        zipEntry = new ZipEntry(parent, null, Path.Combine(entry.ParsingName, path), Path.GetFileName(path), extension);
                        if (parent is ZipEntry ze)
                        {
                            ze._children.Add(zipEntry);
                        }
                        else
                        {
                            archive.Root._children.Add(zipEntry);
                        }

                        archive.Entries[path] = zipEntry;
                        return zipEntry;
                    }

                    return archive;
                });

                // cache expiration
                Task.Run(() =>
                {
                    // REVIEW: check all this
                    if (_archives.Count < 2)
                        return;

                    var toRemove = _archives.Where(kvp => (DateTime.Now - kvp.Value!.CreationTime).TotalMinutes > 10).Select(kvp => kvp.Key).ToList();
                    foreach (var key in toRemove)
                    {
                        if (_archives.TryRemove(key, out var removed))
                        {
                            removed?.Dispose();
                        }
                    }
                });
                return archive;
            }
        }
    }
}

