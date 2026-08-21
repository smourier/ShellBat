namespace ShellN.Extensions;

public class MenuItem : IDisposable
{
    private bool _disposedValue;

    public virtual uint Id { get; protected set; }
    public virtual uint Index { get; protected set; }
    public virtual MENU_ITEM_TYPE Type { get; protected set; }
    public virtual MENU_ITEM_STATE State { get; protected set; }
    public virtual string? Text { get; protected set; } = string.Empty;
    public virtual HBITMAP BitmapHandle { get; protected set; }
    public virtual IComObject<IWICBitmap>? Bitmap { get; protected set; }
    public virtual string? Verb { get; protected set; }
    public virtual IList<MenuItem> Children { get; } = [];
    public bool IsSeparator => Type.HasFlag(MENU_ITEM_TYPE.MFT_SEPARATOR);
    public bool IsRadioCheck => Type.HasFlag(MENU_ITEM_TYPE.MFT_RADIOCHECK);
    public IEnumerable<MenuItem> AllChildren
    {
        get
        {
            foreach (var child in Children)
            {
                yield return child;
                foreach (var grandChild in child.AllChildren)
                {
                    yield return grandChild;
                }
            }
        }
    }

    public override string ToString() => IsSeparator ? "-" : Text ?? Id.ToString();

    public virtual void Extract(HMENU menu, IComObject<IContextMenu2> contextMenu)
    {
        IComObject<IWICImagingFactory>? factory = null;
        Extract(menu, contextMenu, ref factory);
        factory?.Dispose();
    }

    private unsafe void Extract(HMENU menu, IComObject<IContextMenu2> contextMenu, ref IComObject<IWICImagingFactory>? imagingFactory)
    {
        ArgumentNullException.ThrowIfNull(contextMenu);
        if (menu.Value == 0)
            throw new ArgumentException(null, nameof(menu));

        var count = Functions.GetMenuItemCount(menu);
        for (uint i = 0; i < count; i++)
        {
            var mii = new MENUITEMINFOW
            {
                cbSize = (uint)sizeof(MENUITEMINFOW),
                fMask = MENU_ITEM_MASK.MIIM_FTYPE |
                    MENU_ITEM_MASK.MIIM_ID |
                    MENU_ITEM_MASK.MIIM_STATE |
                    MENU_ITEM_MASK.MIIM_STRING |
                    MENU_ITEM_MASK.MIIM_SUBMENU |
                    MENU_ITEM_MASK.MIIM_CHECKMARKS |
                    MENU_ITEM_MASK.MIIM_BITMAP |
                    MENU_ITEM_MASK.MIIM_DATA
            };
            if (!Functions.GetMenuItemInfoW(menu, i, true, ref mii))
                continue;

            var item = new MenuItem();
            if (mii.fType == MENU_ITEM_TYPE.MFT_STRING)
            {
                using var pwstr = new AllocPwstr((mii.cch + 1) * 2);
                mii.dwTypeData = pwstr;
                mii.cch++;
                if (!Functions.GetMenuItemInfoW(menu, i, true, ref mii))
                    continue;

                item.Text = pwstr.ToString();
            }

            if (mii.hSubMenu.Value != 0)
            {
                item.Extract(mii.hSubMenu, contextMenu, ref imagingFactory);
            }

            item.Id = mii.wID;
            item.Type = mii.fType;
            item.State = mii.fState;
            item.Index = (uint)Children.Count;
            item.BitmapHandle = mii.hbmpItem;
            if (IsBitmap(item.BitmapHandle))
            {
                item.Bitmap = getBitmap(item.BitmapHandle, ref imagingFactory);
            }
            else if (item.State.HasFlag(MENU_ITEM_STATE.MFS_CHECKED) && IsBitmap(mii.hbmpChecked))
            {
                item.Bitmap = getBitmap(mii.hbmpChecked, ref imagingFactory);
            }
            else if (item.State.HasFlag(MENU_ITEM_STATE.MFS_UNCHECKED) && IsBitmap(mii.hbmpUnchecked))
            {
                item.Bitmap = getBitmap(mii.hbmpUnchecked, ref imagingFactory);
            }

            if (!item.IsSeparator)
            {
                item.Verb = GetCommandString(contextMenu.Object, item.Id);
            }

            Children.Add(item);

            unsafe static ComObject<IWICBitmap>? getBitmap(HBITMAP h, ref IComObject<IWICImagingFactory>? fac)
            {
                // note sure this is the best way to check if the handle needs alpha or not...
                BITMAP b;
                _ = DirectN.Functions.GetObjectW(new HGDIOBJ(h), sizeof(BITMAP), (nint)(&b));

                fac ??= WicImagingFactory.Create()!;
                var hr = fac.Object.CreateBitmapFromHBITMAP(h, 0, b.bmBits == 0 ? WICBitmapAlphaChannelOption.WICBitmapIgnoreAlpha : WICBitmapAlphaChannelOption.WICBitmapUsePremultipliedAlpha, out var bitmap);
                if (bitmap == null || hr.IsError)
                    return null;

                return new ComObject<IWICBitmap>(bitmap);
            }
        }
    }

    public static string? GetCommandString(IContextMenu cm, uint itemId)
    {
        ArgumentNullException.ThrowIfNull(cm);
        // avoid crashes with invalid item IDs
        if (itemId > short.MaxValue || itemId < 0)
            return null;

        const int size = 256;
        var ptr = Marshal.AllocCoTaskMem(2 * size);
        try
        {
            if (cm.GetCommandString(new nuint(NormalizeId(itemId)), Constants.GCS_VERBW, 0, ptr, size).IsSuccess)
                return Marshal.PtrToStringUni(ptr);
        }
        finally
        {
            Marshal.FreeCoTaskMem(ptr);
        }

        ptr = Marshal.AllocCoTaskMem(size);
        try
        {
            if (cm.GetCommandString(new nuint(NormalizeId(itemId)), Constants.GCS_VERBA, 0, ptr, size).IsSuccess)
                return Marshal.PtrToStringAnsi(ptr);
        }
        finally
        {
            Marshal.FreeCoTaskMem(ptr);
        }
        return null;
    }

    private static uint NormalizeId(uint id)
    {
        if (id > Constants.FCIDM_SHVIEWLAST)
            return Constants.FCIDM_SHVIEWLAST;

        return id;
    }

    private static bool IsBitmap(nint bitmapHandle)
    {
        if (bitmapHandle == 0)
            return false;

        foreach (var handle in Enum.GetValues<HBMMENU>())
        {
            if (bitmapHandle == new nint((int)handle))
                return false;
        }
        return true;
    }

    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                Bitmap?.Dispose();
                Bitmap = null;
            }

            _disposedValue = true;
        }
    }
}
