#nullable enable
namespace ShellN;

// https://learn.microsoft.com/windows/win32/api/commoncontrols/nn-commoncontrols-iimagelist
[SupportedOSPlatform("windows6.0.6000")]
[GeneratedComInterface, Guid("46eb5926-582e-4017-9fdf-e8998daa0950")]
public partial interface IImageList
{
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-add
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Add(HBITMAP hbmImage, HBITMAP hbmMask, out int pi);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-replaceicon
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT ReplaceIcon(int i, HICON hicon, out int pi);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-setoverlayimage
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetOverlayImage(int iImage, int iOverlay);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-replace
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Replace(int i, HBITMAP hbmImage, HBITMAP hbmMask);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-addmasked
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT AddMasked(HBITMAP hbmImage, COLORREF crMask, out int pi);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-draw
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Draw(in IMAGELISTDRAWPARAMS pimldp);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-remove
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Remove(int i);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-geticon
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetIcon(int i, uint flags, out HICON picon);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-getimageinfo
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetImageInfo(int i, out IMAGEINFO pImageInfo);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-copy
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Copy(int iDst, nint punkSrc, int iSrc, uint uFlags);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-merge
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Merge(int i1, nint punk2, int i2, int dx, int dy, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-clone
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Clone(in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-getimagerect
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetImageRect(int i, out RECT prc);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-geticonsize
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetIconSize(out int cx, out int cy);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-seticonsize
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetIconSize(int cx, int cy);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-getimagecount
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetImageCount(out int pi);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-setimagecount
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetImageCount(uint uNewCount);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-setbkcolor
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetBkColor(COLORREF clrBk, out COLORREF pclr);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-getbkcolor
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetBkColor(out COLORREF pclr);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-begindrag
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT BeginDrag(int iTrack, int dxHotspot, int dyHotspot);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-enddrag
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT EndDrag();
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-dragenter
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT DragEnter(HWND hwndLock, int x, int y);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-dragleave
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT DragLeave(HWND hwndLock);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-dragmove
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT DragMove(int x, int y);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-setdragcursorimage
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetDragCursorImage(nint punk, int iDrag, int dxHotspot, int dyHotspot);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-dragshownolock
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT DragShowNolock(BOOL fShow);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-getdragimage
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetDragImage(nint /* optional POINT* */ ppt, nint /* optional POINT* */ pptHotspot, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-getitemflags
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetItemFlags(int i, out IMAGE_LIST_ITEM_FLAGS dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/commoncontrols/nf-commoncontrols-iimagelist-getoverlayimage
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetOverlayImage(int iOverlay, out int piIndex);
}
