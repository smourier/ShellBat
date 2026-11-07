#nullable enable
namespace ShellN;

// https://learn.microsoft.com/windows/win32/api/shobjidl_core/nn-shobjidl_core-ishellmenu
[SupportedOSPlatform("windows5.1.2600")]
[GeneratedComInterface, Guid("ee1f7637-e138-11d1-8379-00c04fd918d0")]
public partial interface IShellMenu
{
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellmenu-initialize
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Initialize([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellMenuCallback?>))] IShellMenuCallback? psmc, uint uId, uint uIdAncestor, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellmenu-getmenuinfo
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetMenuInfo(nint /* optional IShellMenuCallback* */ ppsmc, nint /* optional uint* */ puId, nint /* optional uint* */ puIdAncestor, nint /* optional uint* */ pdwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellmenu-setshellfolder
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetShellFolder([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder?>))] IShellFolder? psf, nint /* optional nint* */ pidlFolder, HKEY hKey, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellmenu-getshellfolder
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetShellFolder(out uint pdwFlags, out nint ppidl, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellmenu-setmenu
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetMenu(HMENU hmenu, HWND hwnd, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellmenu-getmenu
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetMenu(nint /* optional HMENU* */ phmenu, nint /* optional HWND* */ phwnd, nint /* optional uint* */ pdwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellmenu-invalidateitem
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT InvalidateItem(nint /* optional SMDATA* */ psmd, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellmenu-getstate
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetState(out SMDATA psmd);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellmenu-setmenutoolbar
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetMenuToolbar(nint punk, uint dwFlags);
}
