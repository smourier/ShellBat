namespace ShellN;

public static partial class Functions
{
    [LibraryImport("shlwapi", EntryPoint = "#279")]
    [PreserveSig]
    public static partial HRESULT SHInvokeDefaultCommand(HWND hwnd, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder>))] IShellFolder psf, nint pidl);

    [LibraryImport("shlwapi", EntryPoint = "#363")]
    [PreserveSig]
    public static partial HRESULT SHInvokeCommand(HWND hwnd, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder>))] IShellFolder psf, nint pidl, PWSTR lpVerb);

    [LibraryImport("shlwapi", EntryPoint = "#186")]
    [PreserveSig]
    public static partial HRESULT SHSimulateDrop([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDropTarget>))] IDropTarget dropTarget, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject>))] IDataObject dataObject, MODIFIERKEYS_FLAGS flags, ref POINTL pt, ref DROPEFFECT pdwEffect);

    [LibraryImport("shell32")]
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool SHELL32_SHIsVirtualDevice(HANDLE handle);

    [LibraryImport("shell32")]
    [PreserveSig]
    public static partial HRESULT SHELL32_TryVirtualDiscImageDriveEject(HANDLE handle, HWND hwnd, [MarshalAs(UnmanagedType.Bool)] out bool isVirtual);

    [LibraryImport("Windows.Storage.dll")]
    [PreserveSig]
    public static partial HRESULT CreateStorageItemFromShellItem(nint item, in STORAGEITEM_FROM_SHELLITEM_CREATE_OPTIONS options, in Guid riid, out nint ppv);

    [LibraryImport("Windows.Storage.dll", EntryPoint = "STORAGE_GetShellItemFromStorageItem")]
    [PreserveSig]
    public static partial HRESULT GetShellItemFromStorageItem(nint storageItem, in Guid riid, out nint ppv);

    [LibraryImport("Windows.Storage.dll", EntryPoint = "STORAGE_CStorageItem_GetValidatedStorageItem")]
    [PreserveSig]
    public static partial HRESULT GetValidatedStorageItem(nint storageItem, in Guid riid, out nint ppv);

    [LibraryImport("Windows.Storage.dll", EntryPoint = "STORAGE_CStorageItem_GetValidatedStorageItemObject")]
    [PreserveSig]
    public static partial HRESULT GetValidatedStorageItemObject(nint site, nint storageItem, in Guid riid, out nint ppv);
}