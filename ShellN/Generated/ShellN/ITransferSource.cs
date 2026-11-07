#nullable enable
namespace ShellN;

// https://learn.microsoft.com/windows/win32/api/shobjidl_core/nn-shobjidl_core-itransfersource
[SupportedOSPlatform("windows6.0.6000")]
[GeneratedComInterface, Guid("00adb003-bde9-45c6-8e29-d09f9353e108")]
public partial interface ITransferSource
{
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-advise
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Advise([MarshalUsing(typeof(UniqueComInterfaceMarshaller<ITransferAdviseSink>))] ITransferAdviseSink psink, out uint pdwCookie);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-unadvise
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Unadvise(uint dwCookie);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-setproperties
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetProperties([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyChangeArray>))] IPropertyChangeArray pproparray);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-openitem
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OpenItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, uint flags, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-moveitem
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT MoveItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiParentDst, PWSTR pszNameDst, uint flags, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] out IShellItem ppsiNew);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-recycleitem
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT RecycleItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiSource, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiParentDest, uint flags, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] out IShellItem ppsiNewDest);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-removeitem
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT RemoveItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiSource, uint flags);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-renameitem
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT RenameItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiSource, PWSTR pszNewName, uint flags, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] out IShellItem ppsiNewDest);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-linkitem
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT LinkItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiSource, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiParentDest, PWSTR pszNewName, uint flags, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] out IShellItem ppsiNewDest);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-applypropertiestoitem
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT ApplyPropertiesToItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiSource, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] out IShellItem ppsiNew);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-getdefaultdestinationname
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetDefaultDestinationName([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiSource, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiParentDest, out PWSTR ppszDestinationName);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-enterfolder
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT EnterFolder([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiChildFolderDest);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransfersource-leavefolder
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT LeaveFolder([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiChildFolderDest);
}
