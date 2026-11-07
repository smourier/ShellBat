#nullable enable
namespace ShellN;

// https://learn.microsoft.com/windows/win32/api/shobjidl_core/nn-shobjidl_core-itransferdestination
[SupportedOSPlatform("windows6.0.6000")]
[GeneratedComInterface, Guid("48addd32-3ca5-4124-abe3-b5a72531b207")]
public partial interface ITransferDestination
{
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransferdestination-advise
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Advise([MarshalUsing(typeof(UniqueComInterfaceMarshaller<ITransferAdviseSink>))] ITransferAdviseSink psink, out uint pdwCookie);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransferdestination-unadvise
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Unadvise(uint dwCookie);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-itransferdestination-createitem
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT CreateItem(PWSTR pszName, uint dwAttributes, ulong ullSize, uint flags, in Guid riidItem, out nint /* void */ ppvItem, in Guid riidResources, out nint /* void */ ppvResources);
}
