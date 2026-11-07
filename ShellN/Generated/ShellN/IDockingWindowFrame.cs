#nullable enable
namespace ShellN;

// https://learn.microsoft.com/windows/win32/api/shlobj/nn-shlobj-idockingwindowframe
[SupportedOSPlatform("windows5.0")]
[GeneratedComInterface, Guid("47d2657a-7b27-11d0-8ca9-00a0c92dbfe8")]
public partial interface IDockingWindowFrame : IOleWindow
{
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-idockingwindowframe-addtoolbar
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT AddToolbar(nint punkSrc, PWSTR pwszItem, uint dwAddFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-idockingwindowframe-removetoolbar
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT RemoveToolbar(nint punkSrc, uint dwRemoveFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-idockingwindowframe-findtoolbar
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT FindToolbar(PWSTR pwszItem, in Guid riid, out nint /* void */ ppv);
}
