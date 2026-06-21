#nullable enable
namespace ShellN;

[GeneratedComInterface, Guid("4f2b781f-a608-4543-abf0-49c246ebbba9")]
public partial interface IAttachmentExecute2 : IAttachmentExecute
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SaveNoVirusCheck();
    
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SaveWithUINoVirusCheck(HWND hwnd);
}
