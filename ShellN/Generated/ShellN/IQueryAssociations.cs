#nullable enable
namespace ShellN;

// https://learn.microsoft.com/windows/win32/api/shlwapi/nn-shlwapi-iqueryassociations
[SupportedOSPlatform("windows5.0")]
[GeneratedComInterface, Guid("c46ca590-3c3f-11d2-bee6-0000f805ca57")]
public partial interface IQueryAssociations
{
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iqueryassociations-init
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Init(ASSOCF flags, PWSTR pszAssoc, HKEY hkProgid, HWND hwnd);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iqueryassociations-getstring
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetString(ASSOCF flags, ASSOCSTR str, PWSTR pszExtra, [MarshalUsing(CountElementName = nameof(pcchOut))] PWSTR pszOut, ref uint pcchOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iqueryassociations-getkey
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetKey(ASSOCF flags, ASSOCKEY key, PWSTR pszExtra, out HKEY phkeyOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iqueryassociations-getdata
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetData(ASSOCF flags, ASSOCDATA data, PWSTR pszExtra, nint /* optional void* */ pvOut, nint /* optional uint* */ pcbOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iqueryassociations-getenum
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetEnum(ASSOCF flags, ASSOCENUM assocenum, PWSTR pszExtra, in Guid riid, out nint /* void */ ppvOut);
}
