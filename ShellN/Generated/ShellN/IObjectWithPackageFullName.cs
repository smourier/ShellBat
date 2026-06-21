#nullable enable
namespace ShellN;

[GeneratedComInterface, Guid("ed2aa515-602f-469c-a130-ce69fd0fa878")]
public partial interface IObjectWithPackageFullName
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetPackageFullName(out PWSTR packageFullName);
}
