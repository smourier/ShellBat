#nullable enable
namespace ShellN;

[GeneratedComInterface, Guid("4b770da6-d111-4015-96fd-8c1c56f06c55")]
public partial interface IStaticVerbProvider
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT IsVerbSupported(PWSTR verbName, out BOOL result);
}
