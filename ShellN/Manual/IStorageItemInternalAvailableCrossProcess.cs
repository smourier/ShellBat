namespace ShellN;

[GeneratedComInterface, Guid("d75c13bb-3883-4bd2-9b7a-334ff6d83066")]
public partial interface IStorageItemInternalAvailableCrossProcess
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetShellItem(in Guid riid, out nint ppv);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_CreationFlags(out STORAGEITEM_CREATION_FLAGS flags);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetCreationFlagsAndShellItem(out STORAGEITEM_CREATION_FLAGS flags, in Guid riid, out nint ppv);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_CreatorPackageFamilyName(out PWSTR name);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_Condition([MarshalUsing(typeof(UniqueComInterfaceMarshaller<ICondition>))] out ICondition condition);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT CloneAsReadOnly(in Guid riid, out nint ppv);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SerializeToSystemAccessList(PWSTR name, out PWSTR serialized);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT AddStorageProviderProperties(in PROPERTYKEY key, uint count, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyStore>))] IPropertyStore store);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetStorageProviderProperties([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyStore>))] IPropertyStore store);
}
