#nullable enable
namespace ShellN;

// https://learn.microsoft.com/windows/win32/api/shobjidl/nn-shobjidl-inamespacetreecontrolevents
[SupportedOSPlatform("windows6.0.6000")]
[GeneratedComInterface, Guid("93d77985-b3d8-4484-8318-672cdda002ce")]
public partial interface INameSpaceTreeControlEvents
{
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onitemclick
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnItemClick([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, uint nstceHitTest, uint nstceClickType);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onpropertyitemcommit
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnPropertyItemCommit([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onitemstatechanging
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnItemStateChanging([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, uint nstcisMask, uint nstcisState);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onitemstatechanged
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnItemStateChanged([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, uint nstcisMask, uint nstcisState);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onselectionchanged
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnSelectionChanged([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItemArray>))] IShellItemArray psiaSelection);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onkeyboardinput
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnKeyboardInput(uint uMsg, WPARAM wParam, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onbeforeexpand
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnBeforeExpand([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onafterexpand
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnAfterExpand([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onbeginlabeledit
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnBeginLabelEdit([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onendlabeledit
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnEndLabelEdit([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-ongettooltip
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnGetToolTip([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, [MarshalUsing(CountElementName = nameof(cchTip))] PWSTR pszTip, int cchTip);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onbeforeitemdelete
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnBeforeItemDelete([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onitemadded
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnItemAdded([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, BOOL fIsRoot);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onitemdeleted
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnItemDeleted([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, BOOL fIsRoot);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onbeforecontextmenu
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnBeforeContextMenu([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem?>))] IShellItem? psi, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onaftercontextmenu
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnAfterContextMenu([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem?>))] IShellItem? psi, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IContextMenu>))] IContextMenu pcmIn, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-inamespacetreecontrolevents-onbeforestateimagechange
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnBeforeStateImageChange([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi);
    
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnGetDefaultIconIndex([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, out int piDefaultIcon, out int piOpenIcon);
}
