#nullable enable
namespace ShellN;

// https://learn.microsoft.com/windows/win32/api/shdeprecated/nn-shdeprecated-ibrowserservice2
[SupportedOSPlatform("windows5.0")]
[GeneratedComInterface, Guid("68bd21cc-438b-11d2-a560-00a0c92dbfe8")]
public partial interface IBrowserService2 : IBrowserService
{
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-wndprocbs
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    LRESULT WndProcBS(HWND hwnd, uint uMsg, WPARAM wParam, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-setasdeffoldersettings
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetAsDefFolderSettings();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-getviewrect
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetViewRect(out RECT prc);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-onsize
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnSize(WPARAM wParam);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-oncreate
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnCreate(in CREATESTRUCTW pcs);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-oncommand
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    LRESULT OnCommand(WPARAM wParam, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-ondestroy
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnDestroy();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-onnotify
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    LRESULT OnNotify(in NMHDR pnm);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-onsetfocus
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnSetFocus();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-onframewindowactivatebs
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OnFrameWindowActivateBS(BOOL fActive);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-releaseshellview
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT ReleaseShellView();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-activatependingview
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT ActivatePendingView();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-createviewwindow
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT CreateViewWindow([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellView>))] IShellView psvNew, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellView>))] IShellView psvOld, in RECT prcView, out HWND phwnd);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-createbrowserpropsheetext
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT CreateBrowserPropSheetExt(in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-getviewwindow
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetViewWindow(out HWND phwndView);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-getbasebrowserdata
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetBaseBrowserData(out nint pbbd);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-putbasebrowserdata
    [PreserveSig]
    nint PutBaseBrowserData();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-initializetravellog
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT InitializeTravelLog([MarshalUsing(typeof(UniqueComInterfaceMarshaller<ITravelLog>))] ITravelLog ptl, uint dw);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-settopbrowser
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetTopBrowser();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-offline
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Offline(int iCmd);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-allowviewresize
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT AllowViewResize(BOOL f);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-setactivatestate
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetActivateState(uint u);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-updatesecurelockicon
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT UpdateSecureLockIcon(int eSecureLock);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-initializedownloadmanager
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT InitializeDownloadManager();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-initializetransitionsite
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT InitializeTransitionSite();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_initialize
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _Initialize(HWND hwnd, nint pauto);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_cancelpendingnavigationasync
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _CancelPendingNavigationAsync();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_cancelpendingview
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _CancelPendingView();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_maysavechanges
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _MaySaveChanges();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_pauseorresumeview
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _PauseOrResumeView(BOOL fPaused);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_disablemodeless
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _DisableModeless();
    
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _NavigateToPidl2(nint pidl, uint grfHLNF, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_tryshell2rename
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _TryShell2Rename([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellView>))] IShellView psv, nint pidlNew);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_switchactivationnow
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _SwitchActivationNow();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_execchildren
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _ExecChildren(nint punkBar, BOOL fBroadcast, nint /* optional Guid* */ pguidCmdGroup, uint nCmdID, uint nCmdexecopt, nint /* optional VARIANT* */ pvarargIn, nint /* optional VARIANT* */ pvarargOut);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_sendchildren
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _SendChildren(HWND hwndBar, BOOL fBroadcast, uint uMsg, WPARAM wParam, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-getfoldersetdata
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetFolderSetData(ref FOLDERSETDATA pfsd);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_onfocuschange
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _OnFocusChange(uint itb);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-v_showhidechildwindows
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT v_ShowHideChildWindows(BOOL fChildOnly);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_get_itblastfocus
    [PreserveSig]
    uint _get_itbLastFocus();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_put_itblastfocus
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _put_itbLastFocus(uint itbLastFocus);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_uiactivateview
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _UIActivateView(uint uState);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_getviewborderrect
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _GetViewBorderRect(ref RECT prc);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_updateviewrectsize
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _UpdateViewRectSize();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_resizenextborder
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _ResizeNextBorder(uint itb);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_resizeview
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _ResizeView();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_geteffectiveclientarea
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _GetEffectiveClientArea(out RECT lprectBorder, HMONITOR hmon);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-v_getviewstream
    [PreserveSig]
    [return: MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))]
    IStream v_GetViewStream(ref nint pidl, uint grfMode, PWSTR pwszName);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-forwardviewmsg
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    LRESULT ForwardViewMsg(uint uMsg, WPARAM wParam, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-setacceleratormenu
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetAcceleratorMenu(HACCEL hacc);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_gettoolbarcount
    [PreserveSig]
    int _GetToolbarCount();
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_gettoolbaritem
    [PreserveSig]
    nint _GetToolbarItem(int itb);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_savetoolbars
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _SaveToolbars([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_loadtoolbars
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _LoadToolbars([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_closeandreleasetoolbars
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _CloseAndReleaseToolbars(BOOL fClose);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-v_maygetnexttoolbarfocus
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT v_MayGetNextToolbarFocus(in MSG lpMsg, uint itbNext, int citb, out nint pptbi, out HWND phwnd);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_resizenextborderhelper
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _ResizeNextBorderHelper(uint itb, BOOL bUseHmonitor);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_findtbar
    [PreserveSig]
    uint _FindTBar(nint punkSrc);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_setfocus
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _SetFocus(in TOOLBARITEM ptbi, HWND hwnd, in MSG lpMsg);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-v_maytranslateaccelerator
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT v_MayTranslateAccelerator(ref MSG pmsg);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-_getborderdwhelper
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT _GetBorderDWHelper(nint punkSrc, out RECT lprectBorder, BOOL bUseHmonitor);
    
    // https://learn.microsoft.com/windows/win32/api/shdeprecated/nf-shdeprecated-ibrowserservice2-v_checkzonecrossing
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT v_CheckZoneCrossing(nint pidl);
}
