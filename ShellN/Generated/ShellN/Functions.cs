#nullable enable
namespace ShellN;

public static partial class Functions
{
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-assoccreate
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT AssocCreate(Guid clsid, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-assoccreateforclasses
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT AssocCreateForClasses([In][MarshalUsing(CountElementName = nameof(cClasses))] ASSOCIATIONELEMENT[] rgClasses, uint cClasses, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-assocgetdetailsofpropkey
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT AssocGetDetailsOfPropKey([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder>))] IShellFolder psf, nint pidl, in PROPERTYKEY pkey, out VARIANT pv, nint /* optional BOOL* */ pfFoundPropKey);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-assocgetperceivedtype
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT AssocGetPerceivedType(PWSTR pszExt, out PERCEIVED ptype, out uint pflag, nint /* optional PWSTR* */ ppszType);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-associsdangerous
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL AssocIsDangerous(PWSTR pszAssoc);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-assocquerykeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT AssocQueryKeyA(ASSOCF flags, ASSOCKEY key, PSTR pszAssoc, PSTR pszExtra, out HKEY phkeyOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-assocquerykeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT AssocQueryKeyW(ASSOCF flags, ASSOCKEY key, PWSTR pszAssoc, PWSTR pszExtra, out HKEY phkeyOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-assocquerystringa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT AssocQueryStringA(ASSOCF flags, ASSOCSTR str, PSTR pszAssoc, PSTR pszExtra, [MarshalUsing(CountElementName = nameof(pcchOut))] PSTR pszOut, ref uint pcchOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-assocquerystringbykeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT AssocQueryStringByKeyA(ASSOCF flags, ASSOCSTR str, HKEY hkAssoc, PSTR pszExtra, [MarshalUsing(CountElementName = nameof(pcchOut))] PSTR pszOut, ref uint pcchOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-assocquerystringbykeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT AssocQueryStringByKeyW(ASSOCF flags, ASSOCSTR str, HKEY hkAssoc, PWSTR pszExtra, [MarshalUsing(CountElementName = nameof(pcchOut))] PWSTR pszOut, ref uint pcchOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-assocquerystringw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT AssocQueryStringW(ASSOCF flags, ASSOCSTR str, PWSTR pszAssoc, PWSTR pszExtra, [MarshalUsing(CountElementName = nameof(pcchOut))] PWSTR pszOut, ref uint pcchOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-cdeffoldermenu_create2
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT CDefFolderMenu_Create2(nint /* optional nint* */ pidlFolder, HWND hwnd, uint cidl, nint /* optional nint** */ apidl, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder?>))] IShellFolder? psf, LPFNDFMCALLBACK? pfn, uint nKeys, nint /* optional HKEY* */ ahkeys, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IContextMenu>))] out IContextMenu ppcm);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-chrcmpia
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ChrCmpIA(ushort w1, ushort w2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-chrcmpiw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ChrCmpIW(char w1, char w2);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-cidldata_createfromidarray
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT CIDLData_CreateFromIDArray(nint pidlFolder, uint cidl, nint /* optional nint** */ apidl, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject>))] out IDataObject ppdtobj);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-coloradjustluma
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial COLORREF ColorAdjustLuma(COLORREF clrRGB, int n, BOOL fScale);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-colorhlstorgb
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial COLORREF ColorHLSToRGB(ushort wHue, ushort wLuminance, ushort wSaturation);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-colorrgbtohls
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void ColorRGBToHLS(COLORREF clrRGB, out ushort pwHue, out ushort pwLuminance, out ushort pwSaturation);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-commandlinetoargvw
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial nint CommandLineToArgvW(PWSTR lpCmdLine, out int pNumArgs);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-connecttoconnectionpoint
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT ConnectToConnectionPoint(nint punk, in Guid riidEvent, BOOL fConnect, nint punkTarget, out uint pdwCookie, nint /* optional IConnectionPoint* */ ppcpOut);
    
    // https://learn.microsoft.com/windows/win32/api/wingdi/nf-wingdi-createcompatibledc
    [LibraryImport("GDI32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HDC CreateCompatibleDC(HDC hdc);
    
    // https://learn.microsoft.com/windows/win32/api/fileapi/nf-fileapi-createfilew
    [LibraryImport("KERNEL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HANDLE CreateFileW(PWSTR lpFileName, uint dwDesiredAccess, FILE_SHARE_MODE dwShareMode, nint /* optional SECURITY_ATTRIBUTES* */ lpSecurityAttributes, FILE_CREATION_DISPOSITION dwCreationDisposition, FILE_FLAGS_AND_ATTRIBUTES dwFlagsAndAttributes, HANDLE hTemplateFile);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-createmenu
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HMENU CreateMenu();
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-createpopupmenu
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HMENU CreatePopupMenu();
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-createprofile
    [LibraryImport("USERENV")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT CreateProfile(PWSTR pszUserSid, PWSTR pszUserName, [MarshalUsing(CountElementName = nameof(cchProfilePath))] PWSTR pszProfilePath, uint cchProfilePath);
    
    // https://learn.microsoft.com/windows/win32/api/wincrypt/nf-wincrypt-cryptdecodeobjectex
    [LibraryImport("CRYPT32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL CryptDecodeObjectEx(CERT_QUERY_ENCODING_TYPE dwCertEncodingType, PSTR lpszStructType, nint /* byte array */ pbEncoded, uint cbEncoded, uint dwFlags, nint /* optional CRYPT_DECODE_PARA* */ pDecodePara, nint /* optional void* */ pvStructInfo, ref uint pcbStructInfo);
    
    // https://learn.microsoft.com/windows/win32/api/wincrypt/nf-wincrypt-cryptmsgclose
    [LibraryImport("CRYPT32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL CryptMsgClose(nint /* optional void* */ hCryptMsg);
    
    // https://learn.microsoft.com/windows/win32/api/wincrypt/nf-wincrypt-cryptmsggetparam
    [LibraryImport("CRYPT32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL CryptMsgGetParam(nint hCryptMsg, uint dwParamType, uint dwIndex, nint /* optional void* */ pvData, ref uint pcbData);
    
    // https://learn.microsoft.com/windows/win32/api/wincrypt/nf-wincrypt-cryptmsgopentodecode
    [LibraryImport("CRYPT32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint CryptMsgOpenToDecode(uint dwMsgEncodingType, uint dwFlags, uint dwMsgType, HCRYPTPROV_LEGACY hCryptProv, nint /* optional CERT_INFO* */ pRecipientInfo, nint /* optional CMSG_STREAM_INFO* */ pStreamInfo);
    
    // https://learn.microsoft.com/windows/win32/api/wincrypt/nf-wincrypt-cryptmsgupdate
    [LibraryImport("CRYPT32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL CryptMsgUpdate(nint hCryptMsg, nint /* optional byte* */ pbData, uint cbData, BOOL fFinal);
    
    // https://learn.microsoft.com/windows/win32/api/wincrypt/nf-wincrypt-cryptqueryobject
    [LibraryImport("CRYPT32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL CryptQueryObject(CERT_QUERY_OBJECT_TYPE dwObjectType, nint pvObject, CERT_QUERY_CONTENT_TYPE_FLAGS dwExpectedContentTypeFlags, CERT_QUERY_FORMAT_TYPE_FLAGS dwExpectedFormatTypeFlags, uint dwFlags, nint /* optional CERT_QUERY_ENCODING_TYPE* */ pdwMsgAndCertEncodingType, nint /* optional CERT_QUERY_CONTENT_TYPE* */ pdwContentType, nint /* optional CERT_QUERY_FORMAT_TYPE* */ pdwFormatType, nint /* optional HCERTSTORE* */ phCertStore, nint /* optional void** */ phMsg, nint /* optional void** */ ppvContext);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-dad_autoscroll
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DAD_AutoScroll(HWND hwnd, ref AUTO_SCROLL_DATA pad, in POINT pptNow);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-dad_dragenterex
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DAD_DragEnterEx(HWND hwndTarget, POINT ptStart);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-dad_dragenterex2
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DAD_DragEnterEx2(HWND hwndTarget, POINT ptStart, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject?>))] IDataObject? pdtObject);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-dad_dragleave
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DAD_DragLeave();
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-dad_dragmove
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DAD_DragMove(POINT pt);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-dad_setdragimage
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DAD_SetDragImage(HIMAGELIST him, ref POINT pptOffset);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-dad_showdragimage
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DAD_ShowDragImage(BOOL fShow);
    
    // https://learn.microsoft.com/windows/win32/api/commctrl/nf-commctrl-defsubclassproc
    [LibraryImport("COMCTL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial LRESULT DefSubclassProc(HWND hWnd, uint uMsg, WPARAM wParam, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-deleteprofilea
    [LibraryImport("USERENV", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DeleteProfileA(PSTR lpSidString, PSTR lpProfilePath, PSTR lpComputerName);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-deleteprofilew
    [LibraryImport("USERENV", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DeleteProfileW(PWSTR lpSidString, PWSTR lpProfilePath, PWSTR lpComputerName);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-destroymenu
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DestroyMenu(HMENU hMenu);
    
    [LibraryImport("DMProcessXMLFiltered")]
    [PreserveSig]
    public static partial HRESULT DMProcessConfigXMLFiltered(PWSTR pszXmlIn, [In][MarshalUsing(CountElementName = nameof(dwNumAllowedCspNodes))] PWSTR[] rgszAllowedCspNodes, uint dwNumAllowedCspNodes, out BSTR pbstrXmlOut);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-doenvironmentsubsta
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint DoEnvironmentSubstA([MarshalUsing(CountElementName = nameof(cchSrc))] PSTR pszSrc, uint cchSrc);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-doenvironmentsubstw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint DoEnvironmentSubstW([MarshalUsing(CountElementName = nameof(cchSrc))] PWSTR pszSrc, uint cchSrc);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-dragacceptfiles
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void DragAcceptFiles(HWND hWnd, BOOL fAccept);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-dragfinish
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void DragFinish(HDROP hDrop);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-dragqueryfilea
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint DragQueryFileA(HDROP hDrop, uint iFile, [MarshalUsing(CountElementName = nameof(cch))] PSTR lpszFile, uint cch);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-dragqueryfilew
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint DragQueryFileW(HDROP hDrop, uint iFile, [MarshalUsing(CountElementName = nameof(cch))] PWSTR lpszFile, uint cch);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-dragquerypoint
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL DragQueryPoint(HDROP hDrop, out POINT ppt);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-drivetype
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int DriveType(int iDrive);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-duplicateicon
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HICON DuplicateIcon(HINSTANCE hInst, HICON hIcon);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-extractassociatedicona
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HICON ExtractAssociatedIconA(HINSTANCE hInst, [MarshalUsing(ConstantElementCount = 128)] PSTR pszIconPath, ref ushort piIcon);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-extractassociatediconexa
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HICON ExtractAssociatedIconExA(HINSTANCE hInst, [MarshalUsing(ConstantElementCount = 128)] PSTR pszIconPath, ref ushort piIconIndex, ref ushort piIconId);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-extractassociatediconexw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HICON ExtractAssociatedIconExW(HINSTANCE hInst, [MarshalUsing(ConstantElementCount = 128)] PWSTR pszIconPath, ref ushort piIconIndex, ref ushort piIconId);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-extractassociatediconw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HICON ExtractAssociatedIconW(HINSTANCE hInst, [MarshalUsing(ConstantElementCount = 128)] PWSTR pszIconPath, ref ushort piIcon);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-extracticona
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HICON ExtractIconA(HINSTANCE hInst, PSTR pszExeFileName, uint nIconIndex);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-extracticonexa
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint ExtractIconExA(PSTR lpszFile, int nIconIndex, nint /* optional HICON* */ phiconLarge, nint /* optional HICON* */ phiconSmall, uint nIcons);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-extracticonexw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint ExtractIconExW(PWSTR lpszFile, int nIconIndex, nint /* optional HICON* */ phiconLarge, nint /* optional HICON* */ phiconSmall, uint nIcons);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-extracticonw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HICON ExtractIconW(HINSTANCE hInst, PWSTR pszExeFileName, uint nIconIndex);
    
    // https://learn.microsoft.com/windows/win32/shell/fileiconinit
    [LibraryImport("SHELL32", SetLastError = true)]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL FileIconInit(BOOL fRestoreCache);
    
    // https://learn.microsoft.com/windows/win32/api/fileapi/nf-fileapi-filetimetolocalfiletime
    [LibraryImport("KERNEL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL FileTimeToLocalFileTime(in FILETIME lpFileTime, out FILETIME lpLocalFileTime);
    
    // https://learn.microsoft.com/windows/win32/api/timezoneapi/nf-timezoneapi-filetimetosystemtime
    [LibraryImport("KERNEL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL FileTimeToSystemTime(in FILETIME lpFileTime, out SYSTEMTIME lpSystemTime);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-findexecutablea
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HINSTANCE FindExecutableA(PSTR lpFile, PSTR lpDirectory, [MarshalUsing(ConstantElementCount = 260)] PSTR lpResult);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-findexecutablew
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HINSTANCE FindExecutableW(PWSTR lpFile, PWSTR lpDirectory, [MarshalUsing(ConstantElementCount = 260)] PWSTR lpResult);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-getacceptlanguagesa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT GetAcceptLanguagesA([MarshalUsing(CountElementName = nameof(pcchLanguages))] PSTR pszLanguages, ref uint pcchLanguages);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-getacceptlanguagesw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT GetAcceptLanguagesW([MarshalUsing(CountElementName = nameof(pcchLanguages))] PWSTR pszLanguages, ref uint pcchLanguages);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-getallusersprofiledirectorya
    [LibraryImport("USERENV", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetAllUsersProfileDirectoryA([MarshalUsing(CountElementName = nameof(lpcchSize))] PSTR lpProfileDir, ref uint lpcchSize);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-getallusersprofiledirectoryw
    [LibraryImport("USERENV", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetAllUsersProfileDirectoryW([MarshalUsing(CountElementName = nameof(lpcchSize))] PWSTR lpProfileDir, ref uint lpcchSize);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-getcurrentprocessexplicitappusermodelid
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT GetCurrentProcessExplicitAppUserModelID(out PWSTR AppID);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-getdefaultuserprofiledirectorya
    [LibraryImport("USERENV", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetDefaultUserProfileDirectoryA([MarshalUsing(CountElementName = nameof(lpcchSize))] PSTR lpProfileDir, ref uint lpcchSize);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-getdefaultuserprofiledirectoryw
    [LibraryImport("USERENV", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetDefaultUserProfileDirectoryW([MarshalUsing(CountElementName = nameof(lpcchSize))] PWSTR lpProfileDir, ref uint lpcchSize);
    
    // https://learn.microsoft.com/windows/win32/api/wingdi/nf-wingdi-getdibits
    [LibraryImport("GDI32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int GetDIBits(HDC hdc, HBITMAP hbm, uint start, uint cLines, nint /* optional void* */ lpvBits, ref BITMAPINFO lpbmi, DIB_USAGE usage);
    
    // https://learn.microsoft.com/windows/win32/api/shellscalingapi/nf-shellscalingapi-getdpiforshelluicomponent
    [LibraryImport("api-ms-win-shcore-scaling-l1-1-2")]
    [SupportedOSPlatform("windows8.1")]
    [PreserveSig]
    public static partial uint GetDpiForShellUIComponent(SHELL_UI_COMPONENT param0);
    
    // https://learn.microsoft.com/windows/win32/api/fileapi/nf-fileapi-getdrivetypew
    [LibraryImport("KERNEL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint GetDriveTypeW(PWSTR lpRootPathName);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-getfilenamefrombrowse
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetFileNameFromBrowse(HWND hwnd, [MarshalUsing(CountElementName = nameof(cchFilePath))] PWSTR pszFilePath, uint cchFilePath, PWSTR pszWorkingDir, PWSTR pszDefExt, PWSTR pszFilters, PWSTR pszTitle);
    
    // https://learn.microsoft.com/windows/win32/api/fileapi/nf-fileapi-getfinalpathnamebyhandlew
    [LibraryImport("KERNEL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial uint GetFinalPathNameByHandleW(HANDLE hFile, [MarshalUsing(CountElementName = nameof(cchFilePath))] PWSTR lpszFilePath, uint cchFilePath, GETFINALPATHNAMEBYHANDLE_FLAGS dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-getmenucontexthelpid
    [LibraryImport("USER32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint GetMenuContextHelpId(HMENU param0);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-getmenuitemcount
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int GetMenuItemCount(HMENU hMenu);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-getmenuitemid
    [LibraryImport("USER32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial uint GetMenuItemID(HMENU hMenu, int nPos);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-getmenuiteminfow
    [LibraryImport("USER32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetMenuItemInfoW(HMENU hmenu, uint item, BOOL fByPosition, ref MENUITEMINFOW lpmii);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-getmenuposfromid
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int GetMenuPosFromID(HMENU hmenu, uint id);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-getprofilesdirectorya
    [LibraryImport("USERENV", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetProfilesDirectoryA([MarshalUsing(CountElementName = nameof(lpcchSize))] PSTR lpProfileDir, ref uint lpcchSize);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-getprofilesdirectoryw
    [LibraryImport("USERENV", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetProfilesDirectoryW([MarshalUsing(CountElementName = nameof(lpcchSize))] PWSTR lpProfileDir, ref uint lpcchSize);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-getprofiletype
    [LibraryImport("USERENV", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetProfileType(out uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shellscalingapi/nf-shellscalingapi-getscalefactorfordevice
    [LibraryImport("api-ms-win-shcore-scaling-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial DEVICE_SCALE_FACTOR GetScaleFactorForDevice(DISPLAY_DEVICE_TYPE deviceType);
    
    // https://learn.microsoft.com/windows/win32/api/shellscalingapi/nf-shellscalingapi-getscalefactorformonitor
    [LibraryImport("api-ms-win-shcore-scaling-l1-1-1")]
    [SupportedOSPlatform("windows8.1")]
    [PreserveSig]
    public static partial HRESULT GetScaleFactorForMonitor(HMONITOR hMon, out DEVICE_SCALE_FACTOR pScale);
    
    // https://learn.microsoft.com/windows/win32/api/virtdisk/nf-virtdisk-getstoragedependencyinformation
    [LibraryImport("VirtDisk")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial WIN32_ERROR GetStorageDependencyInformation(HANDLE ObjectHandle, GET_STORAGE_DEPENDENCY_FLAG Flags, uint StorageDependencyInfoSize, ref STORAGE_DEPENDENCY_INFO StorageDependencyInfo, nint /* optional uint* */ SizeUsed);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-getuserprofiledirectorya
    [LibraryImport("USERENV", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetUserProfileDirectoryA(HANDLE hToken, [MarshalUsing(CountElementName = nameof(lpcchSize))] PSTR lpProfileDir, ref uint lpcchSize);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-getuserprofiledirectoryw
    [LibraryImport("USERENV", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetUserProfileDirectoryW(HANDLE hToken, [MarshalUsing(CountElementName = nameof(lpcchSize))] PWSTR lpProfileDir, ref uint lpcchSize);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-getwindowcontexthelpid
    [LibraryImport("USER32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint GetWindowContextHelpId(HWND param0);
    
    // https://learn.microsoft.com/windows/win32/api/commctrl/nf-commctrl-getwindowsubclass
    [LibraryImport("COMCTL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL GetWindowSubclass(HWND hWnd, SUBCLASSPROC pfnSubclass, nuint uIdSubclass, nint /* optional nuint* */ pdwRefData);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-hashdata
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT HashData(nint /* byte array */ pbData, uint cbData, nint /* byte array */ pbHash, uint cbHash);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkClone([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlink>))] IHlink pihl, in Guid riid, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkSite>))] IHlinkSite pihlsiteForClone, uint dwSiteData, out nint /* void */ ppvObj);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkCreateBrowseContext(nint piunkOuter, in Guid riid, out nint /* void */ ppvObj);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkCreateExtensionServices(PWSTR pwzAdditionalHeaders, HWND phwnd, PWSTR pszUsername, PWSTR pszPassword, nint piunkOuter, in Guid riid, out nint /* void */ ppvObj);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkCreateFromData([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject>))] IDataObject piDataObj, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkSite>))] IHlinkSite pihlsite, uint dwSiteData, nint piunkOuter, in Guid riid, out nint /* void */ ppvObj);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkCreateFromMoniker([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] IMoniker pimkTrgt, PWSTR pwzLocation, PWSTR pwzFriendlyName, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkSite>))] IHlinkSite pihlsite, uint dwSiteData, nint piunkOuter, in Guid riid, out nint /* void */ ppvObj);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkCreateFromString(PWSTR pwzTarget, PWSTR pwzLocation, PWSTR pwzFriendlyName, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkSite>))] IHlinkSite pihlsite, uint dwSiteData, nint piunkOuter, in Guid riid, out nint /* void */ ppvObj);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkCreateShortcut(uint grfHLSHORTCUTF, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlink>))] IHlink pihl, PWSTR pwzDir, PWSTR pwzFileName, out PWSTR ppwzShortcutFile, uint dwReserved);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkCreateShortcutFromMoniker(uint grfHLSHORTCUTF, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] IMoniker pimkTarget, PWSTR pwzLocation, PWSTR pwzDir, PWSTR pwzFileName, out PWSTR ppwzShortcutFile, uint dwReserved);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkCreateShortcutFromString(uint grfHLSHORTCUTF, PWSTR pwzTarget, PWSTR pwzLocation, PWSTR pwzDir, PWSTR pwzFileName, out PWSTR ppwzShortcutFile, uint dwReserved);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkGetSpecialReference(uint uReference, out PWSTR ppwzReference);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkGetValueFromParams(PWSTR pwzParams, PWSTR pwzName, out PWSTR ppwzValue);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkIsShortcut(PWSTR pwzFileName);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkNavigate([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlink>))] IHlink pihl, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkFrame>))] IHlinkFrame pihlframe, uint grfHLNF, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx>))] IBindCtx pbc, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindStatusCallback>))] IBindStatusCallback pibsc, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkBrowseContext>))] IHlinkBrowseContext pihlbc);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkNavigateToStringReference(PWSTR pwzTarget, PWSTR pwzLocation, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkSite>))] IHlinkSite pihlsite, uint dwSiteData, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkFrame>))] IHlinkFrame pihlframe, uint grfHLNF, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx>))] IBindCtx pibc, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindStatusCallback>))] IBindStatusCallback pibsc, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkBrowseContext>))] IHlinkBrowseContext pihlbc);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkOnNavigate([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkFrame>))] IHlinkFrame pihlframe, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkBrowseContext>))] IHlinkBrowseContext pihlbc, uint grfHLNF, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] IMoniker pimkTarget, PWSTR pwzLocation, PWSTR pwzFriendlyName, ref uint puHLID);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkOnRenameDocument(uint dwReserved, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkBrowseContext>))] IHlinkBrowseContext pihlbc, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] IMoniker pimkOld, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] IMoniker pimkNew);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkParseDisplayName([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx>))] IBindCtx pibc, PWSTR pwzDisplayName, BOOL fNoForceAbs, ref uint pcchEaten, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] out IMoniker ppimk);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkPreprocessMoniker([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx>))] IBindCtx pibc, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] IMoniker pimkIn, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] out IMoniker ppimkOut);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkQueryCreateFromData([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject>))] IDataObject piDataObj);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkResolveMonikerForData([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] IMoniker pimkReference, uint reserved, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx>))] IBindCtx pibc, uint cFmtetc, ref FORMATETC rgFmtetc, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindStatusCallback>))] IBindStatusCallback pibsc, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] IMoniker pimkBase);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkResolveShortcut(PWSTR pwzShortcutFileName, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkSite>))] IHlinkSite pihlsite, uint dwSiteData, nint piunkOuter, in Guid riid, out nint ppvObj);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkResolveShortcutToMoniker(PWSTR pwzShortcutFileName, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] out IMoniker ppimkTarget, out PWSTR ppwzLocation);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkResolveShortcutToString(PWSTR pwzShortcutFileName, out PWSTR ppwzTarget, out PWSTR ppwzLocation);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkResolveStringForData(PWSTR pwzReference, uint reserved, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx>))] IBindCtx pibc, uint cFmtetc, ref FORMATETC rgFmtetc, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindStatusCallback>))] IBindStatusCallback pibsc, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] IMoniker pimkBase);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkSetSpecialReference(uint uReference, PWSTR pwzReference);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkTranslateURL(PWSTR pwzURL, uint grfFlags, out PWSTR ppwzTranslatedURL);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT HlinkUpdateStackItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkFrame>))] IHlinkFrame pihlframe, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IHlinkBrowseContext>))] IHlinkBrowseContext pihlbc, uint uHLID, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] IMoniker pimkTrgt, PWSTR pwzLocation, PWSTR pwzFriendlyName);
    
    [LibraryImport("OLE32")]
    [PreserveSig]
    public static partial void HMONITOR_UserFree(in uint param0, in HMONITOR param1);
    
    [LibraryImport("OLE32")]
    [PreserveSig]
    public static partial void HMONITOR_UserFree64(in uint param0, in HMONITOR param1);
    
    [LibraryImport("OLE32")]
    [PreserveSig]
    public static partial nint HMONITOR_UserMarshal(in uint param0, nint /* byte array */ param1, in HMONITOR param2);
    
    [LibraryImport("OLE32")]
    [PreserveSig]
    public static partial nint HMONITOR_UserMarshal64(in uint param0, nint /* byte array */ param1, in HMONITOR param2);
    
    [LibraryImport("OLE32")]
    [PreserveSig]
    public static partial uint HMONITOR_UserSize(in uint param0, uint param1, in HMONITOR param2);
    
    [LibraryImport("OLE32")]
    [PreserveSig]
    public static partial uint HMONITOR_UserSize64(in uint param0, uint param1, in HMONITOR param2);
    
    [LibraryImport("OLE32")]
    [PreserveSig]
    public static partial nint HMONITOR_UserUnmarshal(in uint param0, nint /* byte array */ param1, out HMONITOR param2);
    
    [LibraryImport("OLE32")]
    [PreserveSig]
    public static partial nint HMONITOR_UserUnmarshal64(in uint param0, nint /* byte array */ param1, out HMONITOR param2);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilappendid
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint ILAppendID(nint /* optional nint* */ pidl, in SHITEMID pmkid, BOOL fAppend);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilclone
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint ILClone(nint pidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilclonefirst
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint ILCloneFirst(nint pidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilcombine
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint ILCombine(nint /* optional nint* */ pidl1, nint /* optional nint* */ pidl2);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilcreatefrompatha
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint ILCreateFromPathA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilcreatefrompathw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint ILCreateFromPathW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilfindchild
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint ILFindChild(nint pidlParent, nint pidlChild);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilfindlastid
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint ILFindLastID(nint pidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilfree
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void ILFree(nint /* optional nint* */ pidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilgetnext
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint ILGetNext(nint /* optional nint* */ pidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilgetsize
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint ILGetSize(nint /* optional nint* */ pidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilisequal
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ILIsEqual(nint pidl1, nint pidl2);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilisparent
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ILIsParent(nint pidl1, nint pidl2, BOOL fImmediate);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-illoadfromstreamex
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT ILLoadFromStreamEx([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm, out nint pidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilremovelastid
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ILRemoveLastID(nint /* optional nint* */ pidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-ilsavetostream
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT ILSaveToStream([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm, nint pidl);
    
    [LibraryImport("SHDOCVW", SetLastError = true)]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ImportPrivacySettings(PWSTR pszFilename, ref BOOL pfParsePrivacyPreferences, ref BOOL pfParsePerSiteRules);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-initnetworkaddresscontrol
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL InitNetworkAddressControl();
    
    // https://learn.microsoft.com/windows/win32/api/propvarutil/nf-propvarutil-initpropvariantfromstrret
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT InitPropVariantFromStrRet(ref STRRET pstrret, nint /* optional nint* */ pidl, out PROPVARIANT ppropvar);
    
    // https://learn.microsoft.com/windows/win32/api/propvarutil/nf-propvarutil-initvariantfromstrret
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT InitVariantFromStrRet(in STRRET pstrret, nint pidl, out VARIANT pvar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-intlstreqworkera
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL IntlStrEqWorkerA(BOOL fCaseSens, [MarshalUsing(CountElementName = nameof(nChar))] PSTR lpString1, [MarshalUsing(CountElementName = nameof(nChar))] PSTR lpString2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-intlstreqworkerw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL IntlStrEqWorkerW(BOOL fCaseSens, [MarshalUsing(CountElementName = nameof(nChar))] PWSTR lpString1, [MarshalUsing(CountElementName = nameof(nChar))] PWSTR lpString2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-ischarspacea
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL IsCharSpaceA(CHAR wch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-ischarspacew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL IsCharSpaceW(char wch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-isinternetescenabled
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL IsInternetESCEnabled();
    
    [LibraryImport("SHELL32", SetLastError = true)]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL IsLFNDriveA(PSTR pszPath);
    
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL IsLFNDriveW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-isnetdrive
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int IsNetDrive(int iDrive);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-isos
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL IsOS(OS dwOS);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-istream_copy
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT IStream_Copy([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstmFrom, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstmTo, uint cb);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-istream_read
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT IStream_Read([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm, nint pv, uint cb);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-istream_readpidl
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT IStream_ReadPidl([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm, out nint ppidlOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-istream_readstr
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT IStream_ReadStr([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm, out PWSTR ppsz);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-istream_reset
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT IStream_Reset([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-istream_size
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT IStream_Size([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm, out ulong pui);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-istream_write
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT IStream_Write([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm, nint pv, uint cb);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-istream_writepidl
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT IStream_WritePidl([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm, nint pidlWrite);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-istream_writestr
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT IStream_WriteStr([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pstm, PWSTR psz);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-isuseranadmin
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL IsUserAnAdmin();
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iunknown_atomicrelease
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void IUnknown_AtomicRelease(nint /* optional void** */ ppunk);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iunknown_getsite
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT IUnknown_GetSite(nint punk, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iunknown_getwindow
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT IUnknown_GetWindow(nint punk, out HWND phwnd);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iunknown_queryservice
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT IUnknown_QueryService(nint punk, in Guid guidService, in Guid riid, out nint /* void */ ppvOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iunknown_set
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void IUnknown_Set(ref nint ppunk, nint punk);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-iunknown_setsite
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT IUnknown_SetSite(nint punk, nint punkSite);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-loaduserprofilea
    [LibraryImport("USERENV", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL LoadUserProfileA(HANDLE hToken, ref PROFILEINFOA lpProfileInfo);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-loaduserprofilew
    [LibraryImport("USERENV", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL LoadUserProfileW(HANDLE hToken, ref PROFILEINFOW lpProfileInfo);
    
    [LibraryImport("hlink")]
    [PreserveSig]
    public static partial HRESULT OleSaveToStreamEx(nint piunk, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream pistm, BOOL fClearDirty);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-openregstream
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [return: MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))]
    public static partial IStream OpenRegStream(HKEY hkey, PWSTR pszSubkey, PWSTR pszValue, uint grfMode);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-parseurla
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT ParseURLA(PSTR pcszURL, ref PARSEDURLA ppu);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-parseurlw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT ParseURLW(PWSTR pcszURL, ref PARSEDURLW ppu);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathaddbackslasha
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR PathAddBackslashA([MarshalUsing(ConstantElementCount = 260)] PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathaddbackslashw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR PathAddBackslashW([MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathaddextensiona
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathAddExtensionA([MarshalUsing(ConstantElementCount = 260)] PSTR pszPath, PSTR pszExt);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathaddextensionw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathAddExtensionW([MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath, PWSTR pszExt);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathalloccanonicalize
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathAllocCanonicalize(PWSTR pszPathIn, PATHCCH_OPTIONS dwFlags, out PWSTR ppszPathOut);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathalloccombine
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathAllocCombine(PWSTR pszPathIn, PWSTR pszMore, PATHCCH_OPTIONS dwFlags, out PWSTR ppszPathOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathappenda
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathAppendA([MarshalUsing(ConstantElementCount = 260)] PSTR pszPath, PSTR pszMore);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathappendw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathAppendW([MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath, PWSTR pszMore);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathbuildroota
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR PathBuildRootA([MarshalUsing(ConstantElementCount = 4)] PSTR pszRoot, int iDrive);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathbuildrootw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR PathBuildRootW([MarshalUsing(ConstantElementCount = 4)] PWSTR pszRoot, int iDrive);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcanonicalizea
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathCanonicalizeA([MarshalUsing(ConstantElementCount = 260)] PSTR pszBuf, PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcanonicalizew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathCanonicalizeW([MarshalUsing(ConstantElementCount = 260)] PWSTR pszBuf, PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchaddbackslash
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchAddBackslash([MarshalUsing(CountElementName = nameof(cchPath))] PWSTR pszPath, nuint cchPath);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchaddbackslashex
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchAddBackslashEx([MarshalUsing(CountElementName = nameof(cchPath))] PWSTR pszPath, nuint cchPath, nint /* optional PWSTR* */ ppszEnd, nint /* optional nuint* */ pcchRemaining);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchaddextension
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchAddExtension([MarshalUsing(CountElementName = nameof(cchPath))] PWSTR pszPath, nuint cchPath, PWSTR pszExt);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchappend
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchAppend([MarshalUsing(CountElementName = nameof(cchPath))] PWSTR pszPath, nuint cchPath, PWSTR pszMore);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchappendex
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchAppendEx([MarshalUsing(CountElementName = nameof(cchPath))] PWSTR pszPath, nuint cchPath, PWSTR pszMore, PATHCCH_OPTIONS dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchcanonicalize
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchCanonicalize([MarshalUsing(CountElementName = nameof(cchPathOut))] PWSTR pszPathOut, nuint cchPathOut, PWSTR pszPathIn);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchcanonicalizeex
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchCanonicalizeEx([MarshalUsing(CountElementName = nameof(cchPathOut))] PWSTR pszPathOut, nuint cchPathOut, PWSTR pszPathIn, PATHCCH_OPTIONS dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchcombine
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchCombine([MarshalUsing(CountElementName = nameof(cchPathOut))] PWSTR pszPathOut, nuint cchPathOut, PWSTR pszPathIn, PWSTR pszMore);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchcombineex
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchCombineEx([MarshalUsing(CountElementName = nameof(cchPathOut))] PWSTR pszPathOut, nuint cchPathOut, PWSTR pszPathIn, PWSTR pszMore, PATHCCH_OPTIONS dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchfindextension
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchFindExtension(PWSTR pszPath, nuint cchPath, out PWSTR ppszExt);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchisroot
    [LibraryImport("api-ms-win-core-path-l1-1-0", SetLastError = true)]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathCchIsRoot(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchremovebackslash
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchRemoveBackslash([MarshalUsing(CountElementName = nameof(cchPath))] PWSTR pszPath, nuint cchPath);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchremovebackslashex
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchRemoveBackslashEx(PWSTR pszPath, nuint cchPath, nint /* optional PWSTR* */ ppszEnd, nint /* optional nuint* */ pcchRemaining);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchremoveextension
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchRemoveExtension(PWSTR pszPath, nuint cchPath);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchremovefilespec
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchRemoveFileSpec(PWSTR pszPath, nuint cchPath);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchrenameextension
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchRenameExtension([MarshalUsing(CountElementName = nameof(cchPath))] PWSTR pszPath, nuint cchPath, PWSTR pszExt);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchskiproot
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchSkipRoot(PWSTR pszPath, out PWSTR ppszRootEnd);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchstripprefix
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchStripPrefix([MarshalUsing(CountElementName = nameof(cchPath))] PWSTR pszPath, nuint cchPath);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathcchstriptoroot
    [LibraryImport("api-ms-win-core-path-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT PathCchStripToRoot(PWSTR pszPath, nuint cchPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pathcleanupspec
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int PathCleanupSpec(PWSTR pszDir, PWSTR pszSpec);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcombinea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR PathCombineA([MarshalUsing(ConstantElementCount = 260)] PSTR pszDest, PSTR pszDir, PSTR pszFile);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcombinew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR PathCombineW([MarshalUsing(ConstantElementCount = 260)] PWSTR pszDest, PWSTR pszDir, PWSTR pszFile);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcommonprefixa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int PathCommonPrefixA(PSTR pszFile1, PSTR pszFile2, [MarshalUsing(ConstantElementCount = 260)] PSTR achPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcommonprefixw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int PathCommonPrefixW(PWSTR pszFile1, PWSTR pszFile2, [MarshalUsing(ConstantElementCount = 260)] PWSTR achPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcompactpatha
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathCompactPathA(HDC hDC, [MarshalUsing(ConstantElementCount = 260)] PSTR pszPath, uint dx);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcompactpathexa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathCompactPathExA([MarshalUsing(CountElementName = nameof(cchMax))] PSTR pszOut, PSTR pszSrc, uint cchMax, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcompactpathexw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathCompactPathExW([MarshalUsing(CountElementName = nameof(cchMax))] PWSTR pszOut, PWSTR pszSrc, uint cchMax, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcompactpathw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathCompactPathW(HDC hDC, [MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath, uint dx);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcreatefromurla
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT PathCreateFromUrlA(PSTR pszUrl, [MarshalUsing(CountElementName = nameof(pcchPath))] PSTR pszPath, ref uint pcchPath, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcreatefromurlalloc
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT PathCreateFromUrlAlloc(PWSTR pszIn, out PWSTR ppszOut, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathcreatefromurlw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT PathCreateFromUrlW(PWSTR pszUrl, [MarshalUsing(CountElementName = nameof(pcchPath))] PWSTR pszPath, ref uint pcchPath, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfileexistsa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathFileExistsA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfileexistsw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathFileExistsW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfindextensiona
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR PathFindExtensionA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfindextensionw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR PathFindExtensionW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfindfilenamea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR PathFindFileNameA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfindfilenamew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR PathFindFileNameW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfindnextcomponenta
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR PathFindNextComponentA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfindnextcomponentw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR PathFindNextComponentW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfindonpatha
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathFindOnPathA([MarshalUsing(ConstantElementCount = 260)] PSTR pszPath, nint /* optional sbyte** */ ppszOtherDirs);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfindonpathw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathFindOnPathW([MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath, nint /* optional ushort** */ ppszOtherDirs);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfindsuffixarraya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR PathFindSuffixArrayA(PSTR pszPath, [In][MarshalUsing(CountElementName = nameof(iArraySize))] PSTR[] apszSuffix, int iArraySize);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathfindsuffixarrayw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR PathFindSuffixArrayW(PWSTR pszPath, [In][MarshalUsing(CountElementName = nameof(iArraySize))] PWSTR[] apszSuffix, int iArraySize);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathgetargsa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR PathGetArgsA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathgetargsw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR PathGetArgsW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathgetchartypea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial uint PathGetCharTypeA(byte ch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathgetchartypew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial uint PathGetCharTypeW(char ch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathgetdrivenumbera
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int PathGetDriveNumberA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathgetdrivenumberw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int PathGetDriveNumberW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pathgetshortpath
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void PathGetShortPath([MarshalUsing(ConstantElementCount = 260)] PWSTR pszLongPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathiscontenttypea
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsContentTypeA(PSTR pszPath, PSTR pszContentType);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathiscontenttypew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsContentTypeW(PWSTR pszPath, PWSTR pszContentType);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisdirectorya
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsDirectoryA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisdirectoryemptya
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsDirectoryEmptyA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisdirectoryemptyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsDirectoryEmptyW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisdirectoryw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsDirectoryW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pathisexe
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsExe(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisfilespeca
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsFileSpecA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisfilespecw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsFileSpecW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathislfnfilespeca
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsLFNFileSpecA(PSTR pszName);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathislfnfilespecw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsLFNFileSpecW(PWSTR pszName);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisnetworkpatha
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsNetworkPathA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisnetworkpathw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsNetworkPathW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisprefixa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsPrefixA(PSTR pszPrefix, PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisprefixw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsPrefixW(PWSTR pszPrefix, PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisrelativea
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsRelativeA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisrelativew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsRelativeW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisroota
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsRootA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisrootw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsRootW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathissameroota
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsSameRootA(PSTR pszPath1, PSTR pszPath2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathissamerootw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsSameRootW(PWSTR pszPath1, PWSTR pszPath2);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-pathisslowa
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsSlowA(PSTR pszFile, uint dwAttr);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-pathissloww
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsSlowW(PWSTR pszFile, uint dwAttr);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathissystemfoldera
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsSystemFolderA(PSTR pszPath, uint dwAttrb);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathissystemfolderw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsSystemFolderW(PWSTR pszPath, uint dwAttrb);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisunca
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsUNCA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/pathcch/nf-pathcch-pathisuncex
    [LibraryImport("api-ms-win-core-path-l1-1-0", SetLastError = true)]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsUNCEx(PWSTR pszPath, nint /* optional PWSTR* */ ppszServer);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisuncservera
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsUNCServerA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisuncserversharea
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsUNCServerShareA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisuncserversharew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsUNCServerShareW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisuncserverw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsUNCServerW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisuncw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsUNCW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisurla
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsURLA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathisurlw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathIsURLW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathmakeprettya
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathMakePrettyA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathmakeprettyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathMakePrettyW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathmakesystemfoldera
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathMakeSystemFolderA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathmakesystemfolderw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathMakeSystemFolderW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pathmakeuniquename
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathMakeUniqueName([MarshalUsing(CountElementName = nameof(cchMax))] PWSTR pszUniqueName, uint cchMax, PWSTR pszTemplate, PWSTR pszLongPlate, PWSTR pszDir);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathmatchspeca
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathMatchSpecA(PSTR pszFile, PSTR pszSpec);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathmatchspecexa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT PathMatchSpecExA(PSTR pszFile, PSTR pszSpec, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathmatchspecexw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT PathMatchSpecExW(PWSTR pszFile, PWSTR pszSpec, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathmatchspecw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathMatchSpecW(PWSTR pszFile, PWSTR pszSpec);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathparseiconlocationa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int PathParseIconLocationA(PSTR pszIconFile);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathparseiconlocationw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int PathParseIconLocationW(PWSTR pszIconFile);
    
    [LibraryImport("SHELL32")]
    [PreserveSig]
    public static partial void PathQualify(PWSTR psz);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathquotespacesa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathQuoteSpacesA([MarshalUsing(ConstantElementCount = 260)] PSTR lpsz);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathquotespacesw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathQuoteSpacesW([MarshalUsing(ConstantElementCount = 260)] PWSTR lpsz);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathrelativepathtoa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathRelativePathToA([MarshalUsing(ConstantElementCount = 260)] PSTR pszPath, PSTR pszFrom, uint dwAttrFrom, PSTR pszTo, uint dwAttrTo);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathrelativepathtow
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathRelativePathToW([MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath, PWSTR pszFrom, uint dwAttrFrom, PWSTR pszTo, uint dwAttrTo);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathremoveargsa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathRemoveArgsA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathremoveargsw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathRemoveArgsW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathremovebackslasha
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR PathRemoveBackslashA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathremovebackslashw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR PathRemoveBackslashW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathremoveblanksa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathRemoveBlanksA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathremoveblanksw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathRemoveBlanksW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathremoveextensiona
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathRemoveExtensionA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathremoveextensionw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathRemoveExtensionW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathremovefilespeca
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathRemoveFileSpecA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathremovefilespecw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathRemoveFileSpecW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathrenameextensiona
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathRenameExtensionA([MarshalUsing(ConstantElementCount = 260)] PSTR pszPath, PSTR pszExt);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathrenameextensionw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathRenameExtensionW([MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath, PWSTR pszExt);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pathresolve
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int PathResolve([MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath, nint /* optional ushort** */ dirs, uint fFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathsearchandqualifya
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathSearchAndQualifyA(PSTR pszPath, [MarshalUsing(CountElementName = nameof(cchBuf))] PSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathsearchandqualifyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathSearchAndQualifyW(PWSTR pszPath, [MarshalUsing(CountElementName = nameof(cchBuf))] PWSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathsetdlgitempatha
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathSetDlgItemPathA(HWND hDlg, int id, PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathsetdlgitempathw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathSetDlgItemPathW(HWND hDlg, int id, PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathskiproota
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR PathSkipRootA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathskiprootw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR PathSkipRootW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathstrippatha
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathStripPathA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathstrippathw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathStripPathW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathstriptoroota
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathStripToRootA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathstriptorootw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathStripToRootW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathundecoratea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathUndecorateA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathundecoratew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void PathUndecorateW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathunexpandenvstringsa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathUnExpandEnvStringsA(PSTR pszPath, [MarshalUsing(CountElementName = nameof(cchBuf))] PSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathunexpandenvstringsw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathUnExpandEnvStringsW(PWSTR pszPath, [MarshalUsing(CountElementName = nameof(cchBuf))] PWSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathunmakesystemfoldera
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathUnmakeSystemFolderA(PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathunmakesystemfolderw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathUnmakeSystemFolderW(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathunquotespacesa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathUnquoteSpacesA(PSTR lpsz);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-pathunquotespacesw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathUnquoteSpacesW(PWSTR lpsz);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pathyetanothermakeuniquename
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL PathYetAnotherMakeUniqueName([MarshalUsing(ConstantElementCount = 260)] PWSTR pszUniqueName, PWSTR pszPath, PWSTR pszShort, PWSTR pszFileSpec);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pickicondlg
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int PickIconDlg(HWND hwnd, [MarshalUsing(CountElementName = nameof(cchIconPath))] PWSTR pszIconPath, uint cchIconPath, nint /* optional int* */ piIconIndex);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pifmgr_closeproperties
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HANDLE PifMgr_CloseProperties(HANDLE hProps, uint flOpt);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pifmgr_getproperties
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int PifMgr_GetProperties(HANDLE hProps, PSTR pszGroup, nint /* optional void* */ lpProps, int cbProps, uint flOpt);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pifmgr_openproperties
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HANDLE PifMgr_OpenProperties(PWSTR pszApp, PWSTR pszPIF, uint hInf, uint flOpt);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-pifmgr_setproperties
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int PifMgr_SetProperties(HANDLE hProps, PSTR pszGroup, nint lpProps, int cbProps, uint flOpt);
    
    // https://learn.microsoft.com/windows/win32/api/propvarutil/nf-propvarutil-propvarianttostrret
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PropVariantToStrRet(in PROPVARIANT propvar, out STRRET pstrret);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pscoercetocanonicalvalue
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSCoerceToCanonicalValue(in PROPERTYKEY key, ref PROPVARIANT ppropvar);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pscreateadapterfrompropertystore
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSCreateAdapterFromPropertyStore([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyStore>))] IPropertyStore pps, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pscreatedelayedmultiplexpropertystore
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSCreateDelayedMultiplexPropertyStore(GETPROPERTYSTOREFLAGS flags, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDelayedPropertyStoreFactory>))] IDelayedPropertyStoreFactory pdpsf, [In][MarshalUsing(CountElementName = nameof(cStores))] uint[] rgStoreIds, uint cStores, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pscreatememorypropertystore
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSCreateMemoryPropertyStore(in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pscreatemultiplexpropertystore
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSCreateMultiplexPropertyStore([In][Out][MarshalUsing(CountElementName = nameof(cStores))] nint[] prgpunkStores, uint cStores, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pscreatepropertychangearray
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSCreatePropertyChangeArray(nint /* optional PROPERTYKEY* */ rgpropkey, nint /* optional PKA_FLAGS* */ rgflags, nint /* optional PROPVARIANT* */ rgpropvar, uint cChanges, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pscreatepropertystorefromobject
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSCreatePropertyStoreFromObject(nint punk, uint grfMode, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pscreatepropertystorefrompropertysetstorage
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSCreatePropertyStoreFromPropertySetStorage([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertySetStorage>))] IPropertySetStorage ppss, uint grfMode, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pscreatesimplepropertychange
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSCreateSimplePropertyChange(PKA_FLAGS flags, in PROPERTYKEY key, in PROPVARIANT propvar, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psenumeratepropertydescriptions
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSEnumeratePropertyDescriptions(PROPDESC_ENUMFILTER filterOn, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psformatfordisplay
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT PSFormatForDisplay(in PROPERTYKEY propkey, in PROPVARIANT propvar, PROPDESC_FORMAT_FLAGS pdfFlags, [MarshalUsing(CountElementName = nameof(cchText))] PWSTR pwszText, uint cchText);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psformatfordisplayalloc
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT PSFormatForDisplayAlloc(in PROPERTYKEY key, in PROPVARIANT propvar, PROPDESC_FORMAT_FLAGS pdff, out PWSTR ppszDisplay);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psformatpropertyvalue
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT PSFormatPropertyValue([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyStore>))] IPropertyStore pps, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyDescription>))] IPropertyDescription ppd, PROPDESC_FORMAT_FLAGS pdff, out PWSTR ppszDisplay);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetimagereferenceforvalue
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSGetImageReferenceForValue(in PROPERTYKEY propkey, in PROPVARIANT propvar, out PWSTR ppszImageRes);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetitempropertyhandler
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetItemPropertyHandler(nint punkItem, BOOL fReadWrite, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetitempropertyhandlerwithcreateobject
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetItemPropertyHandlerWithCreateObject(nint punkItem, BOOL fReadWrite, nint punkCreateObject, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetnamedpropertyfrompropertystorage
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetNamedPropertyFromPropertyStorage(PCUSERIALIZEDPROPSTORAGE psps, uint cb, PWSTR pszName, out PROPVARIANT ppropvar);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetnamefrompropertykey
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetNameFromPropertyKey(in PROPERTYKEY propkey, out PWSTR ppszCanonicalName);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetpropertydescription
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetPropertyDescription(in PROPERTYKEY propkey, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetpropertydescriptionbyname
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetPropertyDescriptionByName(PWSTR pszCanonicalName, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetpropertydescriptionlistfromstring
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetPropertyDescriptionListFromString(PWSTR pszPropList, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetpropertyfrompropertystorage
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetPropertyFromPropertyStorage(PCUSERIALIZEDPROPSTORAGE psps, uint cb, in PROPERTYKEY rpkey, out PROPVARIANT ppropvar);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetpropertykeyfromname
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetPropertyKeyFromName(PWSTR pszName, out PROPERTYKEY ppropkey);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetpropertysystem
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetPropertySystem(in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psgetpropertyvalue
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSGetPropertyValue([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyStore>))] IPropertyStore pps, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyDescription>))] IPropertyDescription ppd, out PROPVARIANT ppropvar);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pslookuppropertyhandlerclsid
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSLookupPropertyHandlerCLSID(PWSTR pszFilePath, out Guid pclsid);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_delete
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_Delete([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readbool
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadBOOL([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out BOOL value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readbstr
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadBSTR([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out BSTR value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readdword
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadDWORD([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out uint value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readguid
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadGUID([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out Guid value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readint
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadInt([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out int value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readlong
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadLONG([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out int value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readpointl
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadPOINTL([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out POINTL value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readpoints
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadPOINTS([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out POINTS value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readpropertykey
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadPropertyKey([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out PROPERTYKEY value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readrectl
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadRECTL([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out RECTL value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readshort
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadSHORT([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out short value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readstr
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadStr([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, [MarshalUsing(CountElementName = nameof(characterCount))] PWSTR value, int characterCount);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readstralloc
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadStrAlloc([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out PWSTR value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readstream
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadStream([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] out IStream value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readtype
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadType([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out VARIANT var, VARENUM type);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readulonglong
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadULONGLONG([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, out ulong value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_readunknown
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_ReadUnknown([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writebool
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteBOOL([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, BOOL value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writebstr
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteBSTR([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, BSTR value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writedword
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteDWORD([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, uint value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writeguid
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteGUID([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, in Guid value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writeint
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteInt([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, int value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writelong
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteLONG([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, int value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writepointl
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WritePOINTL([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, in POINTL value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writepoints
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WritePOINTS([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, in POINTS value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writepropertykey
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WritePropertyKey([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, in PROPERTYKEY value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writerectl
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteRECTL([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, in RECTL value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writeshort
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteSHORT([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, short value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writestr
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteStr([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, PWSTR value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writestream
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteStream([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] IStream value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writeulonglong
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteULONGLONG([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, ulong value);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertybag_writeunknown
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT PSPropertyBag_WriteUnknown([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyBag>))] IPropertyBag propBag, PWSTR propName, nint punk);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pspropertykeyfromstring
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSPropertyKeyFromString(PWSTR pszString, out PROPERTYKEY pkey);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psrefreshpropertyschema
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSRefreshPropertySchema();
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psregisterpropertyschema
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSRegisterPropertySchema(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-pssetpropertyvalue
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSSetPropertyValue([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyStore>))] IPropertyStore pps, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyDescription>))] IPropertyDescription ppd, in PROPVARIANT propvar);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psstringfrompropertykey
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSStringFromPropertyKey(in PROPERTYKEY pkey, [MarshalUsing(CountElementName = nameof(cch))] PWSTR psz, uint cch);
    
    // https://learn.microsoft.com/windows/win32/api/propsys/nf-propsys-psunregisterpropertyschema
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT PSUnregisterPropertySchema(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-qisearch
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT QISearch(nint that, in QITAB pqit, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-readcabinetstate
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ReadCabinetState(nint pcs, int cLength);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-realdrivetype
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int RealDriveType(int iDrive, BOOL fOKToHitNet);
    
    [LibraryImport("api-ms-win-core-psm-appnotify-l1-1-1")]
    [PreserveSig]
    public static partial uint RegisterAppConstrainedChangeNotification(PAPPCONSTRAIN_CHANGE_ROUTINE Routine, nint /* optional void* */ Context, out PAPPCONSTRAIN_REGISTRATION Registration);
    
    // https://learn.microsoft.com/windows/win32/api/appnotify/nf-appnotify-registerappstatechangenotification
    [LibraryImport("api-ms-win-core-psm-appnotify-l1-1-0")]
    [PreserveSig]
    public static partial uint RegisterAppStateChangeNotification(PAPPSTATE_CHANGE_ROUTINE Routine, nint /* optional void* */ Context, out PAPPSTATE_REGISTRATION Registration);
    
    // https://learn.microsoft.com/windows/win32/api/shellscalingapi/nf-shellscalingapi-registerscalechangeevent
    [LibraryImport("api-ms-win-shcore-scaling-l1-1-1")]
    [SupportedOSPlatform("windows8.1")]
    [PreserveSig]
    public static partial HRESULT RegisterScaleChangeEvent(HANDLE hEvent, out nuint pdwCookie);
    
    // https://learn.microsoft.com/windows/win32/api/shellscalingapi/nf-shellscalingapi-registerscalechangenotifications
    [LibraryImport("api-ms-win-shcore-scaling-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT RegisterScaleChangeNotifications(DISPLAY_DEVICE_TYPE displayDevice, HWND hwndNotify, uint uMsgNotify, out uint pdwCookie);
    
    // https://learn.microsoft.com/windows/win32/api/commctrl/nf-commctrl-removewindowsubclass
    [LibraryImport("COMCTL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL RemoveWindowSubclass(HWND hWnd, SUBCLASSPROC pfnSubclass, nuint uIdSubclass);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-restartdialog
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int RestartDialog(HWND hwnd, PWSTR pszPrompt, uint dwReturn);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-restartdialogex
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int RestartDialogEx(HWND hwnd, PWSTR pszPrompt, uint dwReturn, uint dwReasonCode);
    
    // https://learn.microsoft.com/windows/win32/api/shellscalingapi/nf-shellscalingapi-revokescalechangenotifications
    [LibraryImport("api-ms-win-shcore-scaling-l1-1-0")]
    [SupportedOSPlatform("windows8.0")]
    [PreserveSig]
    public static partial HRESULT RevokeScaleChangeNotifications(DISPLAY_DEVICE_TYPE displayDevice, uint dwCookie);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-setcurrentprocessexplicitappusermodelid
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT SetCurrentProcessExplicitAppUserModelID(PWSTR AppID);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-setmenucontexthelpid
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SetMenuContextHelpId(HMENU param0, uint param1);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-setwindowcontexthelpid
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SetWindowContextHelpId(HWND param0, uint param1);
    
    // https://learn.microsoft.com/windows/win32/api/commctrl/nf-commctrl-setwindowsubclass
    [LibraryImport("COMCTL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SetWindowSubclass(HWND hWnd, SUBCLASSPROC pfnSubclass, nuint uIdSubclass, nuint dwRefData);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-shadddefaultpropertiesbyext
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHAddDefaultPropertiesByExt(PWSTR pszExt, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyStore>))] IPropertyStore pPropStore);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shaddfrompropsheetextarray
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial uint SHAddFromPropSheetExtArray(HPSXA hpsxa, LPFNSVADDPROPSHEETPAGE lpfnAddPage, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shaddtorecentdocs
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void SHAddToRecentDocs(uint uFlags, nint /* optional void* */ pv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shalloc
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial nint SHAlloc(nuint cb);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shallocshared
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HANDLE SHAllocShared(nint /* optional void* */ pvData, uint dwSize, uint dwProcessId);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shansitoansi
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHAnsiToAnsi(PSTR pszSrc, [MarshalUsing(CountElementName = nameof(cchBuf))] PSTR pszDst, int cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shansitounicode
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHAnsiToUnicode(PSTR pszSrc, [MarshalUsing(CountElementName = nameof(cwchBuf))] PWSTR pwszDst, int cwchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shappbarmessage
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nuint SHAppBarMessage(uint dwMessage, ref APPBARDATA pData);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shassocenumhandlers
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHAssocEnumHandlers(PWSTR pszExtra, ASSOC_FILTER afFilter, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IEnumAssocHandlers>))] out IEnumAssocHandlers ppEnumHandler);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shassocenumhandlersforprotocolbyapplication
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT SHAssocEnumHandlersForProtocolByApplication(PWSTR protocol, in Guid riid, out nint /* void */ enumHandlers);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shautocomplete
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHAutoComplete(HWND hwndEdit, SHELL_AUTOCOMPLETE_FLAGS dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shbindtofolderidlistparent
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHBindToFolderIDListParent([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder?>))] IShellFolder? psfRoot, nint pidl, in Guid riid, out nint /* void */ ppv, nint /* optional nint** */ ppidlLast);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shbindtofolderidlistparentex
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHBindToFolderIDListParentEx([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder?>))] IShellFolder? psfRoot, nint pidl, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx?>))] IBindCtx? ppbc, in Guid riid, out nint /* void */ ppv, nint /* optional nint** */ ppidlLast);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shbindtoobject
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHBindToObject([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder?>))] IShellFolder? psf, nint pidl, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx?>))] IBindCtx? pbc, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shbindtoparent
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHBindToParent(nint pidl, in Guid riid, out nint /* void */ ppv, nint /* optional nint** */ ppidlLast);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shbrowseforfoldera
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint SHBrowseForFolderA(in BROWSEINFOA lpbi);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shbrowseforfolderw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint SHBrowseForFolderW(in BROWSEINFOW lpbi);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shchangenotification_lock
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HANDLE SHChangeNotification_Lock(HANDLE hChange, uint dwProcId, nint /* optional nint*** */ pppidl, nint /* optional int* */ plEvent);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shchangenotification_unlock
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHChangeNotification_Unlock(HANDLE hLock);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shchangenotify
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void SHChangeNotify(int wEventId, SHCNF_FLAGS uFlags, nint /* optional void* */ dwItem1, nint /* optional void* */ dwItem2);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shchangenotifyderegister
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHChangeNotifyDeregister(uint ulID);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shchangenotifyregister
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial uint SHChangeNotifyRegister(HWND hwnd, SHCNRF_SOURCE fSources, int fEvents, uint wMsg, int cEntries, nint pshcne);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-shchangenotifyregisterthread
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial void SHChangeNotifyRegisterThread(SCNRT_STATUS status);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shclonespecialidlist
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint SHCloneSpecialIDList(HWND hwnd, int csidl, BOOL fCreate);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shclsidfromstring
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHCLSIDFromString(PWSTR psz, out Guid pclsid);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcocreateinstance
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHCoCreateInstance(PWSTR pszCLSID, nint /* optional Guid* */ pclsid, nint pUnkOuter, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shcopykeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHCopyKeyA(HKEY hkeySrc, PSTR pszSrcSubKey, HKEY hkeyDest, uint fReserved);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shcopykeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHCopyKeyW(HKEY hkeySrc, PWSTR pszSrcSubKey, HKEY hkeyDest, uint fReserved);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreateassociationregistration
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateAssociationRegistration(in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcreatedataobject
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateDataObject(nint /* optional nint* */ pidlFolder, uint cidl, nint /* optional nint** */ apidl, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject?>))] IDataObject? pdtInner, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcreatedefaultcontextmenu
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateDefaultContextMenu(in DEFCONTEXTMENU pdcm, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreatedefaultextracticon
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateDefaultExtractIcon(in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-shcreatedefaultpropertiesop
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateDefaultPropertiesOp([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IFileOperation>))] out IFileOperation ppFileOp);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcreatedirectory
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int SHCreateDirectory(HWND hwnd, PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcreatedirectoryexa
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHCreateDirectoryExA(HWND hwnd, PSTR pszPath, nint /* optional SECURITY_ATTRIBUTES* */ psa);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcreatedirectoryexw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHCreateDirectoryExW(HWND hwnd, PWSTR pszPath, nint /* optional SECURITY_ATTRIBUTES* */ psa);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcreatefileextracticonw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHCreateFileExtractIconW(PWSTR pszFile, uint dwFileAttributes, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreateitemfromidlist
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateItemFromIDList(nint pidl, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreateitemfromparsingname
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateItemFromParsingName(PWSTR pszPath, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx?>))] IBindCtx? pbc, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreateitemfromrelativename
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateItemFromRelativeName([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiParent, PWSTR pszName, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx?>))] IBindCtx? pbc, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreateiteminknownfolder
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateItemInKnownFolder(in Guid kfid, uint dwKFFlags, PWSTR pszItem, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreateitemwithparent
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateItemWithParent(nint /* optional nint* */ pidlParent, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder?>))] IShellFolder? psfParent, nint pidl, in Guid riid, out nint /* void */ ppvItem);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shcreatememstream
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [return: MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))]
    public static partial IStream SHCreateMemStream(nint /* optional byte* */ pInit, uint cbInit);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shcreateprocessasuserw
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHCreateProcessAsUserW(ref SHCREATEPROCESSINFOW pscpi);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-shcreatepropsheetextarray
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HPSXA SHCreatePropSheetExtArray(HKEY hKey, PWSTR pszSubKey, uint max_iface);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-shcreatequerycancelautoplaymoniker
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHCreateQueryCancelAutoPlayMoniker([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMoniker>))] out IMoniker ppmoniker);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcreateshellfolderview
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHCreateShellFolderView(in SFV_CREATE pcsfv, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellView>))] out IShellView ppsv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcreateshellfolderviewex
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHCreateShellFolderViewEx(in CSFV pcsfv, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellView>))] out IShellView ppsv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcreateshellitem
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHCreateShellItem(nint /* optional nint* */ pidlParent, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder?>))] IShellFolder? psfParent, nint pidl, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] out IShellItem ppsi);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreateshellitemarray
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateShellItemArray(nint /* optional nint* */ pidlParent, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder?>))] IShellFolder? psf, uint cidl, nint /* optional nint** */ ppidl, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItemArray>))] out IShellItemArray ppsiItemArray);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreateshellitemarrayfromdataobject
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateShellItemArrayFromDataObject([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject>))] IDataObject pdo, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreateshellitemarrayfromidlists
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateShellItemArrayFromIDLists(uint cidl, [In][Out][MarshalUsing(CountElementName = nameof(cidl))] nint[] rgpidl, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItemArray>))] out IShellItemArray ppsiItemArray);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shcreateshellitemarrayfromshellitem
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHCreateShellItemArrayFromShellItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shcreateshellpalette
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HPALETTE SHCreateShellPalette(HDC hdc);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shcreatestdenumfmtetc
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHCreateStdEnumFmtEtc(uint cfmt, [In][MarshalUsing(CountElementName = nameof(cfmt))] FORMATETC[] afmt, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IEnumFORMATETC>))] out IEnumFORMATETC ppenumFormatEtc);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shcreatestreamonfilea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHCreateStreamOnFileA(PSTR pszFile, uint grfMode, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] out IStream ppstm);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shcreatestreamonfileex
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHCreateStreamOnFileEx(PWSTR pszFile, uint grfMode, uint dwAttributes, BOOL fCreate, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream?>))] IStream? pstmTemplate, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] out IStream ppstm);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shcreatestreamonfilew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHCreateStreamOnFileW(PWSTR pszFile, uint grfMode, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))] out IStream ppstm);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shcreatethread
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHCreateThread(LPTHREAD_START_ROUTINE pfnThreadProc, nint /* optional void* */ pData, uint flags, LPTHREAD_START_ROUTINE? pfnCallback);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shcreatethreadref
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHCreateThreadRef(ref int pcRef, out nint ppunk);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shcreatethreadwithhandle
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHCreateThreadWithHandle(LPTHREAD_START_ROUTINE pfnThreadProc, nint /* optional void* */ pData, uint flags, LPTHREAD_START_ROUTINE? pfnCallback, nint /* optional HANDLE* */ pHandle);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shdefextracticona
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHDefExtractIconA(PSTR pszIconFile, int iIndex, uint uFlags, nint /* optional HICON* */ phiconLarge, nint /* optional HICON* */ phiconSmall, uint nIconSize);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shdefextracticonw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHDefExtractIconW(PWSTR pszIconFile, int iIndex, uint uFlags, nint /* optional HICON* */ phiconLarge, nint /* optional HICON* */ phiconSmall, uint nIconSize);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shdeleteemptykeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHDeleteEmptyKeyA(HKEY hkey, PSTR pszSubKey);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shdeleteemptykeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHDeleteEmptyKeyW(HKEY hkey, PWSTR pszSubKey);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shdeletekeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHDeleteKeyA(HKEY hkey, PSTR pszSubKey);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shdeletekeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHDeleteKeyW(HKEY hkey, PWSTR pszSubKey);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shdeletevaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHDeleteValueA(HKEY hkey, PSTR pszSubKey, PSTR pszValue);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shdeletevaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHDeleteValueW(HKEY hkey, PWSTR pszSubKey, PWSTR pszValue);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shdestroypropsheetextarray
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void SHDestroyPropSheetExtArray(HPSXA hpsxa);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shdodragdrop
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHDoDragDrop(HWND hwnd, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject>))] IDataObject pdata, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDropSource?>))] IDropSource? pdsrc, DROPEFFECT dwEffect, out DROPEFFECT pdwEffect);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shell_getcachedimageindex
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int Shell_GetCachedImageIndex(PWSTR pwszIconPath, int iIconIndex, uint uIconFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shell_getcachedimageindexa
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int Shell_GetCachedImageIndexA(PSTR pszIconPath, int iIconIndex, uint uIconFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shell_getcachedimageindexw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int Shell_GetCachedImageIndexW(PWSTR pszIconPath, int iIconIndex, uint uIconFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shell_getimagelists
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL Shell_GetImageLists(nint /* optional HIMAGELIST* */ phiml, nint /* optional HIMAGELIST* */ phimlSmall);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shell_mergemenus
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint Shell_MergeMenus(HMENU hmDst, HMENU hmSrc, uint uInsert, uint uIDAdjust, uint uIDAdjustMax, MM_FLAGS uFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shell_notifyicona
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL Shell_NotifyIconA(NOTIFY_ICON_MESSAGE dwMessage, in NOTIFYICONDATAA lpData);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shell_notifyicongetrect
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT Shell_NotifyIconGetRect(in NOTIFYICONIDENTIFIER identifier, out RECT iconLocation);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shell_notifyiconw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL Shell_NotifyIconW(NOTIFY_ICON_MESSAGE dwMessage, in NOTIFYICONDATAW lpData);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shellabouta
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int ShellAboutA(HWND hWnd, PSTR szApp, PSTR szOtherStuff, HICON hIcon);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shellaboutw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int ShellAboutW(HWND hWnd, PWSTR szApp, PWSTR szOtherStuff, HICON hIcon);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shellexecutea
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HINSTANCE ShellExecuteA(HWND hwnd, PSTR lpOperation, PSTR lpFile, PSTR lpParameters, PSTR lpDirectory, SHOW_WINDOW_CMD nShowCmd);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shellexecuteexa
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ShellExecuteExA(ref SHELLEXECUTEINFOA pExecInfo);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shellexecuteexw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ShellExecuteExW(ref SHELLEXECUTEINFOW pExecInfo);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shellexecutew
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HINSTANCE ShellExecuteW(HWND hwnd, PWSTR lpOperation, PWSTR lpFile, PWSTR lpParameters, PWSTR lpDirectory, SHOW_WINDOW_CMD nShowCmd);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shellmessageboxa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int ShellMessageBoxA(HINSTANCE hAppInst, HWND hWnd, PSTR lpcText, PSTR lpcTitle, MESSAGEBOX_STYLE fuStyle);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shellmessageboxw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int ShellMessageBoxW(HINSTANCE hAppInst, HWND hWnd, PWSTR lpcText, PWSTR lpcTitle, MESSAGEBOX_STYLE fuStyle);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shemptyrecyclebina
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHEmptyRecycleBinA(HWND hwnd, PSTR pszRootPath, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shemptyrecyclebinw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHEmptyRecycleBinW(HWND hwnd, PWSTR pszRootPath, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shenumerateunreadmailaccountsw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHEnumerateUnreadMailAccountsW(HKEY hKeyUser, uint dwIndex, [MarshalUsing(CountElementName = nameof(cchMailAddress))] PWSTR pszMailAddress, int cchMailAddress);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shenumkeyexa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHEnumKeyExA(HKEY hkey, uint dwIndex, [MarshalUsing(CountElementName = nameof(pcchName))] PSTR pszName, ref uint pcchName);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shenumkeyexw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHEnumKeyExW(HKEY hkey, uint dwIndex, [MarshalUsing(CountElementName = nameof(pcchName))] PWSTR pszName, ref uint pcchName);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shenumvaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHEnumValueA(HKEY hkey, uint dwIndex, [MarshalUsing(CountElementName = nameof(pcchValueName))] PSTR pszValueName, nint /* optional uint* */ pcchValueName, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shenumvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHEnumValueW(HKEY hkey, uint dwIndex, [MarshalUsing(CountElementName = nameof(pcchValueName))] PWSTR pszValueName, nint /* optional uint* */ pcchValueName, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shevaluatesystemcommandtemplate
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHEvaluateSystemCommandTemplate(PWSTR pszCmdTemplate, out PWSTR ppszApplication, nint /* optional PWSTR* */ ppszCommandLine, nint /* optional PWSTR* */ ppszParameters);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shfileoperationa
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int SHFileOperationA(ref SHFILEOPSTRUCTA lpFileOp);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shfileoperationw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int SHFileOperationW(ref SHFILEOPSTRUCTW lpFileOp);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shfind_initmenupopup
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [return: MarshalUsing(typeof(UniqueComInterfaceMarshaller<IContextMenu>))]
    public static partial IContextMenu SHFind_InitMenuPopup(HMENU hmenu, HWND hwndOwner, uint idCmdFirst, uint idCmdLast);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shfindfiles
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHFindFiles(nint /* optional nint* */ pidlFolder, nint /* optional nint* */ pidlSaveFile);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shflushsfcache
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void SHFlushSFCache();
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shformatdatetimea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int SHFormatDateTimeA(in FILETIME pft, nint /* optional uint* */ pdwFlags, [MarshalUsing(CountElementName = nameof(cchBuf))] PSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shformatdatetimew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int SHFormatDateTimeW(in FILETIME pft, nint /* optional uint* */ pdwFlags, [MarshalUsing(CountElementName = nameof(cchBuf))] PWSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shformatdrive
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint SHFormatDrive(HWND hwnd, uint drive, SHFMT_ID fmtID, uint options);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shfree
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void SHFree(nint /* optional void* */ pv);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shfreenamemappings
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void SHFreeNameMappings(HANDLE hNameMappings);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shfreeshared
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHFreeShared(HANDLE hData, uint dwProcessId);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetattributesfromdataobject
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetAttributesFromDataObject([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject?>))] IDataObject? pdo, uint dwAttributeMask, nint /* optional uint* */ pdwAttributes, nint /* optional uint* */ pcItems);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetdatafromidlista
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetDataFromIDListA([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder>))] IShellFolder psf, nint pidl, SHGDFIL_FORMAT nFormat, nint pv, int cb);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetdatafromidlistw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetDataFromIDListW([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder>))] IShellFolder psf, nint pidl, SHGDFIL_FORMAT nFormat, nint pv, int cb);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetdesktopfolder
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetDesktopFolder([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder>))] out IShellFolder ppshf);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetdiskfreespaceexa
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHGetDiskFreeSpaceExA(PSTR pszDirectoryName, nint /* optional ulong* */ pulFreeBytesAvailableToCaller, nint /* optional ulong* */ pulTotalNumberOfBytes, nint /* optional ulong* */ pulTotalNumberOfFreeBytes);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetdiskfreespaceexw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHGetDiskFreeSpaceExW(PWSTR pszDirectoryName, nint /* optional ulong* */ pulFreeBytesAvailableToCaller, nint /* optional ulong* */ pulTotalNumberOfBytes, nint /* optional ulong* */ pulTotalNumberOfFreeBytes);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetdrivemedia
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHGetDriveMedia(PWSTR pszDrive, out uint pdwMediaContent);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetfileinfoa
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nuint SHGetFileInfoA(PSTR pszPath, FILE_FLAGS_AND_ATTRIBUTES dwFileAttributes, nint /* optional SHFILEINFOA* */ psfi, uint cbFileInfo, SHGFI_FLAGS uFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetfileinfow
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nuint SHGetFileInfoW(PWSTR pszPath, FILE_FLAGS_AND_ATTRIBUTES dwFileAttributes, nint /* optional SHFILEINFOW* */ psfi, uint cbFileInfo, SHGFI_FLAGS uFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetfolderlocation
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHGetFolderLocation(HWND hwnd, int csidl, HANDLE hToken, uint dwFlags, out nint ppidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetfolderpatha
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHGetFolderPathA(HWND hwnd, int csidl, HANDLE hToken, uint dwFlags, [MarshalUsing(ConstantElementCount = 260)] PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetfolderpathandsubdira
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetFolderPathAndSubDirA(HWND hwnd, int csidl, HANDLE hToken, uint dwFlags, PSTR pszSubDir, [MarshalUsing(ConstantElementCount = 260)] PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetfolderpathandsubdirw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetFolderPathAndSubDirW(HWND hwnd, int csidl, HANDLE hToken, uint dwFlags, PWSTR pszSubDir, [MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetfolderpathw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHGetFolderPathW(HWND hwnd, int csidl, HANDLE hToken, uint dwFlags, [MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgeticonoverlayindexa
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHGetIconOverlayIndexA(PSTR pszIconPath, int iIconIndex);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgeticonoverlayindexw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHGetIconOverlayIndexW(PWSTR pszIconPath, int iIconIndex);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shgetidlistfromobject
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHGetIDListFromObject(nint punk, out nint ppidl);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetimagelist
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetImageList(int iImageList, in Guid riid, out nint /* void */ ppvObj);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetinstanceexplorer
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetInstanceExplorer(out nint ppunk);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shgetinversecmap
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHGetInverseCMAP(nint /* byte array */ pbMap, uint cbMap);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shgetitemfromdataobject
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT SHGetItemFromDataObject([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject>))] IDataObject pdtobj, DATAOBJ_GET_ITEM_FLAGS dwFlags, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shgetitemfromobject
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT SHGetItemFromObject(nint punk, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetknownfolderidlist
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHGetKnownFolderIDList(in Guid rfid, uint dwFlags, HANDLE hToken, out nint ppidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetknownfolderitem
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT SHGetKnownFolderItem(in Guid rfid, KNOWN_FOLDER_FLAG flags, HANDLE hToken, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetknownfolderpath
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHGetKnownFolderPath(in Guid rfid, uint dwFlags, HANDLE hToken, out PWSTR ppszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetlocalizedname
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHGetLocalizedName(PWSTR pszPath, [MarshalUsing(CountElementName = nameof(cch))] PWSTR pszResModule, uint cch, out int pidsRes);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetmalloc
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetMalloc([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IMalloc>))] out IMalloc ppMalloc);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shgetnamefromidlist
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHGetNameFromIDList(nint pidl, SIGDN sigdnName, out PWSTR ppszName);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetnewlinkinfoa
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHGetNewLinkInfoA(PSTR pszLinkTo, PSTR pszDir, [MarshalUsing(ConstantElementCount = 260)] PSTR pszName, out BOOL pfMustCopy, uint uFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetnewlinkinfow
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHGetNewLinkInfoW(PWSTR pszLinkTo, PWSTR pszDir, [MarshalUsing(ConstantElementCount = 260)] PWSTR pszName, out BOOL pfMustCopy, uint uFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetpathfromidlista
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHGetPathFromIDListA(nint pidl, [MarshalUsing(ConstantElementCount = 260)] PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetpathfromidlistex
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHGetPathFromIDListEx(nint pidl, [MarshalUsing(CountElementName = nameof(cchPath))] PWSTR pszPath, uint cchPath, GPFIDL_FLAGS uOpts);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetpathfromidlistw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHGetPathFromIDListW(nint pidl, [MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetpropertystoreforwindow
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT SHGetPropertyStoreForWindow(HWND hwnd, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shgetpropertystorefromidlist
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHGetPropertyStoreFromIDList(nint pidl, GETPROPERTYSTOREFLAGS flags, in Guid riid, out nint ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shgetpropertystorefromparsingname
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHGetPropertyStoreFromParsingName(PWSTR pszPath, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx?>))] IBindCtx? pbc, GETPROPERTYSTOREFLAGS flags, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetrealidl
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetRealIDL([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder>))] IShellFolder psf, nint pidlSimple, out nint ppidlReal);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetsetfoldercustomsettings
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetSetFolderCustomSettings(ref SHFOLDERCUSTOMSETTINGS pfcs, PWSTR pszPath, uint dwReadWrite);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetsetsettings
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void SHGetSetSettings(nint /* optional SHELLSTATEA* */ lpss, SSF_MASK dwMask, BOOL bSet);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetsettings
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void SHGetSettings(out SHELLFLAGSTATE psfs, uint dwMask);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetspecialfolderlocation
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHGetSpecialFolderLocation(HWND hwnd, int csidl, out nint ppidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetspecialfolderpatha
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHGetSpecialFolderPathA(HWND hwnd, [MarshalUsing(ConstantElementCount = 260)] PSTR pszPath, int csidl, BOOL fCreate);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetspecialfolderpathw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHGetSpecialFolderPathW(HWND hwnd, [MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath, int csidl, BOOL fCreate);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetstockiconinfo
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHGetStockIconInfo(SHSTOCKICONID siid, SHGSI_FLAGS uFlags, ref SHSTOCKICONINFO psii);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shgettemporarypropertyforitem
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHGetTemporaryPropertyForItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, in PROPERTYKEY propkey, out PROPVARIANT ppropvar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shgetthreadref
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHGetThreadRef(out nint ppunk);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetunreadmailcountw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetUnreadMailCountW(HKEY hKeyUser, PWSTR pszMailAddress, nint /* optional uint* */ pdwCount, nint /* optional FILETIME* */ pFileTime, [MarshalUsing(CountElementName = nameof(cchShellExecuteCommand))] PWSTR pszShellExecuteCommand, int cchShellExecuteCommand);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shgetvaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHGetValueA(HKEY hkey, PSTR pszSubKey, PSTR pszValue, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shgetvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHGetValueW(HKEY hkey, PWSTR pszSubKey, PWSTR pszValue, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shgetviewstatepropertybag
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHGetViewStatePropertyBag(nint /* optional nint* */ pidl, PWSTR pszBagName, uint dwFlags, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shglobalcounterdecrement
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial int SHGlobalCounterDecrement(SHGLOBALCOUNTER id);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shglobalcountergetvalue
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial int SHGlobalCounterGetValue(SHGLOBALCOUNTER id);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shglobalcounterincrement
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial int SHGlobalCounterIncrement(SHGLOBALCOUNTER id);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shhandleupdateimage
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int SHHandleUpdateImage(nint pidlExtra);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shilcreatefrompath
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHILCreateFromPath(PWSTR pszPath, out nint ppidl, nint /* optional uint* */ rgfInOut);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shinvokeprintercommanda
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHInvokePrinterCommandA(HWND hwnd, uint uAction, PSTR lpBuf1, PSTR lpBuf2, BOOL fModal);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shinvokeprintercommandw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHInvokePrinterCommandW(HWND hwnd, uint uAction, PWSTR lpBuf1, PWSTR lpBuf2, BOOL fModal);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shisfileavailableoffline
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHIsFileAvailableOffline(PWSTR pwszPath, nint /* optional uint* */ pdwStatus);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shislowmemorymachine
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHIsLowMemoryMachine(uint dwType);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shlimitinputedit
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHLimitInputEdit(HWND hwndEdit, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder>))] IShellFolder psf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shloadindirectstring
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHLoadIndirectString(PWSTR pszSource, [MarshalUsing(CountElementName = nameof(cchOutBuf))] PWSTR pszOutBuf, uint cchOutBuf, nint /* optional void** */ ppvReserved);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shloadinproc
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHLoadInProc(in Guid rclsid);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shloadnonloadediconoverlayidentifiers
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHLoadNonloadedIconOverlayIdentifiers();
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shlockshared
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial nint SHLockShared(HANDLE hData, uint dwProcessId);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shmappidltosystemimagelistindex
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int SHMapPIDLToSystemImageListIndex([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellFolder>))] IShellFolder pshf, nint pidl, nint /* optional int* */ piIndexSel);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shmessageboxchecka
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int SHMessageBoxCheckA(HWND hwnd, PSTR pszText, PSTR pszCaption, uint uType, int iDefault, PSTR pszRegVal);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shmessageboxcheckw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int SHMessageBoxCheckW(HWND hwnd, PWSTR pszText, PWSTR pszCaption, uint uType, int iDefault, PWSTR pszRegVal);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-shmultifileproperties
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHMultiFileProperties([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject>))] IDataObject pdtobj, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shobjectproperties
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHObjectProperties(HWND hwnd, uint shopObjectType, PWSTR pszObjectName, PWSTR pszPropertyPage);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shopenfolderandselectitems
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHOpenFolderAndSelectItems(nint pidlFolder, uint cidl, nint /* optional nint** */ apidl, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-shopenpropsheetw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHOpenPropSheetW(PWSTR pszCaption, nint /* optional HKEY* */ ahkeys, uint ckeys, nint /* optional Guid* */ pclsidDefault, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IDataObject>))] IDataObject pdtobj, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellBrowser?>))] IShellBrowser? psb, PWSTR pStartPage);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shopenregstream2a
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [return: MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))]
    public static partial IStream SHOpenRegStream2A(HKEY hkey, PSTR pszSubkey, PSTR pszValue, uint grfMode);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shopenregstream2w
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [return: MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))]
    public static partial IStream SHOpenRegStream2W(HKEY hkey, PWSTR pszSubkey, PWSTR pszValue, uint grfMode);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shopenregstreama
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [return: MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))]
    public static partial IStream SHOpenRegStreamA(HKEY hkey, PSTR pszSubkey, PSTR pszValue, uint grfMode);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shopenregstreamw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [return: MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStream>))]
    public static partial IStream SHOpenRegStreamW(HKEY hkey, PWSTR pszSubkey, PWSTR pszValue, uint grfMode);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shopenwithdialog
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHOpenWithDialog(HWND hwndParent, in OPENASINFO poainfo);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-showcaret
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ShowCaret(HWND hWnd);
    
    // https://learn.microsoft.com/windows/win32/api/gamingtcui/nf-gamingtcui-showchangefriendrelationshipui
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-0")]
    [PreserveSig]
    public static partial HRESULT ShowChangeFriendRelationshipUI(HSTRING targetUserXuid, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-2")]
    [PreserveSig]
    public static partial HRESULT ShowChangeFriendRelationshipUIForUser([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IInspectable>))] IInspectable user, HSTRING targetUserXuid, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("WININET")]
    [PreserveSig]
    public static partial uint ShowClientAuthCerts(HWND hWndParent);
    
    [LibraryImport("KERNEL32")]
    [PreserveSig]
    public static partial int ShowConsoleCursor(HANDLE hConsoleOutput, BOOL bShow);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-showcursor
    [LibraryImport("USER32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int ShowCursor(BOOL bShow);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-4")]
    [PreserveSig]
    public static partial HRESULT ShowCustomizeUserProfileUI(GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-4")]
    [PreserveSig]
    public static partial HRESULT ShowCustomizeUserProfileUIForUser([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IInspectable>))] IInspectable user, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-4")]
    [PreserveSig]
    public static partial HRESULT ShowFindFriendsUI(GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-4")]
    [PreserveSig]
    public static partial HRESULT ShowFindFriendsUIForUser([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IInspectable>))] IInspectable user, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-4")]
    [PreserveSig]
    public static partial HRESULT ShowGameInfoUI(uint titleId, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-4")]
    [PreserveSig]
    public static partial HRESULT ShowGameInfoUIForUser([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IInspectable>))] IInspectable user, uint titleId, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    // https://learn.microsoft.com/windows/win32/api/gamingtcui/nf-gamingtcui-showgameinviteui
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-0")]
    [PreserveSig]
    public static partial HRESULT ShowGameInviteUI(HSTRING serviceConfigurationId, HSTRING sessionTemplateName, HSTRING sessionId, HSTRING invitationDisplayText, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-2")]
    [PreserveSig]
    public static partial HRESULT ShowGameInviteUIForUser([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IInspectable>))] IInspectable user, HSTRING serviceConfigurationId, HSTRING sessionTemplateName, HSTRING sessionId, HSTRING invitationDisplayText, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-3")]
    [PreserveSig]
    public static partial HRESULT ShowGameInviteUIWithContext(HSTRING serviceConfigurationId, HSTRING sessionTemplateName, HSTRING sessionId, HSTRING invitationDisplayText, HSTRING customActivationContext, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-3")]
    [PreserveSig]
    public static partial HRESULT ShowGameInviteUIWithContextForUser([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IInspectable>))] IInspectable user, HSTRING serviceConfigurationId, HSTRING sessionTemplateName, HSTRING sessionId, HSTRING invitationDisplayText, HSTRING customActivationContext, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    // https://learn.microsoft.com/windows/win32/api/commctrl/nf-commctrl-showhidemenuctl
    [LibraryImport("COMCTL32", SetLastError = true)]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ShowHideMenuCtl(HWND hWnd, nuint uFlags, in int lpInfo);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-showownedpopups
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ShowOwnedPopups(HWND hWnd, BOOL fShow);
    
    // https://learn.microsoft.com/windows/win32/api/gamingtcui/nf-gamingtcui-showplayerpickerui
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-0")]
    [PreserveSig]
    public static partial HRESULT ShowPlayerPickerUI(HSTRING promptDisplayText, [In][MarshalUsing(CountElementName = nameof(xuidsCount))] HSTRING[] xuids, nuint xuidsCount, nint /* optional HSTRING* */ preSelectedXuids, nuint preSelectedXuidsCount, nuint minSelectionCount, nuint maxSelectionCount, PlayerPickerUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-2")]
    [PreserveSig]
    public static partial HRESULT ShowPlayerPickerUIForUser([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IInspectable>))] IInspectable user, HSTRING promptDisplayText, [In][MarshalUsing(CountElementName = nameof(xuidsCount))] HSTRING[] xuids, nuint xuidsCount, nint /* optional HSTRING* */ preSelectedXuids, nuint preSelectedXuidsCount, nuint minSelectionCount, nuint maxSelectionCount, PlayerPickerUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    // https://learn.microsoft.com/windows/win32/api/gamingtcui/nf-gamingtcui-showprofilecardui
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-0")]
    [PreserveSig]
    public static partial HRESULT ShowProfileCardUI(HSTRING targetUserXuid, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-2")]
    [PreserveSig]
    public static partial HRESULT ShowProfileCardUIForUser([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IInspectable>))] IInspectable user, HSTRING targetUserXuid, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-showscrollbar
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ShowScrollBar(HWND hWnd, SCROLLBAR_CONSTANTS wBar, BOOL bShow);
    
    [LibraryImport("WININET")]
    [PreserveSig]
    public static partial uint ShowSecurityInfo(HWND hWndParent, in INTERNET_SECURITY_INFO pSecurityInfo);
    
    // https://learn.microsoft.com/windows/win32/api/gamingtcui/nf-gamingtcui-showtitleachievementsui
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-0")]
    [PreserveSig]
    public static partial HRESULT ShowTitleAchievementsUI(uint titleId, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-2")]
    [PreserveSig]
    public static partial HRESULT ShowTitleAchievementsUIForUser([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IInspectable>))] IInspectable user, uint titleId, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-4")]
    [PreserveSig]
    public static partial HRESULT ShowUserSettingsUI(GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    [LibraryImport("api-ms-win-gaming-tcui-l1-1-4")]
    [PreserveSig]
    public static partial HRESULT ShowUserSettingsUIForUser([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IInspectable>))] IInspectable user, GameUICompletionRoutine completionRoutine, nint /* optional void* */ context);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-showwindow
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ShowWindow(HWND hWnd, SHOW_WINDOW_CMD nCmdShow);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-showwindowasync
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL ShowWindowAsync(HWND hWnd, SHOW_WINDOW_CMD nCmdShow);
    
    [LibraryImport("WININET")]
    [PreserveSig]
    public static partial uint ShowX509EncodedCertificate(HWND hWndParent, nint /* byte array */ lpCert, uint cbCert);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shparsedisplayname
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHParseDisplayName(PWSTR pszName, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx?>))] IBindCtx? pbc, out nint ppidl, uint sfgaoIn, nint /* optional uint* */ psfgaoOut);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shpathprepareforwritea
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHPathPrepareForWriteA(HWND hwnd, nint punkEnableModless, PSTR pszPath, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shpathprepareforwritew
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHPathPrepareForWriteW(HWND hwnd, nint punkEnableModless, PWSTR pszPath, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shpropstgcreate
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHPropStgCreate([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertySetStorage>))] IPropertySetStorage psstg, in Guid fmtid, nint /* optional Guid* */ pclsid, uint grfFlags, uint grfMode, uint dwDisposition, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyStorage>))] out IPropertyStorage ppstg, nint /* optional uint* */ puCodePage);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shpropstgreadmultiple
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHPropStgReadMultiple([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyStorage>))] IPropertyStorage pps, uint uCodePage, uint cpspec, [In][MarshalUsing(CountElementName = nameof(cpspec))] PROPSPEC[] rgpspec, [In][Out][MarshalUsing(CountElementName = nameof(cpspec))] PROPVARIANT[] rgvar);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shpropstgwritemultiple
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHPropStgWriteMultiple([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPropertyStorage>))] IPropertyStorage pps, nint /* optional uint* */ puCodePage, uint cpspec, [In][MarshalUsing(CountElementName = nameof(cpspec))] PROPSPEC[] rgpspec, [In][Out][MarshalUsing(CountElementName = nameof(cpspec))] PROPVARIANT[] rgvar, uint propidNameFirst);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shqueryinfokeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHQueryInfoKeyA(HKEY hkey, nint /* optional uint* */ pcSubKeys, nint /* optional uint* */ pcchMaxSubKeyLen, nint /* optional uint* */ pcValues, nint /* optional uint* */ pcchMaxValueNameLen);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shqueryinfokeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHQueryInfoKeyW(HKEY hkey, nint /* optional uint* */ pcSubKeys, nint /* optional uint* */ pcchMaxSubKeyLen, nint /* optional uint* */ pcValues, nint /* optional uint* */ pcchMaxValueNameLen);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shqueryrecyclebina
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHQueryRecycleBinA(PSTR pszRootPath, ref SHQUERYRBINFO pSHQueryRBInfo);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shqueryrecyclebinw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHQueryRecycleBinW(PWSTR pszRootPath, ref SHQUERYRBINFO pSHQueryRBInfo);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shqueryusernotificationstate
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHQueryUserNotificationState(out QUERY_USER_NOTIFICATION_STATE pquns);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shqueryvalueexa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHQueryValueExA(HKEY hkey, PSTR pszValue, nint /* optional uint* */ pdwReserved, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shqueryvalueexw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHQueryValueExW(HKEY hkey, PWSTR pszValue, nint /* optional uint* */ pdwReserved, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregcloseuskey
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegCloseUSKey(nint hUSKey);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregcreateuskeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegCreateUSKeyA(PSTR pszPath, uint samDesired, nint hRelativeUSKey, out nint phNewUSKey, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregcreateuskeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegCreateUSKeyW(PWSTR pwzPath, uint samDesired, nint hRelativeUSKey, out nint phNewUSKey, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregdeleteemptyuskeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegDeleteEmptyUSKeyA(nint hUSKey, PSTR pszSubKey, SHREGDEL_FLAGS delRegFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregdeleteemptyuskeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegDeleteEmptyUSKeyW(nint hUSKey, PWSTR pwzSubKey, SHREGDEL_FLAGS delRegFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregdeleteusvaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegDeleteUSValueA(nint hUSKey, PSTR pszValue, SHREGDEL_FLAGS delRegFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregdeleteusvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegDeleteUSValueW(nint hUSKey, PWSTR pwzValue, SHREGDEL_FLAGS delRegFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregduplicatehkey
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial HKEY SHRegDuplicateHKey(HKEY hkey);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregenumuskeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegEnumUSKeyA(nint hUSKey, uint dwIndex, [MarshalUsing(CountElementName = nameof(pcchName))] PSTR pszName, ref uint pcchName, SHREGENUM_FLAGS enumRegFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregenumuskeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegEnumUSKeyW(nint hUSKey, uint dwIndex, [MarshalUsing(CountElementName = nameof(pcchName))] PWSTR pwzName, ref uint pcchName, SHREGENUM_FLAGS enumRegFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregenumusvaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegEnumUSValueA(nint hUSkey, uint dwIndex, [MarshalUsing(CountElementName = nameof(pcchValueName))] PSTR pszValueName, ref uint pcchValueName, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData, SHREGENUM_FLAGS enumRegFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregenumusvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegEnumUSValueW(nint hUSkey, uint dwIndex, [MarshalUsing(CountElementName = nameof(pcchValueName))] PWSTR pszValueName, ref uint pcchValueName, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData, SHREGENUM_FLAGS enumRegFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreggetboolusvaluea
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHRegGetBoolUSValueA(PSTR pszSubKey, PSTR pszValue, BOOL fIgnoreHKCU, BOOL fDefault);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreggetboolusvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHRegGetBoolUSValueW(PWSTR pszSubKey, PWSTR pszValue, BOOL fIgnoreHKCU, BOOL fDefault);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreggetintw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHRegGetIntW(HKEY hk, PWSTR pwzKey, int iDefault);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreggetpatha
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegGetPathA(HKEY hKey, PSTR pcszSubKey, PSTR pcszValue, [MarshalUsing(ConstantElementCount = 260)] PSTR pszPath, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreggetpathw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegGetPathW(HKEY hKey, PWSTR pcszSubKey, PWSTR pcszValue, [MarshalUsing(ConstantElementCount = 260)] PWSTR pszPath, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreggetusvaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegGetUSValueA(PSTR pszSubKey, PSTR pszValue, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData, BOOL fIgnoreHKCU, nint /* optional void* */ pvDefaultData, uint dwDefaultDataSize);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreggetusvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegGetUSValueW(PWSTR pszSubKey, PWSTR pszValue, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData, BOOL fIgnoreHKCU, nint /* optional void* */ pvDefaultData, uint dwDefaultDataSize);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreggetvaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegGetValueA(HKEY hkey, PSTR pszSubKey, PSTR pszValue, int srrfFlags, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreggetvaluefromhkcuhklm
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegGetValueFromHKCUHKLM(PWSTR pwszKey, PWSTR pwszValue, int srrfFlags, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreggetvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegGetValueW(HKEY hkey, PWSTR pszSubKey, PWSTR pszValue, int srrfFlags, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregopenuskeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegOpenUSKeyA(PSTR pszPath, uint samDesired, nint hRelativeUSKey, out nint phNewUSKey, BOOL fIgnoreHKCU);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregopenuskeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegOpenUSKeyW(PWSTR pwzPath, uint samDesired, nint hRelativeUSKey, out nint phNewUSKey, BOOL fIgnoreHKCU);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregqueryinfouskeya
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegQueryInfoUSKeyA(nint hUSKey, nint /* optional uint* */ pcSubKeys, nint /* optional uint* */ pcchMaxSubKeyLen, nint /* optional uint* */ pcValues, nint /* optional uint* */ pcchMaxValueNameLen, SHREGENUM_FLAGS enumRegFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregqueryinfouskeyw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegQueryInfoUSKeyW(nint hUSKey, nint /* optional uint* */ pcSubKeys, nint /* optional uint* */ pcchMaxSubKeyLen, nint /* optional uint* */ pcValues, nint /* optional uint* */ pcchMaxValueNameLen, SHREGENUM_FLAGS enumRegFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregqueryusvaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegQueryUSValueA(nint hUSKey, PSTR pszValue, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData, BOOL fIgnoreHKCU, nint /* optional void* */ pvDefaultData, uint dwDefaultDataSize);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregqueryusvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegQueryUSValueW(nint hUSKey, PWSTR pszValue, nint /* optional uint* */ pdwType, nint /* optional void* */ pvData, nint /* optional uint* */ pcbData, BOOL fIgnoreHKCU, nint /* optional void* */ pvDefaultData, uint dwDefaultDataSize);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregsetpatha
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegSetPathA(HKEY hKey, PSTR pcszSubKey, PSTR pcszValue, PSTR pcszPath, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregsetpathw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegSetPathW(HKEY hKey, PWSTR pcszSubKey, PWSTR pcszValue, PWSTR pcszPath, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregsetusvaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegSetUSValueA(PSTR pszSubKey, PSTR pszValue, uint dwType, nint /* optional void* */ pvData, uint cbData, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregsetusvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegSetUSValueW(PWSTR pwzSubKey, PWSTR pwzValue, uint dwType, nint /* optional void* */ pvData, uint cbData, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregwriteusvaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegWriteUSValueA(nint hUSKey, PSTR pszValue, uint dwType, nint pvData, uint cbData, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shregwriteusvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR SHRegWriteUSValueW(nint hUSKey, PWSTR pwzValue, uint dwType, nint pvData, uint cbData, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shreleasethreadref
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHReleaseThreadRef();
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shremovelocalizedname
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHRemoveLocalizedName(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shreplacefrompropsheetextarray
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint SHReplaceFromPropSheetExtArray(HPSXA hpsxa, uint uPageID, LPFNSVADDPROPSHEETPAGE lpfnReplaceWith, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shresolvelibrary
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT SHResolveLibrary([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiLibrary);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shrestricted
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint SHRestricted(RESTRICTIONS rest);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shsendmessagebroadcasta
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial LRESULT SHSendMessageBroadcastA(uint uMsg, WPARAM wParam, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shsendmessagebroadcastw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial LRESULT SHSendMessageBroadcastW(uint uMsg, WPARAM wParam, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl/nf-shobjidl-shsetdefaultproperties
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHSetDefaultProperties(HWND hwnd, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, uint dwFileOpFlags, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<IFileOperationProgressSink?>))] IFileOperationProgressSink? pfops);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shsetfolderpatha
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHSetFolderPathA(int csidl, HANDLE hToken, uint dwFlags, PSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shsetfolderpathw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHSetFolderPathW(int csidl, HANDLE hToken, uint dwFlags, PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shsetinstanceexplorer
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial void SHSetInstanceExplorer(nint punk);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shsetknownfolderpath
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHSetKnownFolderPath(in Guid rfid, uint dwFlags, HANDLE hToken, PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shsetlocalizedname
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHSetLocalizedName(PWSTR pszPath, PWSTR pszResModule, int idsRes);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shsettemporarypropertyforitem
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT SHSetTemporaryPropertyForItem([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psi, in PROPERTYKEY propkey, in PROPVARIANT propvar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shsetthreadref
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHSetThreadRef(nint punk);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shsetunreadmailcountw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHSetUnreadMailCountW(PWSTR pszMailAddress, uint dwCount, PWSTR pszShellExecuteCommand);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shsetvaluea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHSetValueA(HKEY hkey, PSTR pszSubKey, PSTR pszValue, uint dwType, nint /* optional void* */ pvData, uint cbData);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shsetvaluew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHSetValueW(HKEY hkey, PWSTR pszSubKey, PWSTR pszValue, uint dwType, nint /* optional void* */ pvData, uint cbData);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shshellfolderview_message
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial LRESULT SHShellFolderView_Message(HWND hwndMain, uint uMsg, LPARAM lParam);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shshowmanagelibraryui
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT SHShowManageLibraryUI([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IShellItem>))] IShellItem psiLibrary, HWND hwndOwner, PWSTR pszTitle, PWSTR pszInstruction, LIBRARYMANAGEDIALOGOPTIONS lmdOptions);
    
    // https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-shsimpleidlistfrompath
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial nint SHSimpleIDListFromPath(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shskipjunction
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHSkipJunction([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IBindCtx?>))] IBindCtx? pbc, in Guid pclsid);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shstartnetconnectiondialogw
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT SHStartNetConnectionDialogW(HWND hwnd, PWSTR pszRemoteName, uint dwType);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shstrdupa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHStrDupA(PSTR psz, out PWSTR ppwsz);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shstrdupw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT SHStrDupW(PWSTR psz, out PWSTR ppwsz);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shstripmneumonica
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial CHAR SHStripMneumonicA(PSTR pszMenu);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shstripmneumonicw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial char SHStripMneumonicW(PWSTR pszMenu);
    
    // https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shtesttokenmembership
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHTestTokenMembership(HANDLE hToken, uint ulRID);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shunicodetoansi
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHUnicodeToAnsi(PWSTR pwszSrc, [MarshalUsing(CountElementName = nameof(cchBuf))] PSTR pszDst, int cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shunicodetounicode
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int SHUnicodeToUnicode(PWSTR pwzSrc, [MarshalUsing(CountElementName = nameof(cwchBuf))] PWSTR pwzDst, int cwchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-shunlockshared
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHUnlockShared(nint pvData);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shupdateimagea
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void SHUpdateImageA(PSTR pszHashItem, int iIndex, uint uFlags, int iImageIndex);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shupdateimagew
    [LibraryImport("SHELL32", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial void SHUpdateImageW(PWSTR pszHashItem, int iIndex, uint uFlags, int iImageIndex);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shvalidateunc
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SHValidateUNC(HWND hwndOwner, PWSTR pszFile, uint fConnect);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-signalfileopen
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL SignalFileOpen(nint pidl);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj/nf-shlobj-softwareupdatemessagebox
    [LibraryImport("SHDOCVW")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial uint SoftwareUpdateMessageBox(HWND hWnd, PWSTR pszDistUnit, uint dwFlags, nint /* optional SOFTDISTINFO* */ psdi);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-stgmakeuniquename
    [LibraryImport("SHELL32")]
    [SupportedOSPlatform("windows6.1")]
    [PreserveSig]
    public static partial HRESULT StgMakeUniqueName([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IStorage>))] IStorage pstgParent, PWSTR pszFileSpec, uint grfMode, in Guid riid, out nint /* void */ ppv);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcatbuffa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrCatBuffA([MarshalUsing(CountElementName = nameof(cchDestBuffSize))] PSTR pszDest, PSTR pszSrc, int cchDestBuffSize);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcatbuffw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrCatBuffW([MarshalUsing(CountElementName = nameof(cchDestBuffSize))] PWSTR pszDest, PWSTR pszSrc, int cchDestBuffSize);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcatchainw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint StrCatChainW([MarshalUsing(CountElementName = nameof(cchDst))] PWSTR pszDst, uint cchDst, uint ichAt, PWSTR pszSrc);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcatw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrCatW(PWSTR psz1, PWSTR psz2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strchra
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrChrA(PSTR pszStart, ushort wMatch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strchria
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrChrIA(PSTR pszStart, ushort wMatch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strchriw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrChrIW(PWSTR pszStart, char wMatch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strchrniw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrChrNIW(PWSTR pszStart, char wMatch, uint cchMax);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strchrnw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrChrNW(PWSTR pszStart, char wMatch, uint cchMax);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strchrw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrChrW(PWSTR pszStart, char wMatch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpca
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpCA(PSTR pszStr1, PSTR pszStr2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpcw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpCW(PWSTR pszStr1, PWSTR pszStr2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpica
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpICA(PSTR pszStr1, PSTR pszStr2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpicw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpICW(PWSTR pszStr1, PWSTR pszStr2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpiw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpIW(PWSTR psz1, PWSTR psz2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmplogicalw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial int StrCmpLogicalW(PWSTR psz1, PWSTR psz2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpna
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpNA(PSTR psz1, PSTR psz2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpnca
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpNCA(PSTR pszStr1, PSTR pszStr2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpncw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpNCW(PWSTR pszStr1, PWSTR pszStr2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpnia
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpNIA(PSTR psz1, PSTR psz2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpnica
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpNICA(PSTR pszStr1, PSTR pszStr2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpnicw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpNICW(PWSTR pszStr1, PWSTR pszStr2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpniw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpNIW(PWSTR psz1, PWSTR psz2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpnw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpNW(PWSTR psz1, PWSTR psz2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcmpw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCmpW(PWSTR psz1, PWSTR psz2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcpynw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrCpyNW([MarshalUsing(CountElementName = nameof(cchMax))] PWSTR pszDst, PWSTR pszSrc, int cchMax);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcpyw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrCpyW(PWSTR psz1, PWSTR psz2);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcspna
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCSpnA(PSTR pszStr, PSTR pszSet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcspnia
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCSpnIA(PSTR pszStr, PSTR pszSet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcspniw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCSpnIW(PWSTR pszStr, PWSTR pszSet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strcspnw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrCSpnW(PWSTR pszStr, PWSTR pszSet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strdupa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrDupA(PSTR pszSrch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strdupw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrDupW(PWSTR pszSrch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strformatbytesize64a
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrFormatByteSize64A(long qdw, [MarshalUsing(CountElementName = nameof(cchBuf))] PSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strformatbytesizea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrFormatByteSizeA(uint dw, [MarshalUsing(CountElementName = nameof(cchBuf))] PSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strformatbytesizeex
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    public static partial HRESULT StrFormatByteSizeEx(ulong ull, SFBS_FLAGS flags, [MarshalUsing(CountElementName = nameof(cchBuf))] PWSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strformatbytesizew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrFormatByteSizeW(long qdw, [MarshalUsing(CountElementName = nameof(cchBuf))] PWSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strformatkbsizea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrFormatKBSizeA(long qdw, [MarshalUsing(CountElementName = nameof(cchBuf))] PSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strformatkbsizew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrFormatKBSizeW(long qdw, [MarshalUsing(CountElementName = nameof(cchBuf))] PWSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strfromtimeintervala
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrFromTimeIntervalA([MarshalUsing(CountElementName = nameof(cchMax))] PSTR pszOut, uint cchMax, uint dwTimeMS, int digits);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strfromtimeintervalw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrFromTimeIntervalW([MarshalUsing(CountElementName = nameof(cchMax))] PWSTR pszOut, uint cchMax, uint dwTimeMS, int digits);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strisintlequala
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL StrIsIntlEqualA(BOOL fCaseSens, PSTR pszString1, PSTR pszString2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strisintlequalw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL StrIsIntlEqualW(BOOL fCaseSens, PWSTR pszString1, PWSTR pszString2, int nChar);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strncata
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrNCatA([MarshalUsing(CountElementName = nameof(cchMax))] PSTR psz1, PSTR psz2, int cchMax);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strncatw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrNCatW([MarshalUsing(CountElementName = nameof(cchMax))] PWSTR psz1, PWSTR psz2, int cchMax);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strpbrka
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrPBrkA(PSTR psz, PSTR pszSet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strpbrkw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrPBrkW(PWSTR psz, PWSTR pszSet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrchra
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrRChrA(PSTR pszStart, PSTR pszEnd, ushort wMatch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrchria
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrRChrIA(PSTR pszStart, PSTR pszEnd, ushort wMatch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrchriw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrRChrIW(PWSTR pszStart, PWSTR pszEnd, char wMatch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrchrw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrRChrW(PWSTR pszStart, PWSTR pszEnd, char wMatch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrettobstr
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT StrRetToBSTR(ref STRRET pstr, nint /* optional nint* */ pidl, out BSTR pbstr);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrettobufa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT StrRetToBufA(ref STRRET pstr, nint /* optional nint* */ pidl, [MarshalUsing(CountElementName = nameof(cchBuf))] PSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrettobufw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT StrRetToBufW(ref STRRET pstr, nint /* optional nint* */ pidl, [MarshalUsing(CountElementName = nameof(cchBuf))] PWSTR pszBuf, uint cchBuf);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrettostra
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT StrRetToStrA(ref STRRET pstr, nint /* optional nint* */ pidl, out PSTR ppsz);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrettostrw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT StrRetToStrW(ref STRRET pstr, nint /* optional nint* */ pidl, out PWSTR ppsz);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrstria
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrRStrIA(PSTR pszSource, PSTR pszLast, PSTR pszSrch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strrstriw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrRStrIW(PWSTR pszSource, PWSTR pszLast, PWSTR pszSrch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strspna
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrSpnA(PSTR psz, PSTR pszSet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strspnw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrSpnW(PWSTR psz, PWSTR pszSet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strstra
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrStrA(PSTR pszFirst, PSTR pszSrch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strstria
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR StrStrIA(PSTR pszFirst, PSTR pszSrch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strstriw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrStrIW(PWSTR pszFirst, PWSTR pszSrch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strstrniw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrStrNIW(PWSTR pszFirst, PWSTR pszSrch, uint cchMax);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strstrnw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows6.0.6000")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrStrNW(PWSTR pszFirst, PWSTR pszSrch, uint cchMax);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strstrw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR StrStrW(PWSTR pszFirst, PWSTR pszSrch);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strtoint64exa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL StrToInt64ExA(PSTR pszString, int dwFlags, out long pllRet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strtoint64exw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL StrToInt64ExW(PWSTR pszString, int dwFlags, out long pllRet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strtointa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrToIntA(PSTR pszSrc);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strtointexa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL StrToIntExA(PSTR pszString, int dwFlags, out int piRet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strtointexw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL StrToIntExW(PWSTR pszString, int dwFlags, out int piRet);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strtointw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int StrToIntW(PWSTR pszSrc);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strtrima
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL StrTrimA(PSTR psz, PSTR pszTrimChars);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-strtrimw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL StrTrimW(PWSTR psz, PWSTR pszTrimChars);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-trackpopupmenu
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial int TrackPopupMenu(HMENU hMenu, TRACK_POPUP_MENU_FLAGS uFlags, int x, int y, int nReserved, HWND hWnd, nint /* optional RECT* */ prcRect);
    
    // https://learn.microsoft.com/windows/win32/api/userenv/nf-userenv-unloaduserprofile
    [LibraryImport("USERENV", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL UnloadUserProfile(HANDLE hToken, HANDLE hProfile);
    
    [LibraryImport("api-ms-win-core-psm-appnotify-l1-1-1")]
    [PreserveSig]
    public static partial void UnregisterAppConstrainedChangeNotification(PAPPCONSTRAIN_REGISTRATION Registration);
    
    // https://learn.microsoft.com/windows/win32/api/appnotify/nf-appnotify-unregisterappstatechangenotification
    [LibraryImport("api-ms-win-core-psm-appnotify-l1-1-0")]
    [PreserveSig]
    public static partial void UnregisterAppStateChangeNotification(PAPPSTATE_REGISTRATION Registration);
    
    // https://learn.microsoft.com/windows/win32/api/shellscalingapi/nf-shellscalingapi-unregisterscalechangeevent
    [LibraryImport("api-ms-win-shcore-scaling-l1-1-1")]
    [SupportedOSPlatform("windows8.1")]
    [PreserveSig]
    public static partial HRESULT UnregisterScaleChangeEvent(nuint dwCookie);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlapplyschemea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlApplySchemeA(PSTR pszIn, [MarshalUsing(CountElementName = nameof(pcchOut))] PSTR pszOut, ref uint pcchOut, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlapplyschemew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlApplySchemeW(PWSTR pszIn, [MarshalUsing(CountElementName = nameof(pcchOut))] PWSTR pszOut, ref uint pcchOut, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlcanonicalizea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlCanonicalizeA(PSTR pszUrl, [MarshalUsing(CountElementName = nameof(pcchCanonicalized))] PSTR pszCanonicalized, ref uint pcchCanonicalized, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlcanonicalizew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlCanonicalizeW(PWSTR pszUrl, [MarshalUsing(CountElementName = nameof(pcchCanonicalized))] PWSTR pszCanonicalized, ref uint pcchCanonicalized, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlcombinea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlCombineA(PSTR pszBase, PSTR pszRelative, [MarshalUsing(CountElementName = nameof(pcchCombined))] PSTR pszCombined, ref uint pcchCombined, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlcombinew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlCombineW(PWSTR pszBase, PWSTR pszRelative, [MarshalUsing(CountElementName = nameof(pcchCombined))] PWSTR pszCombined, ref uint pcchCombined, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlcomparea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int UrlCompareA(PSTR psz1, PSTR psz2, BOOL fIgnoreSlash);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlcomparew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int UrlCompareW(PWSTR psz1, PWSTR psz2, BOOL fIgnoreSlash);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlcreatefrompatha
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlCreateFromPathA(PSTR pszPath, [MarshalUsing(CountElementName = nameof(pcchUrl))] PSTR pszUrl, ref uint pcchUrl, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlcreatefrompathw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlCreateFromPathW(PWSTR pszPath, [MarshalUsing(CountElementName = nameof(pcchUrl))] PWSTR pszUrl, ref uint pcchUrl, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlescapea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlEscapeA(PSTR pszUrl, [MarshalUsing(CountElementName = nameof(pcchEscaped))] PSTR pszEscaped, ref uint pcchEscaped, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlescapew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlEscapeW(PWSTR pszUrl, [MarshalUsing(CountElementName = nameof(pcchEscaped))] PWSTR pszEscaped, ref uint pcchEscaped, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlfixupw
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT UrlFixupW(PWSTR pcszUrl, [MarshalUsing(CountElementName = nameof(cchMax))] PWSTR pszTranslatedUrl, uint cchMax);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlgetlocationa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PSTR UrlGetLocationA(PSTR pszURL);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlgetlocationw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial PWSTR UrlGetLocationW(PWSTR pszURL);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlgetparta
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlGetPartA(PSTR pszIn, [MarshalUsing(CountElementName = nameof(pcchOut))] PSTR pszOut, ref uint pcchOut, uint dwPart, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlgetpartw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlGetPartW(PWSTR pszIn, [MarshalUsing(CountElementName = nameof(pcchOut))] PWSTR pszOut, ref uint pcchOut, uint dwPart, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlhasha
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlHashA(PSTR pszUrl, nint /* byte array */ pbHash, uint cbHash);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlhashw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlHashW(PWSTR pszUrl, nint /* byte array */ pbHash, uint cbHash);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlisa
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL UrlIsA(PSTR pszUrl, URLIS UrlIs);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlisnohistorya
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL UrlIsNoHistoryA(PSTR pszURL);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlisnohistoryw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL UrlIsNoHistoryW(PWSTR pszURL);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlisopaquea
    [LibraryImport("SHLWAPI", SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL UrlIsOpaqueA(PSTR pszURL);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlisopaquew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL UrlIsOpaqueW(PWSTR pszURL);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlisw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL UrlIsW(PWSTR pszUrl, URLIS UrlIs);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlunescapea
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlUnescapeA(PSTR pszUrl, [MarshalUsing(CountElementName = nameof(pcchUnescaped))] PSTR pszUnescaped, nint /* optional uint* */ pcchUnescaped, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-urlunescapew
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial HRESULT UrlUnescapeW(PWSTR pszUrl, [MarshalUsing(CountElementName = nameof(pcchUnescaped))] PWSTR pszUnescaped, nint /* optional uint* */ pcchUnescaped, uint dwFlags);
    
    // https://learn.microsoft.com/windows/win32/api/propvarutil/nf-propvarutil-varianttostrret
    [LibraryImport("PROPSYS")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial HRESULT VariantToStrRet(in VARIANT varIn, out STRRET pstrret);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-whichplatform
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    public static partial uint WhichPlatform();
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-win32deletefile
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL Win32DeleteFile(PWSTR pszPath);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-winhelpa
    [LibraryImport("USER32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL WinHelpA(HWND hWndMain, PSTR lpszHelp, uint uCommand, nuint dwData);
    
    // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-winhelpw
    [LibraryImport("USER32", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL WinHelpW(HWND hWndMain, PWSTR lpszHelp, uint uCommand, nuint dwData);
    
    // https://learn.microsoft.com/windows/win32/api/winnetwk/nf-winnetwk-wnetenumresourcew
    [LibraryImport("MPR", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR WNetEnumResourceW(HANDLE hEnum, ref uint lpcCount, nint lpBuffer, ref uint lpBufferSize);
    
    // https://learn.microsoft.com/windows/win32/api/winnetwk/nf-winnetwk-wnetgetnetworkinformationw
    [LibraryImport("MPR", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR WNetGetNetworkInformationW(PWSTR lpProvider, out NETINFOSTRUCT lpNetInfoStruct);
    
    // https://learn.microsoft.com/windows/win32/api/winnetwk/nf-winnetwk-wnetopenenumw
    [LibraryImport("MPR", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial WIN32_ERROR WNetOpenEnumW(NET_RESOURCE_SCOPE dwScope, NET_RESOURCE_TYPE dwType, WNET_OPEN_ENUM_USAGE dwUsage, nint /* optional NETRESOURCEW* */ lpNetResource, out HANDLE lphEnum);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-wnsprintfa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int wnsprintfA([MarshalUsing(CountElementName = nameof(cchDest))] PSTR pszDest, int cchDest, PSTR pszFmt);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-wnsprintfw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int wnsprintfW([MarshalUsing(CountElementName = nameof(cchDest))] PWSTR pszDest, int cchDest, PWSTR pszFmt);
    
    // https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-writecabinetstate
    [LibraryImport("SHELL32", SetLastError = true)]
    [SupportedOSPlatform("windows5.1.2600")]
    [PreserveSig]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    public static partial BOOL WriteCabinetState(in CABINETSTATE pcs);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-wvnsprintfa
    [LibraryImport("SHLWAPI")]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int wvnsprintfA([MarshalUsing(CountElementName = nameof(cchDest))] PSTR pszDest, int cchDest, PSTR pszFmt, in sbyte arglist);
    
    // https://learn.microsoft.com/windows/win32/api/shlwapi/nf-shlwapi-wvnsprintfw
    [LibraryImport("SHLWAPI", StringMarshalling = StringMarshalling.Utf16)]
    [SupportedOSPlatform("windows5.0")]
    [PreserveSig]
    public static partial int wvnsprintfW([MarshalUsing(CountElementName = nameof(cchDest))] PWSTR pszDest, int cchDest, PWSTR pszFmt, in sbyte arglist);
}
