#nullable enable
namespace ShellN;

public static partial class Constants
{
    public const uint ABE_BOTTOM = 3;
    
    public const uint ABE_LEFT = 0;
    
    public const uint ABE_RIGHT = 2;
    
    public const uint ABE_TOP = 1;
    
    public const uint ABM_ACTIVATE = 6;
    
    public const uint ABM_GETAUTOHIDEBAR = 7;
    
    public const uint ABM_GETAUTOHIDEBAREX = 11;
    
    public const uint ABM_GETSTATE = 4;
    
    public const uint ABM_GETTASKBARPOS = 5;
    
    public const uint ABM_NEW = 0;
    
    public const uint ABM_QUERYPOS = 2;
    
    public const uint ABM_REMOVE = 1;
    
    public const uint ABM_SETAUTOHIDEBAR = 8;
    
    public const uint ABM_SETAUTOHIDEBAREX = 12;
    
    public const uint ABM_SETPOS = 3;
    
    public const uint ABM_SETSTATE = 10;
    
    public const uint ABM_WINDOWPOSCHANGED = 9;
    
    public const uint ABN_FULLSCREENAPP = 2;
    
    public const uint ABN_POSCHANGED = 1;
    
    public const uint ABN_STATECHANGE = 0;
    
    public const uint ABN_WINDOWARRANGE = 3;
    
    public const uint ABS_ALWAYSONTOP = 2;
    
    public const uint ABS_AUTOHIDE = 1;
    
    public static Guid AccessibilityDockingService => new(0x29ce1d46, 0xb481, 0x4aa0, 0xa0, 0x8a, 0xd3, 0xeb, 0xc8, 0xac, 0xa4, 0x02);
    
    public const uint ACDD_VISIBLE = 1;
    
    public const uint AD_APPLY_BUFFERED_REFRESH = 16;
    
    public const uint AD_APPLY_DYNAMICREFRESH = 32;
    
    public const uint AD_APPLY_FORCE = 8;
    
    public const uint AD_APPLY_HTMLGEN = 2;
    
    public const uint AD_APPLY_REFRESH = 4;
    
    public const uint AD_APPLY_SAVE = 1;
    
    public const uint AD_GETWP_BMP = 0;
    
    public const uint AD_GETWP_IMAGE = 1;
    
    public const uint AD_GETWP_LAST_APPLIED = 2;
    
    public const uint ADDURL_SILENT = 1;
    
    public static Guid AlphabeticalCategorizer => new(0x3c2654c6, 0x7372, 0x4f6b, 0xb3, 0x10, 0x55, 0xd6, 0x12, 0x8f, 0x49, 0xd2);
    
    public static Guid ApplicationActivationManager => new(0x45ba127d, 0x10a8, 0x46ea, 0x8a, 0xb7, 0x56, 0xea, 0x90, 0x78, 0x94, 0x3c);
    
    public static Guid ApplicationAssociationRegistration => new(0x591209c7, 0x767b, 0x42b2, 0x9f, 0xba, 0x44, 0xee, 0x46, 0x15, 0xf2, 0xc7);
    
    public static Guid ApplicationAssociationRegistrationUI => new(0x1968106d, 0xf3b5, 0x44cf, 0x89, 0x0e, 0x11, 0x6f, 0xcb, 0x9e, 0xce, 0xf1);
    
    public static Guid ApplicationDesignModeSettings => new(0x958a6fb5, 0xdcb2, 0x4faf, 0xaa, 0xfd, 0x7f, 0xb0, 0x54, 0xad, 0x1a, 0x3b);
    
    public static Guid ApplicationDestinations => new(0x86c14003, 0x4d6b, 0x4ef3, 0xa7, 0xb4, 0x05, 0x06, 0x66, 0x3b, 0x2e, 0x68);
    
    public static Guid ApplicationDocumentLists => new(0x86bec222, 0x30f2, 0x47e0, 0x9f, 0x25, 0x60, 0xd1, 0x1c, 0xd7, 0x5c, 0x28);
    
    public const uint APPNAMEBUFFERLEN = 40;
    
    public static Guid AppShellVerbHandler => new(0x4ed3a719, 0xcea8, 0x4bd9, 0x91, 0x0d, 0xe2, 0x52, 0xf9, 0x97, 0xaf, 0xc2);
    
    public static Guid AppStartupLink => new(0x273eb5e7, 0x88b0, 0x4843, 0xbf, 0xef, 0xe2, 0xc8, 0x1d, 0x43, 0xaa, 0xe5);
    
    public static Guid AppVisibility => new(0x7e5fe3d9, 0x985f, 0x4908, 0x91, 0xf9, 0xee, 0x19, 0xf9, 0xfd, 0x15, 0x14);
    
    public const uint ARCONTENT_AUDIOCD = 4;
    
    public const uint ARCONTENT_AUTOPLAYMUSIC = 256;
    
    public const uint ARCONTENT_AUTOPLAYPIX = 128;
    
    public const uint ARCONTENT_AUTOPLAYVIDEO = 512;
    
    public const uint ARCONTENT_AUTORUNINF = 2;
    
    public const uint ARCONTENT_BLANKBD = 8192;
    
    public const uint ARCONTENT_BLANKCD = 16;
    
    public const uint ARCONTENT_BLANKDVD = 32;
    
    public const uint ARCONTENT_BLURAY = 16384;
    
    public const uint ARCONTENT_CAMERASTORAGE = 32768;
    
    public const uint ARCONTENT_CUSTOMEVENT = 65536;
    
    public const uint ARCONTENT_DVDAUDIO = 4096;
    
    public const uint ARCONTENT_DVDMOVIE = 8;
    
    public const uint ARCONTENT_MASK = 131070;
    
    public const uint ARCONTENT_NONE = 0;
    
    public const uint ARCONTENT_PHASE_FINAL = 1073741824;
    
    public const uint ARCONTENT_PHASE_MASK = 1879048192;
    
    public const uint ARCONTENT_PHASE_PRESNIFF = 268435456;
    
    public const uint ARCONTENT_PHASE_SNIFFING = 536870912;
    
    public const uint ARCONTENT_PHASE_UNKNOWN = 0;
    
    public const uint ARCONTENT_SVCD = 2048;
    
    public const uint ARCONTENT_UNKNOWNCONTENT = 64;
    
    public const uint ARCONTENT_VCD = 1024;
    
    public static Guid AttachmentServices => new(0x4125dd96, 0xe03a, 0x4103, 0x8f, 0x70, 0xe0, 0x59, 0x7d, 0x80, 0x3b, 0x9c);
    
    public const uint BFFM_ENABLEOK = 1125;
    
    public const uint BFFM_INITIALIZED = 1;
    
    public const uint BFFM_IUNKNOWN = 5;
    
    public const uint BFFM_SELCHANGED = 2;
    
    public const uint BFFM_SETEXPANDED = 1130;
    
    public const uint BFFM_SETOKTEXT = 1129;
    
    public const uint BFFM_SETSELECTION = 1127;
    
    public const uint BFFM_SETSELECTIONA = 1126;
    
    public const uint BFFM_SETSELECTIONW = 1127;
    
    public const uint BFFM_SETSTATUSTEXT = 1128;
    
    public const uint BFFM_SETSTATUSTEXTA = 1124;
    
    public const uint BFFM_SETSTATUSTEXTW = 1128;
    
    public const uint BFFM_VALIDATEFAILED = 4;
    
    public const uint BFFM_VALIDATEFAILEDA = 3;
    
    public const uint BFFM_VALIDATEFAILEDW = 4;
    
    public static Guid BHID_AssociationArray => new(0xbea9ef17, 0x82f1, 0x4f60, 0x92, 0x84, 0x4f, 0x8d, 0xb7, 0x5c, 0x3b, 0xe9);
    
    public static Guid BHID_DataObject => new(0xb8c0bd9f, 0xed24, 0x455c, 0x83, 0xe6, 0xd5, 0x39, 0x0c, 0x4f, 0xe8, 0xc4);
    
    public static Guid BHID_EnumAssocHandlers => new(0xb8ab0b9c, 0xc2ec, 0x4f7a, 0x91, 0x8d, 0x31, 0x49, 0x00, 0xe6, 0x28, 0x0a);
    
    public static Guid BHID_EnumItems => new(0x94f60519, 0x2850, 0x4924, 0xaa, 0x5a, 0xd1, 0x5e, 0x84, 0x86, 0x80, 0x39);
    
    public static Guid BHID_FilePlaceholder => new(0x8677dceb, 0xaae0, 0x4005, 0x8d, 0x3d, 0x54, 0x7f, 0xa8, 0x52, 0xf8, 0x25);
    
    public static Guid BHID_Filter => new(0x38d08778, 0xf557, 0x4690, 0x9e, 0xbf, 0xba, 0x54, 0x70, 0x6a, 0xd8, 0xf7);
    
    public static Guid BHID_LinkTargetItem => new(0x3981e228, 0xf559, 0x11d3, 0x8e, 0x3a, 0x00, 0xc0, 0x4f, 0x68, 0x37, 0xd5);
    
    public static Guid BHID_PropertyStore => new(0x0384e1a4, 0x1523, 0x439c, 0xa4, 0xc8, 0xab, 0x91, 0x10, 0x52, 0xf5, 0x86);
    
    public static Guid BHID_RandomAccessStream => new(0xf16fc93b, 0x77ae, 0x4cfe, 0xbd, 0xa7, 0xa8, 0x66, 0xee, 0xa6, 0x87, 0x8d);
    
    public static Guid BHID_SFObject => new(0x3981e224, 0xf559, 0x11d3, 0x8e, 0x3a, 0x00, 0xc0, 0x4f, 0x68, 0x37, 0xd5);
    
    public static Guid BHID_SFUIObject => new(0x3981e225, 0xf559, 0x11d3, 0x8e, 0x3a, 0x00, 0xc0, 0x4f, 0x68, 0x37, 0xd5);
    
    public static Guid BHID_SFViewObject => new(0x3981e226, 0xf559, 0x11d3, 0x8e, 0x3a, 0x00, 0xc0, 0x4f, 0x68, 0x37, 0xd5);
    
    public static Guid BHID_Storage => new(0x3981e227, 0xf559, 0x11d3, 0x8e, 0x3a, 0x00, 0xc0, 0x4f, 0x68, 0x37, 0xd5);
    
    public static Guid BHID_StorageEnum => new(0x4621a4e3, 0xf0d6, 0x4773, 0x8a, 0x9c, 0x46, 0xe7, 0x7b, 0x17, 0x48, 0x40);
    
    public static Guid BHID_StorageItem => new(0x404e2109, 0x77d2, 0x4699, 0xa5, 0xa0, 0x4f, 0xdf, 0x10, 0xdb, 0x98, 0x37);
    
    public static Guid BHID_Stream => new(0x1cebb3ab, 0x7c10, 0x499a, 0xa4, 0x17, 0x92, 0xca, 0x16, 0xc4, 0xcb, 0x83);
    
    public static Guid BHID_ThumbnailHandler => new(0x7b2e650a, 0x8e20, 0x4f4a, 0xb0, 0x9e, 0x65, 0x97, 0xaf, 0xc7, 0x2f, 0xb0);
    
    public static Guid BHID_Transfer => new(0xd5e346a1, 0xf753, 0x4932, 0xb4, 0x03, 0x45, 0x74, 0x80, 0x0e, 0x24, 0x98);
    
    public const uint BIF_BROWSEFILEJUNCTIONS = 65536;
    
    public const uint BIF_BROWSEFORCOMPUTER = 4096;
    
    public const uint BIF_BROWSEFORPRINTER = 8192;
    
    public const uint BIF_BROWSEINCLUDEFILES = 16384;
    
    public const uint BIF_BROWSEINCLUDEURLS = 128;
    
    public const uint BIF_DONTGOBELOWDOMAIN = 2;
    
    public const uint BIF_EDITBOX = 16;
    
    public const uint BIF_NEWDIALOGSTYLE = 64;
    
    public const uint BIF_NONEWFOLDERBUTTON = 512;
    
    public const uint BIF_NOTRANSLATETARGETS = 1024;
    
    public const uint BIF_RETURNFSANCESTORS = 8;
    
    public const uint BIF_RETURNONLYFSDIRS = 1;
    
    public const uint BIF_SHAREABLE = 32768;
    
    public const uint BIF_STATUSTEXT = 4;
    
    public const uint BIF_UAHINT = 256;
    
    public const uint BIF_VALIDATE = 32;
    
    public const uint BIND_INTERRUPTABLE = uint.MaxValue;
    
    public const int BMICON_LARGE = 0;
    
    public const int BMICON_SMALL = 1;
    
    public const uint BSF_CANMAXIMIZE = 1024;
    
    public const uint BSF_DELEGATEDNAVIGATION = 65536;
    
    public const uint BSF_DONTSHOWNAVCANCELPAGE = 16384;
    
    public const uint BSF_FEEDNAVIGATION = 524288;
    
    public const uint BSF_FEEDSUBSCRIBED = 1048576;
    
    public const uint BSF_HTMLNAVCANCELED = 8192;
    
    public const uint BSF_MERGEDMENUS = 262144;
    
    public const uint BSF_NAVNOHISTORY = 4096;
    
    public const uint BSF_NOLOCALFILEWARNING = 16;
    
    public const uint BSF_REGISTERASDROPTARGET = 1;
    
    public const uint BSF_RESIZABLE = 512;
    
    public const uint BSF_SETNAVIGATABLECODEPAGE = 32768;
    
    public const uint BSF_THEATERMODE = 2;
    
    public const uint BSF_TOPBROWSER = 2048;
    
    public const uint BSF_TRUSTEDFORACTIVEX = 131072;
    
    public const uint BSF_UISETBYAUTOMATION = 256;
    
    public const uint BSIM_STATE = 1;
    
    public const uint BSIM_STYLE = 2;
    
    public const uint BSIS_ALWAYSGRIPPER = 2;
    
    public const uint BSIS_AUTOGRIPPER = 0;
    
    public const uint BSIS_FIXEDORDER = 1024;
    
    public const uint BSIS_LEFTALIGN = 4;
    
    public const uint BSIS_LOCKED = 256;
    
    public const uint BSIS_NOCAPTION = 64;
    
    public const uint BSIS_NOCONTEXTMENU = 16;
    
    public const uint BSIS_NODROPTARGET = 32;
    
    public const uint BSIS_NOGRIPPER = 1;
    
    public const uint BSIS_PREFERNOLINEBREAK = 128;
    
    public const uint BSIS_PRESERVEORDERDURINGLAYOUT = 512;
    
    public const uint BSIS_SINGLECLICK = 8;
    
    public const uint BSSF_NOTITLE = 2;
    
    public const uint BSSF_UNDELETEABLE = 4096;
    
    public const uint BSSF_VISIBLE = 1;
    
    public const uint BUFFLEN = 255;
    
    public const uint CABINETSTATE_VERSION = 2;
    
    public static Guid CATID_BrowsableShellExt => new(0x00021490, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid CATID_BrowseInPlace => new(0x00021491, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid CATID_CommBand => new(0x00021494, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid CATID_DeskBand => new(0x00021492, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid CATID_FilePlaceholderMergeHandler => new(0x3e9c9a51, 0xd4aa, 0x4870, 0xb4, 0x7c, 0x74, 0x24, 0xb4, 0x91, 0xf1, 0xcc);
    
    public static Guid CATID_InfoBand => new(0x00021493, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid CATID_LocationFactory => new(0x965c4d51, 0x8b76, 0x4e57, 0x80, 0xb7, 0x56, 0x4d, 0x2e, 0xa4, 0xb5, 0x5e);
    
    public static Guid CATID_LocationProvider => new(0x1b3ca474, 0x2614, 0x414b, 0xb8, 0x13, 0x1a, 0xce, 0xca, 0x3e, 0x3d, 0xd8);
    
    public static Guid CATID_SearchableApplication => new(0x366c292a, 0xd9b3, 0x4dbf, 0xbb, 0x70, 0xe6, 0x2e, 0xc3, 0xd0, 0xbb, 0xbf);
    
    public const uint CDB2GVF_ADDSHIELD = 64;
    
    public const uint CDB2GVF_ALLOWPREVIEWPANE = 4;
    
    public const uint CDB2GVF_ISFILESAVE = 2;
    
    public const uint CDB2GVF_ISFOLDERPICKER = 32;
    
    public const uint CDB2GVF_NOINCLUDEITEM = 16;
    
    public const uint CDB2GVF_NOSELECTVERB = 8;
    
    public const uint CDB2GVF_SHOWALLFILES = 1;
    
    public const uint CDB2N_CONTEXTMENU_DONE = 1;
    
    public const uint CDB2N_CONTEXTMENU_START = 2;
    
    public const uint CDBOSC_KILLFOCUS = 1;
    
    public const uint CDBOSC_RENAME = 3;
    
    public const uint CDBOSC_SELCHANGE = 2;
    
    public const uint CDBOSC_SETFOCUS = 0;
    
    public const uint CDBOSC_STATECHANGE = 4;
    
    public static Guid CDBurn => new(0xfbeb8a05, 0xbeee, 0x4442, 0x80, 0x4e, 0x40, 0x9d, 0x6c, 0x45, 0x15, 0xe9);
    
    public const string CFSTR_AUTOPLAY_SHELLIDLISTS = @"Autoplay Enumerated IDList Array";
    
    public const string CFSTR_DROPDESCRIPTION = @"DropDescription";
    
    public const string CFSTR_FILE_ATTRIBUTES_ARRAY = @"File Attributes Array";
    
    public const string CFSTR_FILECONTENTS = @"FileContents";
    
    public const string CFSTR_FILEDESCRIPTOR = @"FileGroupDescriptorW";
    
    public const string CFSTR_FILEDESCRIPTORA = @"FileGroupDescriptor";
    
    public const string CFSTR_FILEDESCRIPTORW = @"FileGroupDescriptorW";
    
    public const string CFSTR_FILENAME = @"FileNameW";
    
    public const string CFSTR_FILENAMEA = @"FileName";
    
    public const string CFSTR_FILENAMEMAP = @"FileNameMapW";
    
    public const string CFSTR_FILENAMEMAPA = @"FileNameMap";
    
    public const string CFSTR_FILENAMEMAPW = @"FileNameMapW";
    
    public const string CFSTR_FILENAMEW = @"FileNameW";
    
    public const string CFSTR_INDRAGLOOP = @"InShellDragLoop";
    
    public const string CFSTR_INETURL = @"UniformResourceLocatorW";
    
    public const string CFSTR_INETURLA = @"UniformResourceLocator";
    
    public const string CFSTR_INETURLW = @"UniformResourceLocatorW";
    
    public const string CFSTR_INVOKECOMMAND_DROPPARAM = @"InvokeCommand DropParam";
    
    public const string CFSTR_LOGICALPERFORMEDDROPEFFECT = @"Logical Performed DropEffect";
    
    public const string CFSTR_MOUNTEDVOLUME = @"MountedVolume";
    
    public const string CFSTR_NETRESOURCES = @"Net Resource";
    
    public const string CFSTR_PASTESUCCEEDED = @"Paste Succeeded";
    
    public const string CFSTR_PERFORMEDDROPEFFECT = @"Performed DropEffect";
    
    public const string CFSTR_PERSISTEDDATAOBJECT = @"PersistedDataObject";
    
    public const string CFSTR_PREFERREDDROPEFFECT = @"Preferred DropEffect";
    
    public const string CFSTR_PRINTERGROUP = @"PrinterFriendlyName";
    
    public const string CFSTR_SHELLDROPHANDLER = @"DropHandlerCLSID";
    
    public const string CFSTR_SHELLIDLIST = @"Shell IDList Array";
    
    public const string CFSTR_SHELLIDLISTOFFSET = @"Shell Object Offsets";
    
    public const string CFSTR_SHELLURL = @"UniformResourceLocator";
    
    public const string CFSTR_TARGETCLSID = @"TargetCLSID";
    
    public const string CFSTR_UNTRUSTEDDRAGDROP = @"UntrustedDragDrop";
    
    public const string CFSTR_ZONEIDENTIFIER = @"ZoneIdentifier";
    
    public static Guid CGID_DefView => new(0x4af07f10, 0xd231, 0x11d0, 0xb9, 0x42, 0x00, 0xa0, 0xc9, 0x03, 0x12, 0xe1);
    
    public static Guid CGID_Explorer => new(0x000214d0, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid CGID_ExplorerBarDoc => new(0x000214d3, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid CGID_MENUDESKBAR => new(0x5c9f0a12, 0x959e, 0x11d0, 0xa3, 0xa4, 0x00, 0xa0, 0xc9, 0x08, 0x26, 0x36);
    
    public static Guid CGID_ShellDocView => new(0x000214d1, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid CGID_ShellServiceObject => new(0x000214d2, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid CGID_ShortCut => new(0x93a68750, 0x951a, 0x11d1, 0x94, 0x6f, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00);
    
    public const uint CLOSEPROPS_DISCARD = 1;
    
    public const uint CLOSEPROPS_NONE = 0;
    
    public static Guid CLSID_ACLCustomMRU => new(0x6935db93, 0x21e8, 0x4ccc, 0xbe, 0xb9, 0x9f, 0xe3, 0xc7, 0x7a, 0x29, 0x7a);
    
    public static Guid CLSID_ACLHistory => new(0x00bb2764, 0x6a77, 0x11d0, 0xa5, 0x35, 0x00, 0xc0, 0x4f, 0xd7, 0xd0, 0x62);
    
    public static Guid CLSID_ACListISF => new(0x03c036f1, 0xa186, 0x11d0, 0x82, 0x4a, 0x00, 0xaa, 0x00, 0x5b, 0x43, 0x83);
    
    public static Guid CLSID_ACLMRU => new(0x6756a641, 0xde71, 0x11d0, 0x83, 0x1b, 0x00, 0xaa, 0x00, 0x5b, 0x43, 0x83);
    
    public static Guid CLSID_ACLMulti => new(0x00bb2765, 0x6a77, 0x11d0, 0xa5, 0x35, 0x00, 0xc0, 0x4f, 0xd7, 0xd0, 0x62);
    
    public static Guid CLSID_ActiveDesktop => new(0x75048700, 0xef1f, 0x11d0, 0x98, 0x88, 0x00, 0x60, 0x97, 0xde, 0xac, 0xf9);
    
    public static Guid CLSID_AutoComplete => new(0x00bb2763, 0x6a77, 0x11d0, 0xa5, 0x35, 0x00, 0xc0, 0x4f, 0xd7, 0xd0, 0x62);
    
    public static Guid CLSID_CAnchorBrowsePropertyPage => new(0x3050f3bb, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b);
    
    public static Guid CLSID_CDocBrowsePropertyPage => new(0x3050f3b4, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b);
    
    public static Guid CLSID_CFSIconOverlayManager => new(0x63b51f81, 0xc868, 0x11d0, 0x99, 0x9c, 0x00, 0xc0, 0x4f, 0xd6, 0x55, 0xe1);
    
    public static Guid CLSID_CImageBrowsePropertyPage => new(0x3050f3b3, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b);
    
    public static Guid CLSID_ControlPanel => new(0x21ec2020, 0x3aea, 0x1069, 0xa2, 0xdd, 0x08, 0x00, 0x2b, 0x30, 0x30, 0x9d);
    
    public static Guid CLSID_CUrlHistory => new(0x3c374a40, 0xbae4, 0x11cf, 0xbf, 0x7d, 0x00, 0xaa, 0x00, 0x69, 0x46, 0xee);
    
    public static Guid CLSID_CUrlHistoryBoth => new(0x6659983c, 0x8476, 0x4eb4, 0xb7, 0x8c, 0xe5, 0x96, 0x8f, 0x32, 0x6b, 0xa0);
    
    public static Guid CLSID_CURLSearchHook => new(0xcfbfae00, 0x17a6, 0x11d0, 0x99, 0xcb, 0x00, 0xc0, 0x4f, 0xd6, 0x44, 0x97);
    
    public static Guid CLSID_DarwinAppPublisher => new(0xcfccc7a0, 0xa282, 0x11d1, 0x90, 0x82, 0x00, 0x60, 0x08, 0x05, 0x93, 0x82);
    
    public static Guid CLSID_DocHostUIHandler => new(0x7057e952, 0xbd1b, 0x11d1, 0x89, 0x19, 0x00, 0xc0, 0x4f, 0xc2, 0xc8, 0x36);
    
    public static Guid CLSID_DragDropHelper => new(0x4657278a, 0x411b, 0x11d2, 0x83, 0x9a, 0x00, 0xc0, 0x4f, 0xd9, 0x18, 0xd0);
    
    public static Guid CLSID_FileTypes => new(0xb091e540, 0x83e3, 0x11cf, 0xa7, 0x13, 0x00, 0x20, 0xaf, 0xd7, 0x97, 0x62);
    
    public static Guid CLSID_FolderItemsMultiLevel => new(0x53c74826, 0xab99, 0x4d33, 0xac, 0xa4, 0x31, 0x17, 0xf5, 0x1d, 0x37, 0x88);
    
    public static Guid CLSID_FolderShortcut => new(0x0afaced1, 0xe828, 0x11d1, 0x91, 0x87, 0xb5, 0x32, 0xf1, 0xe9, 0x57, 0x5d);
    
    public static Guid CLSID_HWShellExecute => new(0xffb8655f, 0x81b9, 0x4fce, 0xb8, 0x9c, 0x9a, 0x6b, 0xa7, 0x6d, 0x13, 0xe7);
    
    public static Guid CLSID_Internet => new(0x871c5380, 0x42a0, 0x1069, 0xa2, 0xea, 0x08, 0x00, 0x2b, 0x30, 0x30, 0x9d);
    
    public static Guid CLSID_InternetButtons => new(0x1e796980, 0x9cc5, 0x11d1, 0xa8, 0x3f, 0x00, 0xc0, 0x4f, 0xc9, 0x9d, 0x61);
    
    public static Guid CLSID_InternetShortcut => new(0xfbf23b40, 0xe3f0, 0x101b, 0x84, 0x88, 0x00, 0xaa, 0x00, 0x3e, 0x56, 0xf8);
    
    public static Guid CLSID_ISFBand => new(0xd82be2b0, 0x5764, 0x11d0, 0xa9, 0x6e, 0x00, 0xc0, 0x4f, 0xd7, 0x05, 0xa2);
    
    public static Guid CLSID_LinkColumnProvider => new(0x24f14f02, 0x7b1c, 0x11d1, 0x83, 0x8f, 0x00, 0x00, 0xf8, 0x04, 0x61, 0xcf);
    
    public static Guid CLSID_MenuBand => new(0x5b4dae26, 0xb807, 0x11d0, 0x98, 0x15, 0x00, 0xc0, 0x4f, 0xd9, 0x19, 0x72);
    
    public static Guid CLSID_MenuBandSite => new(0xe13ef4e4, 0xd2f2, 0x11d0, 0x98, 0x16, 0x00, 0xc0, 0x4f, 0xd9, 0x19, 0x72);
    
    public static Guid CLSID_MenuToolbarBase => new(0x40b96610, 0xb522, 0x11d1, 0xb3, 0xb4, 0x00, 0xaa, 0x00, 0x6e, 0xfd, 0xe7);
    
    public static Guid CLSID_MSOButtons => new(0x178f34b8, 0xa282, 0x11d2, 0x86, 0xc5, 0x00, 0xc0, 0x4f, 0x8e, 0xea, 0x99);
    
    public static Guid CLSID_MyComputer => new(0x20d04fe0, 0x3aea, 0x1069, 0xa2, 0xd8, 0x08, 0x00, 0x2b, 0x30, 0x30, 0x9d);
    
    public static Guid CLSID_MyDocuments => new(0x450d8fba, 0xad25, 0x11d0, 0x98, 0xa8, 0x08, 0x00, 0x36, 0x1b, 0x11, 0x03);
    
    public static Guid CLSID_NetworkDomain => new(0x46e06680, 0x4bf0, 0x11d1, 0x83, 0xee, 0x00, 0xa0, 0xc9, 0x0d, 0xc8, 0x49);
    
    public static Guid CLSID_NetworkServer => new(0xc0542a90, 0x4bf0, 0x11d1, 0x83, 0xee, 0x00, 0xa0, 0xc9, 0x0d, 0xc8, 0x49);
    
    public static Guid CLSID_NetworkShare => new(0x54a754c0, 0x4bf0, 0x11d1, 0x83, 0xee, 0x00, 0xa0, 0xc9, 0x0d, 0xc8, 0x49);
    
    public static Guid CLSID_NewMenu => new(0xd969a300, 0xe7ff, 0x11d0, 0xa9, 0x3b, 0x00, 0xa0, 0xc9, 0x0f, 0x27, 0x19);
    
    public static Guid CLSID_Printers => new(0x2227a280, 0x3aea, 0x1069, 0xa2, 0xde, 0x08, 0x00, 0x2b, 0x30, 0x30, 0x9d);
    
    public static Guid CLSID_ProgressDialog => new(0xf8383852, 0xfcd3, 0x11d1, 0xa6, 0xb9, 0x00, 0x60, 0x97, 0xdf, 0x5b, 0xd4);
    
    public static Guid CLSID_QueryAssociations => new(0xa07034fd, 0x6caa, 0x4954, 0xac, 0x3f, 0x97, 0xa2, 0x72, 0x16, 0xf9, 0x8a);
    
    public static Guid CLSID_QuickLinks => new(0x0e5cbf21, 0xd15f, 0x11d0, 0x83, 0x01, 0x00, 0xaa, 0x00, 0x5b, 0x43, 0x83);
    
    public static Guid CLSID_RecycleBin => new(0x645ff040, 0x5081, 0x101b, 0x9f, 0x08, 0x00, 0xaa, 0x00, 0x2f, 0x95, 0x4e);
    
    public static Guid CLSID_ShellFldSetExt => new(0x6d5313c0, 0x8c62, 0x11d1, 0xb2, 0xcd, 0x00, 0x60, 0x97, 0xdf, 0x8c, 0x11);
    
    public static Guid CLSID_ShellThumbnailDiskCache => new(0x1ebdcf80, 0xa200, 0x11d0, 0xa3, 0xa4, 0x00, 0xc0, 0x4f, 0xd7, 0x06, 0xec);
    
    public static Guid CLSID_ToolbarExtButtons => new(0x2ce4b5d8, 0xa28f, 0x11d2, 0x86, 0xc5, 0x00, 0xc0, 0x4f, 0x8e, 0xea, 0x99);
    
    public static Guid CLSID_WPD_NAMESPACE_EXTENSION => new(0x35786d3c, 0xb075, 0x49b9, 0x88, 0xdd, 0x02, 0x98, 0x76, 0xe1, 0x1c, 0x01);
    
    public const int CMDID_INTSHORTCUTCREATE = 1;
    
    public const string CMDSTR_NEWFOLDER = @"NewFolder";
    
    public const string CMDSTR_NEWFOLDERA = @"NewFolder";
    
    public const string CMDSTR_NEWFOLDERW = @"NewFolder";
    
    public const string CMDSTR_VIEWDETAILS = @"ViewDetails";
    
    public const string CMDSTR_VIEWDETAILSA = @"ViewDetails";
    
    public const string CMDSTR_VIEWDETAILSW = @"ViewDetails";
    
    public const string CMDSTR_VIEWLIST = @"ViewList";
    
    public const string CMDSTR_VIEWLISTA = @"ViewList";
    
    public const string CMDSTR_VIEWLISTW = @"ViewList";
    
    public const uint CMF_ASYNCVERBSTATE = 1024;
    
    public const uint CMF_CANRENAME = 16;
    
    public const uint CMF_DEFAULTONLY = 1;
    
    public const uint CMF_DISABLEDVERBS = 512;
    
    public const uint CMF_DONOTPICKDEFAULT = 8192;
    
    public const uint CMF_EXPLORE = 4;
    
    public const uint CMF_EXTENDEDVERBS = 256;
    
    public const uint CMF_INCLUDESTATIC = 64;
    
    public const uint CMF_ITEMMENU = 128;
    
    public const uint CMF_NODEFAULT = 32;
    
    public const uint CMF_NORMAL = 0;
    
    public const uint CMF_NOVERBS = 8;
    
    public const uint CMF_OPTIMIZEFORINVOKE = 2048;
    
    public const uint CMF_RESERVED = 4294901760;
    
    public const uint CMF_SYNCCASCADEMENU = 4096;
    
    public const uint CMF_VERBSONLY = 2;
    
    public const uint CMIC_MASK_CONTROL_DOWN = 1073741824;
    
    public const uint CMIC_MASK_PTINVOKE = 536870912;
    
    public const uint CMIC_MASK_SHIFT_DOWN = 268435456;
    
    public const uint COMP_ELEM_CHECKED = 2;
    
    public const uint COMP_ELEM_CURITEMSTATE = 16384;
    
    public const uint COMP_ELEM_DIRTY = 4;
    
    public const uint COMP_ELEM_FRIENDLYNAME = 1024;
    
    public const uint COMP_ELEM_NOSCROLL = 8;
    
    public const uint COMP_ELEM_ORIGINAL_CSI = 4096;
    
    public const uint COMP_ELEM_POS_LEFT = 16;
    
    public const uint COMP_ELEM_POS_TOP = 32;
    
    public const uint COMP_ELEM_POS_ZINDEX = 256;
    
    public const uint COMP_ELEM_RESTORED_CSI = 8192;
    
    public const uint COMP_ELEM_SIZE_HEIGHT = 128;
    
    public const uint COMP_ELEM_SIZE_WIDTH = 64;
    
    public const uint COMP_ELEM_SOURCE = 512;
    
    public const uint COMP_ELEM_SUBSCRIBEDURL = 2048;
    
    public const uint COMP_ELEM_TYPE = 1;
    
    public const uint COMP_TYPE_CFHTML = 4;
    
    public const uint COMP_TYPE_CONTROL = 3;
    
    public const uint COMP_TYPE_HTMLDOC = 0;
    
    public const uint COMP_TYPE_MAX = 4;
    
    public const uint COMP_TYPE_PICTURE = 1;
    
    public const uint COMP_TYPE_WEBSITE = 2;
    
    public const uint COMPONENT_DEFAULT_LEFT = 65535;
    
    public const uint COMPONENT_DEFAULT_TOP = 65535;
    
    public const uint COMPONENT_TOP = 1073741823;
    
    public const string CONFLICT_RESOLUTION_CLSID_KEY = @"ConflictResolutionCLSID";
    
    public static Guid ConflictFolder => new(0x289978ac, 0xa101, 0x4341, 0xa8, 0x17, 0x21, 0xeb, 0xa7, 0xfd, 0x04, 0x6d);
    
    public static Guid CPFG_CREDENTIAL_PROVIDER_LABEL => new(0x286bbff3, 0xbad4, 0x438f, 0xb0, 0x07, 0x79, 0xb7, 0x26, 0x7c, 0x3d, 0x48);
    
    public static Guid CPFG_CREDENTIAL_PROVIDER_LOGO => new(0x2d837775, 0xf6cd, 0x464e, 0xa7, 0x45, 0x48, 0x2f, 0xd0, 0xb4, 0x74, 0x93);
    
    public static Guid CPFG_LOGON_PASSWORD => new(0x60624cfa, 0xa477, 0x47b1, 0x8a, 0x8e, 0x3a, 0x4a, 0x19, 0x98, 0x18, 0x27);
    
    public static Guid CPFG_LOGON_USERNAME => new(0xda15bbe8, 0x954d, 0x4fd3, 0xb0, 0xf4, 0x1f, 0xb5, 0xb9, 0x0b, 0x17, 0x4b);
    
    public static Guid CPFG_SMARTCARD_PIN => new(0x4fe5263b, 0x9181, 0x46c1, 0xb0, 0xa4, 0x9d, 0xed, 0xd4, 0xdb, 0x7d, 0xea);
    
    public static Guid CPFG_SMARTCARD_USERNAME => new(0x3e1ecf69, 0x568c, 0x4d96, 0x9d, 0x59, 0x46, 0x44, 0x41, 0x74, 0xe2, 0xd6);
    
    public static Guid CPFG_STANDALONE_SUBMIT_BUTTON => new(0x0b7b0ad8, 0xcc36, 0x4d59, 0x80, 0x2b, 0x82, 0xf7, 0x14, 0xfa, 0x70, 0x22);
    
    public static Guid CPFG_STYLE_LINK_AS_BUTTON => new(0x088fa508, 0x94a6, 0x4430, 0xa4, 0xcb, 0x6f, 0xc6, 0xe3, 0xc0, 0xb9, 0xe2);
    
    public const uint CPL_DBLCLK = 5;
    
    public const uint CPL_DYNAMIC_RES = 0;
    
    public const uint CPL_EXIT = 7;
    
    public const uint CPL_GETCOUNT = 2;
    
    public const uint CPL_INIT = 1;
    
    public const uint CPL_INQUIRE = 3;
    
    public const uint CPL_NEWINQUIRE = 8;
    
    public const uint CPL_SELECT = 4;
    
    public const uint CPL_SETUP = 200;
    
    public const uint CPL_STARTWPARMS = 10;
    
    public const uint CPL_STARTWPARMSA = 9;
    
    public const uint CPL_STARTWPARMSW = 10;
    
    public const uint CPL_STOP = 6;
    
    public const uint CPLPAGE_DISPLAY_BACKGROUND = 1;
    
    public const uint CPLPAGE_KEYBOARD_SPEED = 1;
    
    public const uint CPLPAGE_MOUSE_BUTTONS = 1;
    
    public const uint CPLPAGE_MOUSE_PTRMOTION = 2;
    
    public const uint CPLPAGE_MOUSE_WHEEL = 3;
    
    public const uint CREDENTIAL_PROVIDER_NO_DEFAULT = uint.MaxValue;
    
    public static Guid CScriptErrorList => new(0xefd01300, 0x160f, 0x11d2, 0xbb, 0x2e, 0x00, 0x80, 0x5f, 0xf7, 0xef, 0xca);
    
    public static Guid CSearchManager => new(0x7d096c5f, 0xac08, 0x4f1f, 0xbe, 0xb7, 0x5c, 0x22, 0xc5, 0x17, 0xce, 0x39);
    
    public const uint CSIDL_ADMINTOOLS = 48;
    
    public const uint CSIDL_ALTSTARTUP = 29;
    
    public const uint CSIDL_APPDATA = 26;
    
    public const uint CSIDL_BITBUCKET = 10;
    
    public const uint CSIDL_CDBURN_AREA = 59;
    
    public const uint CSIDL_COMMON_ADMINTOOLS = 47;
    
    public const uint CSIDL_COMMON_ALTSTARTUP = 30;
    
    public const uint CSIDL_COMMON_APPDATA = 35;
    
    public const uint CSIDL_COMMON_DESKTOPDIRECTORY = 25;
    
    public const uint CSIDL_COMMON_DOCUMENTS = 46;
    
    public const uint CSIDL_COMMON_FAVORITES = 31;
    
    public const uint CSIDL_COMMON_MUSIC = 53;
    
    public const uint CSIDL_COMMON_OEM_LINKS = 58;
    
    public const uint CSIDL_COMMON_PICTURES = 54;
    
    public const uint CSIDL_COMMON_PROGRAMS = 23;
    
    public const uint CSIDL_COMMON_STARTMENU = 22;
    
    public const uint CSIDL_COMMON_STARTUP = 24;
    
    public const uint CSIDL_COMMON_TEMPLATES = 45;
    
    public const uint CSIDL_COMMON_VIDEO = 55;
    
    public const uint CSIDL_COMPUTERSNEARME = 61;
    
    public const uint CSIDL_CONNECTIONS = 49;
    
    public const uint CSIDL_CONTROLS = 3;
    
    public const uint CSIDL_COOKIES = 33;
    
    public const uint CSIDL_DESKTOP = 0;
    
    public const uint CSIDL_DESKTOPDIRECTORY = 16;
    
    public const uint CSIDL_DRIVES = 17;
    
    public const uint CSIDL_FAVORITES = 6;
    
    public const uint CSIDL_FLAG_CREATE = 32768;
    
    public const uint CSIDL_FLAG_DONT_UNEXPAND = 8192;
    
    public const uint CSIDL_FLAG_DONT_VERIFY = 16384;
    
    public const uint CSIDL_FLAG_MASK = 65280;
    
    public const uint CSIDL_FLAG_NO_ALIAS = 4096;
    
    public const uint CSIDL_FLAG_PER_USER_INIT = 2048;
    
    public const uint CSIDL_FLAG_PFTI_TRACKTARGET = 16384;
    
    public const uint CSIDL_FONTS = 20;
    
    public const uint CSIDL_HISTORY = 34;
    
    public const uint CSIDL_INTERNET = 1;
    
    public const uint CSIDL_INTERNET_CACHE = 32;
    
    public const uint CSIDL_LOCAL_APPDATA = 28;
    
    public const uint CSIDL_MYDOCUMENTS = 5;
    
    public const uint CSIDL_MYMUSIC = 13;
    
    public const uint CSIDL_MYPICTURES = 39;
    
    public const uint CSIDL_MYVIDEO = 14;
    
    public const uint CSIDL_NETHOOD = 19;
    
    public const uint CSIDL_NETWORK = 18;
    
    public const uint CSIDL_PERSONAL = 5;
    
    public const uint CSIDL_PRINTERS = 4;
    
    public const uint CSIDL_PRINTHOOD = 27;
    
    public const uint CSIDL_PROFILE = 40;
    
    public const uint CSIDL_PROGRAM_FILES = 38;
    
    public const uint CSIDL_PROGRAM_FILES_COMMON = 43;
    
    public const uint CSIDL_PROGRAM_FILES_COMMONX86 = 44;
    
    public const uint CSIDL_PROGRAM_FILESX86 = 42;
    
    public const uint CSIDL_PROGRAMS = 2;
    
    public const uint CSIDL_RECENT = 8;
    
    public const uint CSIDL_RESOURCES = 56;
    
    public const uint CSIDL_RESOURCES_LOCALIZED = 57;
    
    public const uint CSIDL_SENDTO = 9;
    
    public const uint CSIDL_STARTMENU = 11;
    
    public const uint CSIDL_STARTUP = 7;
    
    public const uint CSIDL_SYSTEM = 37;
    
    public const uint CSIDL_SYSTEMX86 = 41;
    
    public const uint CSIDL_TEMPLATES = 21;
    
    public const uint CSIDL_WINDOWS = 36;
    
    public const int CTF_COINIT = 8;
    
    public const int CTF_COINIT_MTA = 4096;
    
    public const int CTF_COINIT_STA = 8;
    
    public const int CTF_FREELIBANDEXIT = 16;
    
    public const int CTF_INHERITWOW64 = 256;
    
    public const int CTF_INSIST = 1;
    
    public const int CTF_KEYBOARD_LOCALE = 1024;
    
    public const int CTF_NOADDREFLIB = 8192;
    
    public const int CTF_OLEINITIALIZE = 2048;
    
    public const int CTF_PROCESS_REF = 4;
    
    public const int CTF_REF_COUNTED = 32;
    
    public const int CTF_THREAD_REF = 2;
    
    public const int CTF_UNUSED = 128;
    
    public const int CTF_WAIT_ALLOWCOM = 64;
    
    public const int CTF_WAIT_NO_REENTRANCY = 512;
    
    public const uint DBC_GS_IDEAL = 0;
    
    public const uint DBC_GS_SIZEDOWN = 1;
    
    public const uint DBC_HIDE = 0;
    
    public const uint DBC_SHOW = 1;
    
    public const uint DBC_SHOWOBSCURE = 2;
    
    public const int DBCID_CLSIDOFBAR = 2;
    
    public const int DBCID_EMPTY = 0;
    
    public const int DBCID_GETBAR = 4;
    
    public const int DBCID_ONDRAG = 1;
    
    public const int DBCID_RESIZE = 3;
    
    public const int DBCID_UPDATESIZE = 5;
    
    public const uint DBIF_VIEWMODE_FLOATING = 2;
    
    public const uint DBIF_VIEWMODE_NORMAL = 0;
    
    public const uint DBIF_VIEWMODE_TRANSPARENT = 4;
    
    public const uint DBIF_VIEWMODE_VERTICAL = 1;
    
    public const uint DBIM_ACTUAL = 8;
    
    public const uint DBIM_BKCOLOR = 64;
    
    public const uint DBIM_INTEGRAL = 4;
    
    public const uint DBIM_MAXSIZE = 2;
    
    public const uint DBIM_MINSIZE = 1;
    
    public const uint DBIM_MODEFLAGS = 32;
    
    public const uint DBIM_TITLE = 16;
    
    public const uint DBIMF_ADDTOFRONT = 512;
    
    public const uint DBIMF_ALWAYSGRIPPER = 4096;
    
    public const uint DBIMF_BKCOLOR = 64;
    
    public const uint DBIMF_BREAK = 256;
    
    public const uint DBIMF_DEBOSSED = 32;
    
    public const uint DBIMF_FIXED = 1;
    
    public const uint DBIMF_FIXEDBMP = 4;
    
    public const uint DBIMF_NOGRIPPER = 2048;
    
    public const uint DBIMF_NOMARGINS = 8192;
    
    public const uint DBIMF_NORMAL = 0;
    
    public const uint DBIMF_TOPALIGN = 1024;
    
    public const uint DBIMF_UNDELETEABLE = 16;
    
    public const uint DBIMF_USECHEVRON = 128;
    
    public const uint DBIMF_VARIABLEHEIGHT = 8;
    
    public const uint DBPC_SELECTFIRST = uint.MaxValue;
    
    public const uint DBT_APPYBEGIN = 0;
    
    public const uint DBT_APPYEND = 1;
    
    public const uint DBT_CONFIGCHANGECANCELED = 25;
    
    public const uint DBT_CONFIGCHANGED = 24;
    
    public const uint DBT_CONFIGMGAPI32 = 34;
    
    public const uint DBT_CONFIGMGPRIVATE = 32767;
    
    public const uint DBT_CUSTOMEVENT = 32774;
    
    public const uint DBT_DEVICEARRIVAL = 32768;
    
    public const uint DBT_DEVICEQUERYREMOVE = 32769;
    
    public const uint DBT_DEVICEQUERYREMOVEFAILED = 32770;
    
    public const uint DBT_DEVICEREMOVECOMPLETE = 32772;
    
    public const uint DBT_DEVICEREMOVEPENDING = 32771;
    
    public const uint DBT_DEVICETYPESPECIFIC = 32773;
    
    public const uint DBT_DEVNODES_CHANGED = 7;
    
    public const uint DBT_DEVTYP_DEVNODE = 1;
    
    public const uint DBT_DEVTYP_NET = 4;
    
    public const uint DBT_LOW_DISK_SPACE = 72;
    
    public const uint DBT_MONITORCHANGE = 27;
    
    public const uint DBT_NO_DISK_SPACE = 71;
    
    public const uint DBT_QUERYCHANGECONFIG = 23;
    
    public const uint DBT_SHELLLOGGEDON = 32;
    
    public const uint DBT_USERDEFINED = 65535;
    
    public const uint DBT_VOLLOCKLOCKFAILED = 32835;
    
    public const uint DBT_VOLLOCKLOCKRELEASED = 32837;
    
    public const uint DBT_VOLLOCKLOCKTAKEN = 32834;
    
    public const uint DBT_VOLLOCKQUERYLOCK = 32833;
    
    public const uint DBT_VOLLOCKQUERYUNLOCK = 32836;
    
    public const uint DBT_VOLLOCKUNLOCKFAILED = 32838;
    
    public const uint DBT_VPOWERDAPI = 33024;
    
    public const uint DBT_VXDINITCOMPLETE = 35;
    
    public static Guid DefFolderMenu => new(0xc63382be, 0x7933, 0x48d0, 0x9a, 0xc8, 0x85, 0xfb, 0x46, 0xbe, 0x2f, 0xdd);
    
    public static Guid DesktopGadget => new(0x924ccc1b, 0x6562, 0x4c85, 0x86, 0x57, 0xd1, 0x77, 0x92, 0x52, 0x22, 0xb6);
    
    public static Guid DesktopWallpaper => new(0xc2cf3110, 0x460e, 0x4fc1, 0xb9, 0xd0, 0x8a, 0x1c, 0x0c, 0x9c, 0xc4, 0xbd);
    
    public static Guid DestinationList => new(0x77f10cf0, 0x3db5, 0x4966, 0xb5, 0x20, 0xb7, 0xc5, 0x4f, 0xd3, 0x5e, 0xd6);
    
    public static Guid DestinationListBoth => new(0x38fe0cf4, 0x6a59, 0x4729, 0x8e, 0x4a, 0x2d, 0x58, 0x00, 0x59, 0xed, 0xe4);
    
    public static DEVPROPKEY DEVPKEY_MTPBTH_IsConnected => new(new(3927062522, 22685, 17522, 132, 228, 10, 190, 54, 253, 98, 239), 2);
    
    public const uint DEVSVC_SERVICEINFO_VERSION = 100;
    
    public const uint DEVSVCTYPE_ABSTRACT = 1;
    
    public const uint DEVSVCTYPE_DEFAULT = 0;
    
    public const string DI_GETDRAGIMAGE = @"ShellGetDragImage";
    
    public const uint DISPID_BEGINDRAG = 204;
    
    public const uint DISPID_CHECKSTATECHANGED = 209;
    
    public const uint DISPID_COLUMNSCHANGED = 212;
    
    public const uint DISPID_CONTENTSCHANGED = 207;
    
    public const uint DISPID_CTRLMOUSEWHEEL = 213;
    
    public const uint DISPID_DEFAULTVERBINVOKED = 203;
    
    public const uint DISPID_ENTERPRESSED = 200;
    
    public const uint DISPID_ENTERPRISEIDCHANGED = 224;
    
    public const uint DISPID_EXPLORERWINDOWREADY = 221;
    
    public const uint DISPID_FILELISTENUMDONE = 201;
    
    public const uint DISPID_FILTERINVOKED = 218;
    
    public const uint DISPID_FOCUSCHANGED = 208;
    
    public const uint DISPID_FOLDERCHANGED = 217;
    
    public const uint DISPID_IADCCTL_DEFAULTCAT = 262;
    
    public const uint DISPID_IADCCTL_DIRTY = 256;
    
    public const uint DISPID_IADCCTL_FORCEX86 = 259;
    
    public const uint DISPID_IADCCTL_ONDOMAIN = 261;
    
    public const uint DISPID_IADCCTL_PUBCAT = 257;
    
    public const uint DISPID_IADCCTL_SHOWPOSTSETUP = 260;
    
    public const uint DISPID_IADCCTL_SORT = 258;
    
    public const uint DISPID_ICONSIZECHANGED = 215;
    
    public const uint DISPID_INITIALENUMERATIONDONE = 223;
    
    public const uint DISPID_NOITEMSTATE_CHANGED = 206;
    
    public const uint DISPID_ORDERCHANGED = 210;
    
    public const uint DISPID_SEARCHCOMMAND_ABORT = 3;
    
    public const uint DISPID_SEARCHCOMMAND_COMPLETE = 2;
    
    public const uint DISPID_SEARCHCOMMAND_ERROR = 6;
    
    public const uint DISPID_SEARCHCOMMAND_PROGRESSTEXT = 5;
    
    public const uint DISPID_SEARCHCOMMAND_RESTORE = 7;
    
    public const uint DISPID_SEARCHCOMMAND_START = 1;
    
    public const uint DISPID_SEARCHCOMMAND_UPDATE = 4;
    
    public const uint DISPID_SELECTEDITEMCHANGED = 220;
    
    public const uint DISPID_SELECTIONCHANGED = 200;
    
    public const uint DISPID_SORTDONE = 214;
    
    public const uint DISPID_UPDATEIMAGE = 222;
    
    public const uint DISPID_VERBINVOKED = 202;
    
    public const uint DISPID_VIEWMODECHANGED = 205;
    
    public const uint DISPID_VIEWPAINTDONE = 211;
    
    public const uint DISPID_WORDWHEELEDITED = 219;
    
    public const uint DLG_SCRNSAVECONFIGURE = 2003;
    
    public const ulong DLLVER_BUILD_MASK = 4294901760;
    
    public const ulong DLLVER_MAJOR_MASK = 18446462598732840960;
    
    public const ulong DLLVER_MINOR_MASK = 281470681743360;
    
    public const uint DLLVER_PLATFORM_NT = 2;
    
    public const uint DLLVER_PLATFORM_WINDOWS = 1;
    
    public const ulong DLLVER_QFE_MASK = 65535;
    
    public static Guid DocPropShellExtension => new(0x883373c3, 0xbf89, 0x11d1, 0xbe, 0x35, 0x08, 0x00, 0x36, 0xb1, 0x1a, 0x03);
    
    public static Guid DriveSizeCategorizer => new(0x94357b53, 0xca29, 0x4b78, 0x83, 0xae, 0xe8, 0xfe, 0x74, 0x09, 0x13, 0x4f);
    
    public static Guid DriveTypeCategorizer => new(0xb0a8f3cf, 0x4333, 0x4bab, 0x88, 0x73, 0x1c, 0xcb, 0x1c, 0xad, 0xa4, 0x8b);
    
    public const uint DVASPECT_COPY = 3;
    
    public const uint DVASPECT_LINK = 4;
    
    public const uint DVASPECT_SHORTNAME = 2;
    
    public const uint DWFAF_AUTOHIDE = 16;
    
    public const uint DWFAF_GROUP1 = 2;
    
    public const uint DWFAF_GROUP2 = 4;
    
    public const uint DWFAF_HIDDEN = 1;
    
    public const uint DWFRF_DELETECONFIGDATA = 1;
    
    public const uint DWFRF_NORMAL = 0;
    
    public const uint ENUM_AnchorResults_AnchorStateInvalid = 1;
    
    public const uint ENUM_AnchorResults_AnchorStateNormal = 0;
    
    public const uint ENUM_AnchorResults_AnchorStateOld = 2;
    
    public const uint ENUM_AnchorResults_ItemStateChanged = 4;
    
    public const uint ENUM_AnchorResults_ItemStateCreated = 2;
    
    public const uint ENUM_AnchorResults_ItemStateDeleted = 1;
    
    public const uint ENUM_AnchorResults_ItemStateInvalid = 0;
    
    public const uint ENUM_AnchorResults_ItemStateUpdated = 3;
    
    public const uint ENUM_CalendarObj_BusyStatusBusy = 1;
    
    public const uint ENUM_CalendarObj_BusyStatusFree = 0;
    
    public const uint ENUM_CalendarObj_BusyStatusOutOfOffice = 2;
    
    public const uint ENUM_CalendarObj_BusyStatusTentative = 3;
    
    public const uint ENUM_DeviceMetadataObj_DefaultCABFalse = 0;
    
    public const uint ENUM_DeviceMetadataObj_DefaultCABTrue = 1;
    
    public const uint ENUM_MessageObj_PatternInstanceFirst = 1;
    
    public const uint ENUM_MessageObj_PatternInstanceFourth = 4;
    
    public const uint ENUM_MessageObj_PatternInstanceLast = 5;
    
    public const uint ENUM_MessageObj_PatternInstanceNone = 0;
    
    public const uint ENUM_MessageObj_PatternInstanceSecond = 2;
    
    public const uint ENUM_MessageObj_PatternInstanceThird = 3;
    
    public const uint ENUM_MessageObj_PatternTypeDaily = 1;
    
    public const uint ENUM_MessageObj_PatternTypeMonthly = 3;
    
    public const uint ENUM_MessageObj_PatternTypeWeekly = 2;
    
    public const uint ENUM_MessageObj_PatternTypeYearly = 4;
    
    public const uint ENUM_MessageObj_PriorityHighest = 2;
    
    public const uint ENUM_MessageObj_PriorityLowest = 0;
    
    public const uint ENUM_MessageObj_PriorityNormal = 1;
    
    public const uint ENUM_MessageObj_ReadFalse = 0;
    
    public const uint ENUM_MessageObj_ReadTrue = 255;
    
    public const uint ENUM_StatusSvc_ChargingActive = 1;
    
    public const uint ENUM_StatusSvc_ChargingInactive = 0;
    
    public const uint ENUM_StatusSvc_ChargingUnknown = 2;
    
    public const uint ENUM_StatusSvc_RoamingActive = 1;
    
    public const uint ENUM_StatusSvc_RoamingInactive = 0;
    
    public const uint ENUM_StatusSvc_RoamingUnknown = 2;
    
    public const uint ENUM_SyncSvc_SyncObjectReferencesDisabled = 0;
    
    public const uint ENUM_SyncSvc_SyncObjectReferencesEnabled = 255;
    
    public const uint ENUM_TaskObj_CompleteFalse = 0;
    
    public const uint ENUM_TaskObj_CompleteTrue = 255;
    
    public static Guid EnumBthMtpConnectors => new(0xa1570149, 0xe645, 0x4f43, 0x8b, 0x0d, 0x40, 0x9b, 0x06, 0x1d, 0xb2, 0xfc);
    
    public static Guid EnumerableObjectCollection => new(0x2d3468c1, 0x36a7, 0x43b6, 0xac, 0x24, 0xd3, 0xf0, 0x2f, 0xd9, 0x60, 0x7a);
    
    public static Guid EP_AdvQueryPane => new(0xb4e9db8b, 0x34ba, 0x4c39, 0xb5, 0xcc, 0x16, 0xa1, 0xbd, 0x2c, 0x41, 0x1c);
    
    public static Guid EP_Commands => new(0xd9745868, 0xca5f, 0x4a76, 0x91, 0xcd, 0xf5, 0xa1, 0x29, 0xfb, 0xb0, 0x76);
    
    public static Guid EP_Commands_Organize => new(0x72e81700, 0xe3ec, 0x4660, 0xbf, 0x24, 0x3c, 0x3b, 0x7b, 0x64, 0x88, 0x06);
    
    public static Guid EP_Commands_View => new(0x21f7c32d, 0xeeaa, 0x439b, 0xbb, 0x51, 0x37, 0xb9, 0x6f, 0xd6, 0xa9, 0x43);
    
    public static Guid EP_DetailsPane => new(0x43abf98b, 0x89b8, 0x472d, 0xb9, 0xce, 0xe6, 0x9b, 0x82, 0x29, 0xf0, 0x19);
    
    public static Guid EP_NavPane => new(0xcb316b22, 0x25f7, 0x42b8, 0x8a, 0x09, 0x54, 0x0d, 0x23, 0xa4, 0x3c, 0x2f);
    
    public static Guid EP_PreviewPane => new(0x893c63d1, 0x45c8, 0x4d17, 0xbe, 0x19, 0x22, 0x3b, 0xe7, 0x1b, 0xe3, 0x65);
    
    public static Guid EP_QueryPane => new(0x65bcde4f, 0x4f07, 0x4f27, 0x83, 0xa7, 0x1a, 0xfc, 0xa4, 0xdf, 0x7d, 0xdd);
    
    public static Guid EP_Ribbon => new(0xd27524a8, 0xc9f2, 0x4834, 0xa1, 0x06, 0xdf, 0x88, 0x89, 0xfd, 0x4f, 0x37);
    
    public static Guid EP_StatusBar => new(0x65fe56ce, 0x5cfe, 0x4bc4, 0xad, 0x8a, 0x7a, 0xe3, 0xfe, 0x7e, 0x8f, 0x7c);
    
    public static Guid ExecuteFolder => new(0x11dbb47c, 0xa525, 0x400b, 0x9e, 0x80, 0xa5, 0x46, 0x15, 0xa0, 0x90, 0xc0);
    
    public static Guid ExecuteUnknown => new(0xe44e9428, 0xbdbc, 0x4987, 0xa0, 0x99, 0x40, 0xdc, 0x8f, 0xd2, 0x55, 0xe7);
    
    public const uint EXP_DARWIN_ID_SIG = 2684354566;
    
    public const uint EXP_PROPERTYSTORAGE_SIG = 2684354569;
    
    public const uint EXP_SPECIAL_FOLDER_SIG = 2684354565;
    
    public const uint EXP_SZ_ICON_SIG = 2684354567;
    
    public const uint EXP_SZ_LINK_SIG = 2684354561;
    
    public static Guid ExplorerBrowser => new(0x71f96385, 0xddd6, 0x48d3, 0xa0, 0xc1, 0xae, 0x06, 0xe8, 0xb0, 0x55, 0xfb);
    
    public const uint FACILITY_WPD = 42;
    
    public const uint FCIDM_BROWSERFIRST = 40960;
    
    public const uint FCIDM_BROWSERLAST = 48896;
    
    public const uint FCIDM_GLOBALFIRST = 32768;
    
    public const uint FCIDM_GLOBALLAST = 40959;
    
    public const uint FCIDM_MENU_EDIT = 32832;
    
    public const uint FCIDM_MENU_EXPLORE = 33104;
    
    public const uint FCIDM_MENU_FAVORITES = 33136;
    
    public const uint FCIDM_MENU_FILE = 32768;
    
    public const uint FCIDM_MENU_FIND = 33088;
    
    public const uint FCIDM_MENU_HELP = 33024;
    
    public const uint FCIDM_MENU_TOOLS = 32960;
    
    public const uint FCIDM_MENU_TOOLS_SEP_GOTO = 32961;
    
    public const uint FCIDM_MENU_VIEW = 32896;
    
    public const uint FCIDM_MENU_VIEW_SEP_OPTIONS = 32897;
    
    public const uint FCIDM_SHVIEWFIRST = 0;
    
    public const uint FCIDM_SHVIEWLAST = 32767;
    
    public const uint FCIDM_STATUS = 40961;
    
    public const uint FCIDM_TOOLBAR = 40960;
    
    public const uint FCS_FLAG_DRAGDROP = 2;
    
    public const uint FCS_FORCEWRITE = 2;
    
    public const uint FCS_READ = 1;
    
    public const uint FCSM_CLSID = 8;
    
    public const uint FCSM_FLAGS = 64;
    
    public const uint FCSM_ICONFILE = 16;
    
    public const uint FCSM_INFOTIP = 4;
    
    public const uint FCSM_LOGO = 32;
    
    public const uint FCSM_VIEWID = 1;
    
    public const uint FCSM_WEBVIEWTEMPLATE = 2;
    
    public const uint FCT_ADDTOEND = 4;
    
    public const uint FCT_CONFIGABLE = 2;
    
    public const uint FCT_MERGE = 1;
    
    public const uint FCW_INTERNETBAR = 6;
    
    public const uint FCW_PROGRESS = 8;
    
    public const uint FCW_STATUS = 1;
    
    public const uint FCW_TOOLBAR = 2;
    
    public const uint FCW_TREE = 3;
    
    public const uint FDTF_LONGDATE = 4;
    
    public const uint FDTF_LONGTIME = 8;
    
    public const uint FDTF_LTRDATE = 256;
    
    public const uint FDTF_NOAUTOREADINGORDER = 1024;
    
    public const uint FDTF_RELATIVE = 16;
    
    public const uint FDTF_RTLDATE = 512;
    
    public const uint FDTF_SHORTDATE = 2;
    
    public const uint FDTF_SHORTTIME = 1;
    
    public static Guid FileOpenDialog => new(0xdc1c5a9c, 0xe88a, 0x4dde, 0xa5, 0xa1, 0x60, 0xf8, 0x2a, 0x20, 0xae, 0xf7);
    
    public static Guid FileOperation => new(0x3ad05575, 0x8857, 0x4850, 0x92, 0x77, 0x11, 0xb8, 0x5b, 0xdb, 0x8e, 0x09);
    
    public static Guid FileSaveDialog => new(0xc0b4e2f3, 0xba21, 0x4773, 0x8d, 0xba, 0x33, 0x5e, 0xc9, 0x46, 0xeb, 0x8b);
    
    public static Guid FileSearchBand => new(0xc4ee31f3, 0x4768, 0x11d2, 0xbe, 0x5c, 0x00, 0xa0, 0xc9, 0xa8, 0x3d, 0xa1);
    
    public const uint FLAG_MessageObj_DayOfWeekFriday = 32;
    
    public const uint FLAG_MessageObj_DayOfWeekMonday = 2;
    
    public const uint FLAG_MessageObj_DayOfWeekNone = 0;
    
    public const uint FLAG_MessageObj_DayOfWeekSaturday = 64;
    
    public const uint FLAG_MessageObj_DayOfWeekSunday = 1;
    
    public const uint FLAG_MessageObj_DayOfWeekThursday = 16;
    
    public const uint FLAG_MessageObj_DayOfWeekTuesday = 4;
    
    public const uint FLAG_MessageObj_DayOfWeekWednesday = 8;
    
    public static Guid FMTID_Briefcase => new(0x328d8b21, 0x7729, 0x4bfc, 0x95, 0x4c, 0x90, 0x2b, 0x32, 0x9d, 0x56, 0xb0);
    
    public static Guid FMTID_CustomImageProperties => new(0x7ecd8b0e, 0xc136, 0x4a9b, 0x94, 0x11, 0x4e, 0xbd, 0x66, 0x73, 0xcc, 0xc3);
    
    public static Guid FMTID_Displaced => new(0x9b174b33, 0x40ff, 0x11d2, 0xa2, 0x7e, 0x00, 0xc0, 0x4f, 0xc3, 0x08, 0x71);
    
    public static Guid FMTID_DRM => new(0xaeac19e4, 0x89ae, 0x4508, 0xb9, 0xb7, 0xbb, 0x86, 0x7a, 0xbe, 0xe2, 0xed);
    
    public static Guid FMTID_ImageProperties => new(0x14b81da1, 0x0135, 0x4d31, 0x96, 0xd9, 0x6c, 0xbf, 0xc9, 0x67, 0x1a, 0x99);
    
    public static Guid FMTID_InternetSite => new(0x000214a1, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid FMTID_Intshcut => new(0x000214a0, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid FMTID_LibraryProperties => new(0x5d76b67f, 0x9b3d, 0x44bb, 0xb6, 0xae, 0x25, 0xda, 0x4f, 0x63, 0x8a, 0x67);
    
    public static Guid FMTID_Misc => new(0x9b174b34, 0x40ff, 0x11d2, 0xa2, 0x7e, 0x00, 0xc0, 0x4f, 0xc3, 0x08, 0x71);
    
    public static Guid FMTID_MUSIC => new(0x56a3372e, 0xce9c, 0x11d2, 0x9f, 0x0e, 0x00, 0x60, 0x97, 0xc6, 0x86, 0xf6);
    
    public static Guid FMTID_Query => new(0x49691c90, 0x7e17, 0x101a, 0xa9, 0x1c, 0x08, 0x00, 0x2b, 0x2e, 0xcd, 0xa9);
    
    public static Guid FMTID_ShellDetails => new(0x28636aa6, 0x953d, 0x11d2, 0xb5, 0xd6, 0x00, 0xc0, 0x4f, 0xd9, 0x18, 0xd0);
    
    public static Guid FMTID_Storage => new(0xb725f130, 0x47ef, 0x101a, 0xa5, 0xf1, 0x02, 0x60, 0x8c, 0x9e, 0xeb, 0xac);
    
    public static Guid FMTID_Volume => new(0x9b174b35, 0x40ff, 0x11d2, 0xa2, 0x7e, 0x00, 0xc0, 0x4f, 0xc3, 0x08, 0x71);
    
    public static Guid FMTID_WebView => new(0xf2275480, 0xf782, 0x4291, 0xbd, 0x94, 0xf1, 0x36, 0x93, 0x51, 0x3a, 0xec);
    
    public const uint FO_COPY = 2;
    
    public const uint FO_DELETE = 3;
    
    public const uint FO_MOVE = 1;
    
    public const uint FO_RENAME = 4;
    
    public static Guid FOLDERID_AccountPictures => new(0x008ca0b1, 0x55b4, 0x4c56, 0xb8, 0xa8, 0x4d, 0xe4, 0xb2, 0x99, 0xd3, 0xbe);
    
    public static Guid FOLDERID_AddNewPrograms => new(0xde61d971, 0x5ebc, 0x4f02, 0xa3, 0xa9, 0x6c, 0x82, 0x89, 0x5e, 0x5c, 0x04);
    
    public static Guid FOLDERID_AdminTools => new(0x724ef170, 0xa42d, 0x4fef, 0x9f, 0x26, 0xb6, 0x0e, 0x84, 0x6f, 0xba, 0x4f);
    
    public static Guid FOLDERID_AllAppMods => new(0x7ad67899, 0x66af, 0x43ba, 0x91, 0x56, 0x6a, 0xad, 0x42, 0xe6, 0xc5, 0x96);
    
    public static Guid FOLDERID_AppCaptures => new(0xedc0fe71, 0x98d8, 0x4f4a, 0xb9, 0x20, 0xc8, 0xdc, 0x13, 0x3c, 0xb1, 0x65);
    
    public static Guid FOLDERID_AppDataDesktop => new(0xb2c5e279, 0x7add, 0x439f, 0xb2, 0x8c, 0xc4, 0x1f, 0xe1, 0xbb, 0xf6, 0x72);
    
    public static Guid FOLDERID_AppDataDocuments => new(0x7be16610, 0x1f7f, 0x44ac, 0xbf, 0xf0, 0x83, 0xe1, 0x5f, 0x2f, 0xfc, 0xa1);
    
    public static Guid FOLDERID_AppDataFavorites => new(0x7cfbefbc, 0xde1f, 0x45aa, 0xb8, 0x43, 0xa5, 0x42, 0xac, 0x53, 0x6c, 0xc9);
    
    public static Guid FOLDERID_AppDataProgramData => new(0x559d40a3, 0xa036, 0x40fa, 0xaf, 0x61, 0x84, 0xcb, 0x43, 0x0a, 0x4d, 0x34);
    
    public static Guid FOLDERID_ApplicationShortcuts => new(0xa3918781, 0xe5f2, 0x4890, 0xb3, 0xd9, 0xa7, 0xe5, 0x43, 0x32, 0x32, 0x8c);
    
    public static Guid FOLDERID_AppsFolder => new(0x1e87508d, 0x89c2, 0x42f0, 0x8a, 0x7e, 0x64, 0x5a, 0x0f, 0x50, 0xca, 0x58);
    
    public static Guid FOLDERID_AppUpdates => new(0xa305ce99, 0xf527, 0x492b, 0x8b, 0x1a, 0x7e, 0x76, 0xfa, 0x98, 0xd6, 0xe4);
    
    public static Guid FOLDERID_CameraRoll => new(0xab5fb87b, 0x7ce2, 0x4f83, 0x91, 0x5d, 0x55, 0x08, 0x46, 0xc9, 0x53, 0x7b);
    
    public static Guid FOLDERID_CameraRollLibrary => new(0x2b20df75, 0x1eda, 0x4039, 0x80, 0x97, 0x38, 0x79, 0x82, 0x27, 0xd5, 0xb7);
    
    public static Guid FOLDERID_CDBurning => new(0x9e52ab10, 0xf80d, 0x49df, 0xac, 0xb8, 0x43, 0x30, 0xf5, 0x68, 0x78, 0x55);
    
    public static Guid FOLDERID_ChangeRemovePrograms => new(0xdf7266ac, 0x9274, 0x4867, 0x8d, 0x55, 0x3b, 0xd6, 0x61, 0xde, 0x87, 0x2d);
    
    public static Guid FOLDERID_CommonAdminTools => new(0xd0384e7d, 0xbac3, 0x4797, 0x8f, 0x14, 0xcb, 0xa2, 0x29, 0xb3, 0x92, 0xb5);
    
    public static Guid FOLDERID_CommonOEMLinks => new(0xc1bae2d0, 0x10df, 0x4334, 0xbe, 0xdd, 0x7a, 0xa2, 0x0b, 0x22, 0x7a, 0x9d);
    
    public static Guid FOLDERID_CommonPrograms => new(0x0139d44e, 0x6afe, 0x49f2, 0x86, 0x90, 0x3d, 0xaf, 0xca, 0xe6, 0xff, 0xb8);
    
    public static Guid FOLDERID_CommonStartMenu => new(0xa4115719, 0xd62e, 0x491d, 0xaa, 0x7c, 0xe7, 0x4b, 0x8b, 0xe3, 0xb0, 0x67);
    
    public static Guid FOLDERID_CommonStartMenuPlaces => new(0xa440879f, 0x87a0, 0x4f7d, 0xb7, 0x00, 0x02, 0x07, 0xb9, 0x66, 0x19, 0x4a);
    
    public static Guid FOLDERID_CommonStartup => new(0x82a5ea35, 0xd9cd, 0x47c5, 0x96, 0x29, 0xe1, 0x5d, 0x2f, 0x71, 0x4e, 0x6e);
    
    public static Guid FOLDERID_CommonTemplates => new(0xb94237e7, 0x57ac, 0x4347, 0x91, 0x51, 0xb0, 0x8c, 0x6c, 0x32, 0xd1, 0xf7);
    
    public static Guid FOLDERID_ComputerFolder => new(0x0ac0837c, 0xbbf8, 0x452a, 0x85, 0x0d, 0x79, 0xd0, 0x8e, 0x66, 0x7c, 0xa7);
    
    public static Guid FOLDERID_ConflictFolder => new(0x4bfefb45, 0x347d, 0x4006, 0xa5, 0xbe, 0xac, 0x0c, 0xb0, 0x56, 0x71, 0x92);
    
    public static Guid FOLDERID_ConnectionsFolder => new(0x6f0cd92b, 0x2e97, 0x45d1, 0x88, 0xff, 0xb0, 0xd1, 0x86, 0xb8, 0xde, 0xdd);
    
    public static Guid FOLDERID_Contacts => new(0x56784854, 0xc6cb, 0x462b, 0x81, 0x69, 0x88, 0xe3, 0x50, 0xac, 0xb8, 0x82);
    
    public static Guid FOLDERID_ControlPanelFolder => new(0x82a74aeb, 0xaeb4, 0x465c, 0xa0, 0x14, 0xd0, 0x97, 0xee, 0x34, 0x6d, 0x63);
    
    public static Guid FOLDERID_Cookies => new(0x2b0f765d, 0xc0e9, 0x4171, 0x90, 0x8e, 0x08, 0xa6, 0x11, 0xb8, 0x4f, 0xf6);
    
    public static Guid FOLDERID_CurrentAppMods => new(0x3db40b20, 0x2a30, 0x4dbe, 0x91, 0x7e, 0x77, 0x1d, 0xd2, 0x1d, 0xd0, 0x99);
    
    public static Guid FOLDERID_Desktop => new(0xb4bfcc3a, 0xdb2c, 0x424c, 0xb0, 0x29, 0x7f, 0xe9, 0x9a, 0x87, 0xc6, 0x41);
    
    public static Guid FOLDERID_DevelopmentFiles => new(0xdbe8e08e, 0x3053, 0x4bbc, 0xb1, 0x83, 0x2a, 0x7b, 0x2b, 0x19, 0x1e, 0x59);
    
    public static Guid FOLDERID_Device => new(0x1c2ac1dc, 0x4358, 0x4b6c, 0x97, 0x33, 0xaf, 0x21, 0x15, 0x65, 0x76, 0xf0);
    
    public static Guid FOLDERID_DeviceMetadataStore => new(0x5ce4a5e9, 0xe4eb, 0x479d, 0xb8, 0x9f, 0x13, 0x0c, 0x02, 0x88, 0x61, 0x55);
    
    public static Guid FOLDERID_Documents => new(0xfdd39ad0, 0x238f, 0x46af, 0xad, 0xb4, 0x6c, 0x85, 0x48, 0x03, 0x69, 0xc7);
    
    public static Guid FOLDERID_DocumentsLibrary => new(0x7b0db17d, 0x9cd2, 0x4a93, 0x97, 0x33, 0x46, 0xcc, 0x89, 0x02, 0x2e, 0x7c);
    
    public static Guid FOLDERID_Downloads => new(0x374de290, 0x123f, 0x4565, 0x91, 0x64, 0x39, 0xc4, 0x92, 0x5e, 0x46, 0x7b);
    
    public static Guid FOLDERID_Favorites => new(0x1777f761, 0x68ad, 0x4d8a, 0x87, 0xbd, 0x30, 0xb7, 0x59, 0xfa, 0x33, 0xdd);
    
    public static Guid FOLDERID_Fonts => new(0xfd228cb7, 0xae11, 0x4ae3, 0x86, 0x4c, 0x16, 0xf3, 0x91, 0x0a, 0xb8, 0xfe);
    
    public static Guid FOLDERID_Games => new(0xcac52c1a, 0xb53d, 0x4edc, 0x92, 0xd7, 0x6b, 0x2e, 0x8a, 0xc1, 0x94, 0x34);
    
    public static Guid FOLDERID_GameTasks => new(0x054fae61, 0x4dd8, 0x4787, 0x80, 0xb6, 0x09, 0x02, 0x20, 0xc4, 0xb7, 0x00);
    
    public static Guid FOLDERID_History => new(0xd9dc8a3b, 0xb784, 0x432e, 0xa7, 0x81, 0x5a, 0x11, 0x30, 0xa7, 0x59, 0x63);
    
    public static Guid FOLDERID_HomeGroup => new(0x52528a6b, 0xb9e3, 0x4add, 0xb6, 0x0d, 0x58, 0x8c, 0x2d, 0xba, 0x84, 0x2d);
    
    public static Guid FOLDERID_HomeGroupCurrentUser => new(0x9b74b6a3, 0x0dfd, 0x4f11, 0x9e, 0x78, 0x5f, 0x78, 0x00, 0xf2, 0xe7, 0x72);
    
    public static Guid FOLDERID_ImplicitAppShortcuts => new(0xbcb5256f, 0x79f6, 0x4cee, 0xb7, 0x25, 0xdc, 0x34, 0xe4, 0x02, 0xfd, 0x46);
    
    public static Guid FOLDERID_InternetCache => new(0x352481e8, 0x33be, 0x4251, 0xba, 0x85, 0x60, 0x07, 0xca, 0xed, 0xcf, 0x9d);
    
    public static Guid FOLDERID_InternetFolder => new(0x4d9f7874, 0x4e0c, 0x4904, 0x96, 0x7b, 0x40, 0xb0, 0xd2, 0x0c, 0x3e, 0x4b);
    
    public static Guid FOLDERID_Libraries => new(0x1b3ea5dc, 0xb587, 0x4786, 0xb4, 0xef, 0xbd, 0x1d, 0xc3, 0x32, 0xae, 0xae);
    
    public static Guid FOLDERID_Links => new(0xbfb9d5e0, 0xc6a9, 0x404c, 0xb2, 0xb2, 0xae, 0x6d, 0xb6, 0xaf, 0x49, 0x68);
    
    public static Guid FOLDERID_LocalAppData => new(0xf1b32785, 0x6fba, 0x4fcf, 0x9d, 0x55, 0x7b, 0x8e, 0x7f, 0x15, 0x70, 0x91);
    
    public static Guid FOLDERID_LocalAppDataLow => new(0xa520a1a4, 0x1780, 0x4ff6, 0xbd, 0x18, 0x16, 0x73, 0x43, 0xc5, 0xaf, 0x16);
    
    public static Guid FOLDERID_LocalDocuments => new(0xf42ee2d3, 0x909f, 0x4907, 0x88, 0x71, 0x4c, 0x22, 0xfc, 0x0b, 0xf7, 0x56);
    
    public static Guid FOLDERID_LocalDownloads => new(0x7d83ee9b, 0x2244, 0x4e70, 0xb1, 0xf5, 0x53, 0x93, 0x04, 0x2a, 0xf1, 0xe4);
    
    public static Guid FOLDERID_LocalizedResourcesDir => new(0x2a00375e, 0x224c, 0x49de, 0xb8, 0xd1, 0x44, 0x0d, 0xf7, 0xef, 0x3d, 0xdc);
    
    public static Guid FOLDERID_LocalMusic => new(0xa0c69a99, 0x21c8, 0x4671, 0x87, 0x03, 0x79, 0x34, 0x16, 0x2f, 0xcf, 0x1d);
    
    public static Guid FOLDERID_LocalPictures => new(0x0ddd015d, 0xb06c, 0x45d5, 0x8c, 0x4c, 0xf5, 0x97, 0x13, 0x85, 0x46, 0x39);
    
    public static Guid FOLDERID_LocalStorage => new(0xb3eb08d3, 0xa1f3, 0x496b, 0x86, 0x5a, 0x42, 0xb5, 0x36, 0xcd, 0xa0, 0xec);
    
    public static Guid FOLDERID_LocalVideos => new(0x35286a68, 0x3c57, 0x41a1, 0xbb, 0xb1, 0x0e, 0xae, 0x73, 0xd7, 0x6c, 0x95);
    
    public static Guid FOLDERID_Music => new(0x4bd8d571, 0x6d19, 0x48d3, 0xbe, 0x97, 0x42, 0x22, 0x20, 0x08, 0x0e, 0x43);
    
    public static Guid FOLDERID_MusicLibrary => new(0x2112ab0a, 0xc86a, 0x4ffe, 0xa3, 0x68, 0x0d, 0xe9, 0x6e, 0x47, 0x01, 0x2e);
    
    public static Guid FOLDERID_NetHood => new(0xc5abbf53, 0xe17f, 0x4121, 0x89, 0x00, 0x86, 0x62, 0x6f, 0xc2, 0xc9, 0x73);
    
    public static Guid FOLDERID_NetworkFolder => new(0xd20beec4, 0x5ca8, 0x4905, 0xae, 0x3b, 0xbf, 0x25, 0x1e, 0xa0, 0x9b, 0x53);
    
    public static Guid FOLDERID_Objects3D => new(0x31c0dd25, 0x9439, 0x4f12, 0xbf, 0x41, 0x7f, 0xf4, 0xed, 0xa3, 0x87, 0x22);
    
    public static Guid FOLDERID_OneDrive => new(0xa52bba46, 0xe9e1, 0x435f, 0xb3, 0xd9, 0x28, 0xda, 0xa6, 0x48, 0xc0, 0xf6);
    
    public static Guid FOLDERID_OriginalImages => new(0x2c36c0aa, 0x5812, 0x4b87, 0xbf, 0xd0, 0x4c, 0xd0, 0xdf, 0xb1, 0x9b, 0x39);
    
    public static Guid FOLDERID_PhotoAlbums => new(0x69d2cf90, 0xfc33, 0x4fb7, 0x9a, 0x0c, 0xeb, 0xb0, 0xf0, 0xfc, 0xb4, 0x3c);
    
    public static Guid FOLDERID_Pictures => new(0x33e28130, 0x4e1e, 0x4676, 0x83, 0x5a, 0x98, 0x39, 0x5c, 0x3b, 0xc3, 0xbb);
    
    public static Guid FOLDERID_PicturesLibrary => new(0xa990ae9f, 0xa03b, 0x4e80, 0x94, 0xbc, 0x99, 0x12, 0xd7, 0x50, 0x41, 0x04);
    
    public static Guid FOLDERID_Playlists => new(0xde92c1c7, 0x837f, 0x4f69, 0xa3, 0xbb, 0x86, 0xe6, 0x31, 0x20, 0x4a, 0x23);
    
    public static Guid FOLDERID_PrintersFolder => new(0x76fc4e2d, 0xd6ad, 0x4519, 0xa6, 0x63, 0x37, 0xbd, 0x56, 0x06, 0x81, 0x85);
    
    public static Guid FOLDERID_PrintHood => new(0x9274bd8d, 0xcfd1, 0x41c3, 0xb3, 0x5e, 0xb1, 0x3f, 0x55, 0xa7, 0x58, 0xf4);
    
    public static Guid FOLDERID_Profile => new(0x5e6c858f, 0x0e22, 0x4760, 0x9a, 0xfe, 0xea, 0x33, 0x17, 0xb6, 0x71, 0x73);
    
    public static Guid FOLDERID_ProgramData => new(0x62ab5d82, 0xfdc1, 0x4dc3, 0xa9, 0xdd, 0x07, 0x0d, 0x1d, 0x49, 0x5d, 0x97);
    
    public static Guid FOLDERID_ProgramFiles => new(0x905e63b6, 0xc1bf, 0x494e, 0xb2, 0x9c, 0x65, 0xb7, 0x32, 0xd3, 0xd2, 0x1a);
    
    public static Guid FOLDERID_ProgramFilesCommon => new(0xf7f1ed05, 0x9f6d, 0x47a2, 0xaa, 0xae, 0x29, 0xd3, 0x17, 0xc6, 0xf0, 0x66);
    
    public static Guid FOLDERID_ProgramFilesCommonX64 => new(0x6365d5a7, 0x0f0d, 0x45e5, 0x87, 0xf6, 0x0d, 0xa5, 0x6b, 0x6a, 0x4f, 0x7d);
    
    public static Guid FOLDERID_ProgramFilesCommonX86 => new(0xde974d24, 0xd9c6, 0x4d3e, 0xbf, 0x91, 0xf4, 0x45, 0x51, 0x20, 0xb9, 0x17);
    
    public static Guid FOLDERID_ProgramFilesX64 => new(0x6d809377, 0x6af0, 0x444b, 0x89, 0x57, 0xa3, 0x77, 0x3f, 0x02, 0x20, 0x0e);
    
    public static Guid FOLDERID_ProgramFilesX86 => new(0x7c5a40ef, 0xa0fb, 0x4bfc, 0x87, 0x4a, 0xc0, 0xf2, 0xe0, 0xb9, 0xfa, 0x8e);
    
    public static Guid FOLDERID_Programs => new(0xa77f5d77, 0x2e2b, 0x44c3, 0xa6, 0xa2, 0xab, 0xa6, 0x01, 0x05, 0x4a, 0x51);
    
    public static Guid FOLDERID_Public => new(0xdfdf76a2, 0xc82a, 0x4d63, 0x90, 0x6a, 0x56, 0x44, 0xac, 0x45, 0x73, 0x85);
    
    public static Guid FOLDERID_PublicDesktop => new(0xc4aa340d, 0xf20f, 0x4863, 0xaf, 0xef, 0xf8, 0x7e, 0xf2, 0xe6, 0xba, 0x25);
    
    public static Guid FOLDERID_PublicDocuments => new(0xed4824af, 0xdce4, 0x45a8, 0x81, 0xe2, 0xfc, 0x79, 0x65, 0x08, 0x36, 0x34);
    
    public static Guid FOLDERID_PublicDownloads => new(0x3d644c9b, 0x1fb8, 0x4f30, 0x9b, 0x45, 0xf6, 0x70, 0x23, 0x5f, 0x79, 0xc0);
    
    public static Guid FOLDERID_PublicGameTasks => new(0xdebf2536, 0xe1a8, 0x4c59, 0xb6, 0xa2, 0x41, 0x45, 0x86, 0x47, 0x6a, 0xea);
    
    public static Guid FOLDERID_PublicLibraries => new(0x48daf80b, 0xe6cf, 0x4f4e, 0xb8, 0x00, 0x0e, 0x69, 0xd8, 0x4e, 0xe3, 0x84);
    
    public static Guid FOLDERID_PublicMusic => new(0x3214fab5, 0x9757, 0x4298, 0xbb, 0x61, 0x92, 0xa9, 0xde, 0xaa, 0x44, 0xff);
    
    public static Guid FOLDERID_PublicPictures => new(0xb6ebfb86, 0x6907, 0x413c, 0x9a, 0xf7, 0x4f, 0xc2, 0xab, 0xf0, 0x7c, 0xc5);
    
    public static Guid FOLDERID_PublicRingtones => new(0xe555ab60, 0x153b, 0x4d17, 0x9f, 0x04, 0xa5, 0xfe, 0x99, 0xfc, 0x15, 0xec);
    
    public static Guid FOLDERID_PublicUserTiles => new(0x0482af6c, 0x08f1, 0x4c34, 0x8c, 0x90, 0xe1, 0x7e, 0xc9, 0x8b, 0x1e, 0x17);
    
    public static Guid FOLDERID_PublicVideos => new(0x2400183a, 0x6185, 0x49fb, 0xa2, 0xd8, 0x4a, 0x39, 0x2a, 0x60, 0x2b, 0xa3);
    
    public static Guid FOLDERID_QuickLaunch => new(0x52a4f021, 0x7b75, 0x48a9, 0x9f, 0x6b, 0x4b, 0x87, 0xa2, 0x10, 0xbc, 0x8f);
    
    public static Guid FOLDERID_Recent => new(0xae50c081, 0xebd2, 0x438a, 0x86, 0x55, 0x8a, 0x09, 0x2e, 0x34, 0x98, 0x7a);
    
    public static Guid FOLDERID_RecordedCalls => new(0x2f8b40c2, 0x83ed, 0x48ee, 0xb3, 0x83, 0xa1, 0xf1, 0x57, 0xec, 0x6f, 0x9a);
    
    public static Guid FOLDERID_RecordedTVLibrary => new(0x1a6fdba2, 0xf42d, 0x4358, 0xa7, 0x98, 0xb7, 0x4d, 0x74, 0x59, 0x26, 0xc5);
    
    public static Guid FOLDERID_RecycleBinFolder => new(0xb7534046, 0x3ecb, 0x4c18, 0xbe, 0x4e, 0x64, 0xcd, 0x4c, 0xb7, 0xd6, 0xac);
    
    public static Guid FOLDERID_ResourceDir => new(0x8ad10c31, 0x2adb, 0x4296, 0xa8, 0xf7, 0xe4, 0x70, 0x12, 0x32, 0xc9, 0x72);
    
    public static Guid FOLDERID_RetailDemo => new(0x12d4c69e, 0x24ad, 0x4923, 0xbe, 0x19, 0x31, 0x32, 0x1c, 0x43, 0xa7, 0x67);
    
    public static Guid FOLDERID_Ringtones => new(0xc870044b, 0xf49e, 0x4126, 0xa9, 0xc3, 0xb5, 0x2a, 0x1f, 0xf4, 0x11, 0xe8);
    
    public static Guid FOLDERID_RoamedTileImages => new(0xaaa8d5a5, 0xf1d6, 0x4259, 0xba, 0xa8, 0x78, 0xe7, 0xef, 0x60, 0x83, 0x5e);
    
    public static Guid FOLDERID_RoamingAppData => new(0x3eb685db, 0x65f9, 0x4cf6, 0xa0, 0x3a, 0xe3, 0xef, 0x65, 0x72, 0x9f, 0x3d);
    
    public static Guid FOLDERID_RoamingTiles => new(0x00bcfc5a, 0xed94, 0x4e48, 0x96, 0xa1, 0x3f, 0x62, 0x17, 0xf2, 0x19, 0x90);
    
    public static Guid FOLDERID_SampleMusic => new(0xb250c668, 0xf57d, 0x4ee1, 0xa6, 0x3c, 0x29, 0x0e, 0xe7, 0xd1, 0xaa, 0x1f);
    
    public static Guid FOLDERID_SamplePictures => new(0xc4900540, 0x2379, 0x4c75, 0x84, 0x4b, 0x64, 0xe6, 0xfa, 0xf8, 0x71, 0x6b);
    
    public static Guid FOLDERID_SamplePlaylists => new(0x15ca69b3, 0x30ee, 0x49c1, 0xac, 0xe1, 0x6b, 0x5e, 0xc3, 0x72, 0xaf, 0xb5);
    
    public static Guid FOLDERID_SampleVideos => new(0x859ead94, 0x2e85, 0x48ad, 0xa7, 0x1a, 0x09, 0x69, 0xcb, 0x56, 0xa6, 0xcd);
    
    public static Guid FOLDERID_SavedGames => new(0x4c5c32ff, 0xbb9d, 0x43b0, 0xb5, 0xb4, 0x2d, 0x72, 0xe5, 0x4e, 0xaa, 0xa4);
    
    public static Guid FOLDERID_SavedPictures => new(0x3b193882, 0xd3ad, 0x4eab, 0x96, 0x5a, 0x69, 0x82, 0x9d, 0x1f, 0xb5, 0x9f);
    
    public static Guid FOLDERID_SavedPicturesLibrary => new(0xe25b5812, 0xbe88, 0x4bd9, 0x94, 0xb0, 0x29, 0x23, 0x34, 0x77, 0xb6, 0xc3);
    
    public static Guid FOLDERID_SavedSearches => new(0x7d1d3a04, 0xdebb, 0x4115, 0x95, 0xcf, 0x2f, 0x29, 0xda, 0x29, 0x20, 0xda);
    
    public static Guid FOLDERID_Screenshots => new(0xb7bede81, 0xdf94, 0x4682, 0xa7, 0xd8, 0x57, 0xa5, 0x26, 0x20, 0xb8, 0x6f);
    
    public static Guid FOLDERID_SEARCH_CSC => new(0xee32e446, 0x31ca, 0x4aba, 0x81, 0x4f, 0xa5, 0xeb, 0xd2, 0xfd, 0x6d, 0x5e);
    
    public static Guid FOLDERID_SEARCH_MAPI => new(0x98ec0e18, 0x2098, 0x4d44, 0x86, 0x44, 0x66, 0x97, 0x93, 0x15, 0xa2, 0x81);
    
    public static Guid FOLDERID_SearchHistory => new(0x0d4c3db6, 0x03a3, 0x462f, 0xa0, 0xe6, 0x08, 0x92, 0x4c, 0x41, 0xb5, 0xd4);
    
    public static Guid FOLDERID_SearchHome => new(0x190337d1, 0xb8ca, 0x4121, 0xa6, 0x39, 0x6d, 0x47, 0x2d, 0x16, 0x97, 0x2a);
    
    public static Guid FOLDERID_SearchTemplates => new(0x7e636bfe, 0xdfa9, 0x4d5e, 0xb4, 0x56, 0xd7, 0xb3, 0x98, 0x51, 0xd8, 0xa9);
    
    public static Guid FOLDERID_SendTo => new(0x8983036c, 0x27c0, 0x404b, 0x8f, 0x08, 0x10, 0x2d, 0x10, 0xdc, 0xfd, 0x74);
    
    public static Guid FOLDERID_SidebarDefaultParts => new(0x7b396e54, 0x9ec5, 0x4300, 0xbe, 0x0a, 0x24, 0x82, 0xeb, 0xae, 0x1a, 0x26);
    
    public static Guid FOLDERID_SidebarParts => new(0xa75d362e, 0x50fc, 0x4fb7, 0xac, 0x2c, 0xa8, 0xbe, 0xaa, 0x31, 0x44, 0x93);
    
    public static Guid FOLDERID_SkyDrive => new(0xa52bba46, 0xe9e1, 0x435f, 0xb3, 0xd9, 0x28, 0xda, 0xa6, 0x48, 0xc0, 0xf6);
    
    public static Guid FOLDERID_SkyDriveCameraRoll => new(0x767e6811, 0x49cb, 0x4273, 0x87, 0xc2, 0x20, 0xf3, 0x55, 0xe1, 0x08, 0x5b);
    
    public static Guid FOLDERID_SkyDriveDocuments => new(0x24d89e24, 0x2f19, 0x4534, 0x9d, 0xde, 0x6a, 0x66, 0x71, 0xfb, 0xb8, 0xfe);
    
    public static Guid FOLDERID_SkyDriveMusic => new(0xc3f2459e, 0x80d6, 0x45dc, 0xbf, 0xef, 0x1f, 0x76, 0x9f, 0x2b, 0xe7, 0x30);
    
    public static Guid FOLDERID_SkyDrivePictures => new(0x339719b5, 0x8c47, 0x4894, 0x94, 0xc2, 0xd8, 0xf7, 0x7a, 0xdd, 0x44, 0xa6);
    
    public static Guid FOLDERID_StartMenu => new(0x625b53c3, 0xab48, 0x4ec1, 0xba, 0x1f, 0xa1, 0xef, 0x41, 0x46, 0xfc, 0x19);
    
    public static Guid FOLDERID_StartMenuAllPrograms => new(0xf26305ef, 0x6948, 0x40b9, 0xb2, 0x55, 0x81, 0x45, 0x3d, 0x09, 0xc7, 0x85);
    
    public static Guid FOLDERID_Startup => new(0xb97d20bb, 0xf46a, 0x4c97, 0xba, 0x10, 0x5e, 0x36, 0x08, 0x43, 0x08, 0x54);
    
    public static Guid FOLDERID_SyncManagerFolder => new(0x43668bf8, 0xc14e, 0x49b2, 0x97, 0xc9, 0x74, 0x77, 0x84, 0xd7, 0x84, 0xb7);
    
    public static Guid FOLDERID_SyncResultsFolder => new(0x289a9a43, 0xbe44, 0x4057, 0xa4, 0x1b, 0x58, 0x7a, 0x76, 0xd7, 0xe7, 0xf9);
    
    public static Guid FOLDERID_SyncSetupFolder => new(0x0f214138, 0xb1d3, 0x4a90, 0xbb, 0xa9, 0x27, 0xcb, 0xc0, 0xc5, 0x38, 0x9a);
    
    public static Guid FOLDERID_System => new(0x1ac14e77, 0x02e7, 0x4e5d, 0xb7, 0x44, 0x2e, 0xb1, 0xae, 0x51, 0x98, 0xb7);
    
    public static Guid FOLDERID_SystemX86 => new(0xd65231b0, 0xb2f1, 0x4857, 0xa4, 0xce, 0xa8, 0xe7, 0xc6, 0xea, 0x7d, 0x27);
    
    public static Guid FOLDERID_Templates => new(0xa63293e8, 0x664e, 0x48db, 0xa0, 0x79, 0xdf, 0x75, 0x9e, 0x05, 0x09, 0xf7);
    
    public static Guid FOLDERID_UserPinned => new(0x9e3995ab, 0x1f9c, 0x4f13, 0xb8, 0x27, 0x48, 0xb2, 0x4b, 0x6c, 0x71, 0x74);
    
    public static Guid FOLDERID_UserProfiles => new(0x0762d272, 0xc50a, 0x4bb0, 0xa3, 0x82, 0x69, 0x7d, 0xcd, 0x72, 0x9b, 0x80);
    
    public static Guid FOLDERID_UserProgramFiles => new(0x5cd7aee2, 0x2219, 0x4a67, 0xb8, 0x5d, 0x6c, 0x9c, 0xe1, 0x56, 0x60, 0xcb);
    
    public static Guid FOLDERID_UserProgramFilesCommon => new(0xbcbd3057, 0xca5c, 0x4622, 0xb4, 0x2d, 0xbc, 0x56, 0xdb, 0x0a, 0xe5, 0x16);
    
    public static Guid FOLDERID_UsersFiles => new(0xf3ce0f7c, 0x4901, 0x4acc, 0x86, 0x48, 0xd5, 0xd4, 0x4b, 0x04, 0xef, 0x8f);
    
    public static Guid FOLDERID_UsersLibraries => new(0xa302545d, 0xdeff, 0x464b, 0xab, 0xe8, 0x61, 0xc8, 0x64, 0x8d, 0x93, 0x9b);
    
    public static Guid FOLDERID_Videos => new(0x18989b1d, 0x99b5, 0x455b, 0x84, 0x1c, 0xab, 0x7c, 0x74, 0xe4, 0xdd, 0xfc);
    
    public static Guid FOLDERID_VideosLibrary => new(0x491e922f, 0x5643, 0x4af4, 0xa7, 0xeb, 0x4e, 0x7a, 0x13, 0x8d, 0x81, 0x74);
    
    public static Guid FOLDERID_Windows => new(0xf38bf404, 0x1d43, 0x42f2, 0x93, 0x05, 0x67, 0xde, 0x0b, 0x28, 0xfc, 0x23);
    
    public static Guid FOLDERTYPEID_AccountPictures => new(0xdb2a5d8f, 0x06e6, 0x4007, 0xab, 0xa6, 0xaf, 0x87, 0x7d, 0x52, 0x6e, 0xa6);
    
    public static Guid FOLDERTYPEID_Communications => new(0x91475fe5, 0x586b, 0x4eba, 0x8d, 0x75, 0xd1, 0x74, 0x34, 0xb8, 0xcd, 0xf6);
    
    public static Guid FOLDERTYPEID_CompressedFolder => new(0x80213e82, 0xbcfd, 0x4c4f, 0x88, 0x17, 0xbb, 0x27, 0x60, 0x12, 0x67, 0xa9);
    
    public static Guid FOLDERTYPEID_Contacts => new(0xde2b70ec, 0x9bf7, 0x4a93, 0xbd, 0x3d, 0x24, 0x3f, 0x78, 0x81, 0xd4, 0x92);
    
    public static Guid FOLDERTYPEID_ControlPanelCategory => new(0xde4f0660, 0xfa10, 0x4b8f, 0xa4, 0x94, 0x06, 0x8b, 0x20, 0xb2, 0x23, 0x07);
    
    public static Guid FOLDERTYPEID_ControlPanelClassic => new(0x0c3794f3, 0xb545, 0x43aa, 0xa3, 0x29, 0xc3, 0x74, 0x30, 0xc5, 0x8d, 0x2a);
    
    public static Guid FOLDERTYPEID_Documents => new(0x7d49d726, 0x3c21, 0x4f05, 0x99, 0xaa, 0xfd, 0xc2, 0xc9, 0x47, 0x46, 0x56);
    
    public static Guid FOLDERTYPEID_Downloads => new(0x885a186e, 0xa440, 0x4ada, 0x81, 0x2b, 0xdb, 0x87, 0x1b, 0x94, 0x22, 0x59);
    
    public static Guid FOLDERTYPEID_Games => new(0xb689b0d0, 0x76d3, 0x4cbb, 0x87, 0xf7, 0x58, 0x5d, 0x0e, 0x0c, 0xe0, 0x70);
    
    public static Guid FOLDERTYPEID_Generic => new(0x5c4f28b5, 0xf869, 0x4e84, 0x8e, 0x60, 0xf1, 0x1d, 0xb9, 0x7c, 0x5c, 0xc7);
    
    public static Guid FOLDERTYPEID_GenericLibrary => new(0x5f4eab9a, 0x6833, 0x4f61, 0x89, 0x9d, 0x31, 0xcf, 0x46, 0x97, 0x9d, 0x49);
    
    public static Guid FOLDERTYPEID_GenericSearchResults => new(0x7fde1a1e, 0x8b31, 0x49a5, 0x93, 0xb8, 0x6b, 0xe1, 0x4c, 0xfa, 0x49, 0x43);
    
    public static Guid FOLDERTYPEID_Invalid => new(0x57807898, 0x8c4f, 0x4462, 0xbb, 0x63, 0x71, 0x04, 0x23, 0x80, 0xb1, 0x09);
    
    public static Guid FOLDERTYPEID_Music => new(0x94d6ddcc, 0x4a68, 0x4175, 0xa3, 0x74, 0xbd, 0x58, 0x4a, 0x51, 0x0b, 0x78);
    
    public static Guid FOLDERTYPEID_NetworkExplorer => new(0x25cc242b, 0x9a7c, 0x4f51, 0x80, 0xe0, 0x7a, 0x29, 0x28, 0xfe, 0xbe, 0x42);
    
    public static Guid FOLDERTYPEID_OpenSearch => new(0x8faf9629, 0x1980, 0x46ff, 0x80, 0x23, 0x9d, 0xce, 0xab, 0x9c, 0x3e, 0xe3);
    
    public static Guid FOLDERTYPEID_OtherUsers => new(0xb337fd00, 0x9dd5, 0x4635, 0xa6, 0xd4, 0xda, 0x33, 0xfd, 0x10, 0x2b, 0x7a);
    
    public static Guid FOLDERTYPEID_Pictures => new(0xb3690e58, 0xe961, 0x423b, 0xb6, 0x87, 0x38, 0x6e, 0xbf, 0xd8, 0x32, 0x39);
    
    public static Guid FOLDERTYPEID_Printers => new(0x2c7bbec6, 0xc844, 0x4a0a, 0x91, 0xfa, 0xce, 0xf6, 0xf5, 0x9c, 0xfd, 0xa1);
    
    public static Guid FOLDERTYPEID_PublishedItems => new(0x7f2f5b96, 0xff74, 0x41da, 0xaf, 0xd8, 0x1c, 0x78, 0xa5, 0xf3, 0xae, 0xa2);
    
    public static Guid FOLDERTYPEID_RecordedTV => new(0x5557a28f, 0x5da6, 0x4f83, 0x88, 0x09, 0xc2, 0xc9, 0x8a, 0x11, 0xa6, 0xfa);
    
    public static Guid FOLDERTYPEID_RecycleBin => new(0xd6d9e004, 0xcd87, 0x442b, 0x9d, 0x57, 0x5e, 0x0a, 0xeb, 0x4f, 0x6f, 0x72);
    
    public static Guid FOLDERTYPEID_SavedGames => new(0xd0363307, 0x28cb, 0x4106, 0x9f, 0x23, 0x29, 0x56, 0xe3, 0xe5, 0xe0, 0xe7);
    
    public static Guid FOLDERTYPEID_SearchConnector => new(0x982725ee, 0x6f47, 0x479e, 0xb4, 0x47, 0x81, 0x2b, 0xfa, 0x7d, 0x2e, 0x8f);
    
    public static Guid FOLDERTYPEID_Searches => new(0x0b0ba2e3, 0x405f, 0x415e, 0xa6, 0xee, 0xca, 0xd6, 0x25, 0x20, 0x78, 0x53);
    
    public static Guid FOLDERTYPEID_SearchHome => new(0x834d8a44, 0x0974, 0x4ed6, 0x86, 0x6e, 0xf2, 0x03, 0xd8, 0x0b, 0x38, 0x10);
    
    public static Guid FOLDERTYPEID_SoftwareExplorer => new(0xd674391b, 0x52d9, 0x4e07, 0x83, 0x4e, 0x67, 0xc9, 0x86, 0x10, 0xf3, 0x9d);
    
    public static Guid FOLDERTYPEID_StartMenu => new(0xef87b4cb, 0xf2ce, 0x4785, 0x86, 0x58, 0x4c, 0xa6, 0xc6, 0x3e, 0x38, 0xc6);
    
    public static Guid FOLDERTYPEID_StorageProviderDocuments => new(0xdd61bd66, 0x70e8, 0x48dd, 0x96, 0x55, 0x65, 0xc5, 0xe1, 0xaa, 0xc2, 0xd1);
    
    public static Guid FOLDERTYPEID_StorageProviderGeneric => new(0x4f01ebc5, 0x2385, 0x41f2, 0xa2, 0x8e, 0x2c, 0x5c, 0x91, 0xfb, 0x56, 0xe0);
    
    public static Guid FOLDERTYPEID_StorageProviderMusic => new(0x672ecd7e, 0xaf04, 0x4399, 0x87, 0x5c, 0x02, 0x90, 0x84, 0x5b, 0x62, 0x47);
    
    public static Guid FOLDERTYPEID_StorageProviderPictures => new(0x71d642a9, 0xf2b1, 0x42cd, 0xad, 0x92, 0xeb, 0x93, 0x00, 0xc7, 0xcc, 0x0a);
    
    public static Guid FOLDERTYPEID_StorageProviderVideos => new(0x51294da1, 0xd7b1, 0x485b, 0x9e, 0x9a, 0x17, 0xcf, 0xfe, 0x33, 0xe1, 0x87);
    
    public static Guid FOLDERTYPEID_UserFiles => new(0xcd0fc69b, 0x71e2, 0x46e5, 0x96, 0x90, 0x5b, 0xcd, 0x9f, 0x57, 0xaa, 0xb3);
    
    public static Guid FOLDERTYPEID_UsersLibraries => new(0xc4d98f09, 0x6124, 0x4fe0, 0x99, 0x42, 0x82, 0x64, 0x16, 0x08, 0x2d, 0xa9);
    
    public static Guid FOLDERTYPEID_VersionControl => new(0x69f1e26b, 0xec64, 0x4280, 0xbc, 0x83, 0xf1, 0xeb, 0x88, 0x7e, 0xc3, 0x5a);
    
    public static Guid FOLDERTYPEID_Videos => new(0x5fa96407, 0x7e77, 0x483c, 0xac, 0x93, 0x69, 0x1d, 0x05, 0x85, 0x0d, 0xe8);
    
    public static Guid FolderViewHost => new(0x20b1cb23, 0x6968, 0x4eb9, 0xb7, 0xd4, 0xa6, 0x6d, 0x00, 0xd0, 0x7c, 0xee);
    
    public static Guid FrameworkInputPane => new(0xd5120aa3, 0x46ba, 0x44c5, 0x82, 0x2d, 0xca, 0x80, 0x92, 0xc1, 0xfc, 0x72);
    
    public static Guid FreeSpaceCategorizer => new(0xb5607793, 0x24ac, 0x44c7, 0x82, 0xe2, 0x83, 0x17, 0x26, 0xaa, 0x6c, 0xb7);
    
    public static Guid FSCopyHandler => new(0xd197380a, 0x0a79, 0x4dc8, 0xa0, 0x33, 0xed, 0x88, 0x2c, 0x2f, 0xa1, 0x4b);
    
    public const uint FVSIF_CANVIEWIT = 1073741824;
    
    public const uint FVSIF_NEWFAILED = 134217728;
    
    public const uint FVSIF_NEWFILE = 2147483648;
    
    public const uint FVSIF_PINNED = 2;
    
    public const uint FVSIF_RECT = 1;
    
    public const uint GADOF_DIRTY = 1;
    
    public const uint GCS_HELPTEXT = 5;
    
    public const uint GCS_HELPTEXTA = 1;
    
    public const uint GCS_HELPTEXTW = 5;
    
    public const uint GCS_UNICODE = 4;
    
    public const uint GCS_VALIDATE = 6;
    
    public const uint GCS_VALIDATEA = 2;
    
    public const uint GCS_VALIDATEW = 6;
    
    public const uint GCS_VERB = 4;
    
    public const uint GCS_VERBA = 0;
    
    public const uint GCS_VERBICONW = 20;
    
    public const uint GCS_VERBW = 4;
    
    public const uint GCT_INVALID = 0;
    
    public const uint GCT_LFNCHAR = 1;
    
    public const uint GCT_SEPARATOR = 8;
    
    public const uint GCT_SHORTCHAR = 2;
    
    public const uint GCT_WILD = 4;
    
    public static Guid GenericCredentialProvider => new(0x25cbb996, 0x92ed, 0x457e, 0xb2, 0x8c, 0x47, 0x74, 0x08, 0x4b, 0xd5, 0x62);
    
    public const uint GETPROPS_NONE = 0;
    
    public const uint GIL_ASYNC = 32;
    
    public const uint GIL_CHECKSHIELD = 512;
    
    public const uint GIL_DEFAULTICON = 64;
    
    public const uint GIL_DONTCACHE = 16;
    
    public const uint GIL_FORCENOSHIELD = 1024;
    
    public const uint GIL_FORSHELL = 2;
    
    public const uint GIL_FORSHORTCUT = 128;
    
    public const uint GIL_NOTFILENAME = 8;
    
    public const uint GIL_OPENICON = 1;
    
    public const uint GIL_PERCLASS = 4;
    
    public const uint GIL_PERINSTANCE = 2;
    
    public const uint GIL_SHIELD = 512;
    
    public const uint GIL_SIMULATEDOC = 1;
    
    public static Guid GUID_DEVINTERFACE_WPD => new(0x6ac27878, 0xa6fa, 0x4155, 0xba, 0x85, 0xf9, 0x8f, 0x49, 0x1d, 0x4f, 0x33);
    
    public static Guid GUID_DEVINTERFACE_WPD_PRIVATE => new(0xba0c718f, 0x4ded, 0x49b7, 0xbd, 0xd3, 0xfa, 0xbe, 0x28, 0x66, 0x12, 0x11);
    
    public static Guid GUID_DEVINTERFACE_WPD_SERVICE => new(0x9ef44f80, 0x3d64, 0x4246, 0xa6, 0xaa, 0x20, 0x6f, 0x32, 0x8d, 0x1e, 0xdc);
    
    public static Guid HideInputPaneAnimationCoordinator => new(0x384742b1, 0x2a77, 0x4cb3, 0x8c, 0xf8, 0x11, 0x36, 0xf5, 0xe1, 0x7e, 0x59);
    
    public const int HLINK_S_DONTHIDE = 262400;
    
    public const uint HLNF_ALLOW_AUTONAVIGATE = 536870912;
    
    public const uint HLNF_CALLERUNTRUSTED = 2097152;
    
    public const uint HLNF_DISABLEWINDOWRESTRICTIONS = 8388608;
    
    public const uint HLNF_EXTERNALNAVIGATE = 268435456;
    
    public const uint HLNF_NEWWINDOWSMANAGED = 2147483648;
    
    public const uint HLNF_TRUSTEDFORACTIVEX = 4194304;
    
    public const uint HLNF_TRUSTFIRSTDOWNLOAD = 16777216;
    
    public const uint HLNF_UNTRUSTEDFORDOWNLOAD = 33554432;
    
    public static Guid HomeGroup => new(0xde77ba04, 0x3c92, 0x4d11, 0xa1, 0xa5, 0x42, 0x35, 0x2a, 0x53, 0xe0, 0xe3);
    
    public const string HOMEGROUP_SECURITY_GROUP = @"HomeUsers";
    
    public const string HOMEGROUP_SECURITY_GROUP_MULTI = @"HUG";
    
    public const uint ID_APP = 100;
    
    public const uint IDC_OFFLINE_HAND = 103;
    
    public const uint IDC_PANTOOL_HAND_CLOSED = 105;
    
    public const uint IDC_PANTOOL_HAND_OPEN = 104;
    
    public const uint IDD_WIZEXTN_FIRST = 20480;
    
    public const uint IDD_WIZEXTN_LAST = 20736;
    
    public static Guid Identity_LocalUserProvider => new(0xa198529b, 0x730f, 0x4089, 0xb6, 0x46, 0xa1, 0x25, 0x57, 0xf5, 0x66, 0x5e);
    
    public const ulong IDO_SHGIOI_DEFAULT = 4294967292;
    
    public const uint IDO_SHGIOI_LINK = 268435454;
    
    public const uint IDO_SHGIOI_SHARE = 268435455;
    
    public const ulong IDO_SHGIOI_SLOWFILE = 4294967293;
    
    public const uint IDS_DESCRIPTION = 1;
    
    public const uint idsAppName = 1007;
    
    public const uint idsBadOldPW = 1006;
    
    public const uint idsChangePW = 1005;
    
    public const uint idsDefKeyword = 1010;
    
    public const uint idsDifferentPW = 1004;
    
    public const uint idsHelpFile = 1009;
    
    public const uint idsIniFile = 1001;
    
    public const uint idsIsPassword = 1000;
    
    public const uint idsNoHelpMemory = 1008;
    
    public const uint idsPassword = 1003;
    
    public const uint idsScreenSaver = 1002;
    
    public const uint IEI_PRIORITY_MAX = 2147483647;
    
    public const uint IEI_PRIORITY_MIN = 0;
    
    public const uint IEIFLAG_ASPECT = 4;
    
    public const uint IEIFLAG_ASYNC = 1;
    
    public const uint IEIFLAG_CACHE = 2;
    
    public const uint IEIFLAG_GLEAM = 16;
    
    public const uint IEIFLAG_NOBORDER = 256;
    
    public const uint IEIFLAG_NOSTAMP = 128;
    
    public const uint IEIFLAG_OFFLINE = 8;
    
    public const uint IEIFLAG_ORIGSIZE = 64;
    
    public const uint IEIFLAG_QUALITY = 512;
    
    public const uint IEIFLAG_REFRESH = 1024;
    
    public const uint IEIFLAG_SCREEN = 32;
    
    public const uint IEIT_PRIORITY_NORMAL = 268435456;
    
    public static Guid IENamespaceTreeControl => new(0xace52d03, 0xe5cd, 0x4b20, 0x82, 0xff, 0xe7, 0x1b, 0x11, 0xbe, 0xae, 0x1d);
    
    public const uint ILMM_IE4 = 0;
    
    public static Guid ImageProperties => new(0x7ab770c7, 0x0e23, 0x4d7a, 0x8a, 0xa2, 0x19, 0xbf, 0xad, 0x47, 0x98, 0x29);
    
    public static Guid ImageRecompress => new(0x6e33091c, 0xd2f8, 0x4740, 0xb5, 0x5e, 0x2e, 0x11, 0xd1, 0x47, 0x7a, 0x2c);
    
    public static Guid ImageTranscode => new(0x17b75166, 0x928f, 0x417d, 0x96, 0x85, 0x64, 0xaa, 0x13, 0x55, 0x65, 0xc1);
    
    public static Guid InMemoryPropertyStore => new(0x9a02e012, 0x6303, 0x4e1e, 0xb9, 0xa1, 0x63, 0x0f, 0x80, 0x25, 0x92, 0xc5);
    
    public static Guid InMemoryPropertyStoreMarshalByValue => new(0xd4ca0e2d, 0x6da7, 0x4b75, 0xa9, 0x7c, 0x5f, 0x30, 0x6f, 0x0e, 0xae, 0xdc);
    
    public static Guid InputPanelConfiguration => new(0x2853add3, 0xf096, 0x4c63, 0xa7, 0x8f, 0x7f, 0xa3, 0xea, 0x83, 0x7f, 0xb7);
    
    public const uint INTERNET_MAX_PATH_LENGTH = 2048;
    
    public const uint INTERNET_MAX_SCHEME_LENGTH = 32;
    
    public static Guid InternetExplorer => new(0x0002df01, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid InternetExplorerMedium => new(0xd5e8041d, 0x920f, 0x45e9, 0xb8, 0xfb, 0xb1, 0xde, 0xb8, 0x2c, 0x6e, 0x5e);
    
    public static Guid InternetPrintOrdering => new(0xadd36aa8, 0x751a, 0x4579, 0xa2, 0x66, 0xd6, 0x6f, 0x52, 0x02, 0xcc, 0xbb);
    
    public const uint IOCTL_WPD_MESSAGE_READ_ACCESS = 4210952;
    
    public const uint IOCTL_WPD_MESSAGE_READWRITE_ACCESS = 4243720;
    
    public const uint IRTIR_TASK_FINISHED = 4;
    
    public const uint IRTIR_TASK_NOT_RUNNING = 0;
    
    public const uint IRTIR_TASK_PENDING = 3;
    
    public const uint IRTIR_TASK_RUNNING = 1;
    
    public const uint IRTIR_TASK_SUSPENDED = 2;
    
    public const uint IS_FULLSCREEN = 2;
    
    public const uint IS_NORMAL = 1;
    
    public const uint IS_SPLIT = 4;
    
    public const uint ISFB_MASK_BKCOLOR = 2;
    
    public const uint ISFB_MASK_COLORS = 32;
    
    public const uint ISFB_MASK_IDLIST = 16;
    
    public const uint ISFB_MASK_SHELLFOLDER = 8;
    
    public const uint ISFB_MASK_STATE = 1;
    
    public const uint ISFB_MASK_VIEWMODE = 4;
    
    public const uint ISFB_STATE_ALLOWRENAME = 2;
    
    public const uint ISFB_STATE_BTNMINSIZE = 256;
    
    public const uint ISFB_STATE_CHANNELBAR = 16;
    
    public const uint ISFB_STATE_DEBOSSED = 1;
    
    public const uint ISFB_STATE_DEFAULT = 0;
    
    public const uint ISFB_STATE_FULLOPEN = 64;
    
    public const uint ISFB_STATE_NONAMESORT = 128;
    
    public const uint ISFB_STATE_NOSHOWTEXT = 4;
    
    public const uint ISFB_STATE_QLINKSMODE = 32;
    
    public const uint ISFBVIEWMODE_LARGEICONS = 2;
    
    public const uint ISFBVIEWMODE_LOGOS = 3;
    
    public const uint ISFBVIEWMODE_SMALLICONS = 1;
    
    public const int ISHCUTCMDID_COMMITHISTORY = 2;
    
    public const int ISHCUTCMDID_DOWNLOADICON = 0;
    
    public const int ISHCUTCMDID_INTSHORTCUTCREATE = 1;
    
    public const int ISHCUTCMDID_SETUSERAWURL = 3;
    
    public const uint ISIOI_ICONFILE = 1;
    
    public const uint ISIOI_ICONINDEX = 2;
    
    public static Guid ItemCount_Property_GUID => new(0xabbf5c45, 0x5ccc, 0x47b7, 0xbb, 0x4e, 0x87, 0xcb, 0x87, 0xbb, 0xd1, 0x62);
    
    public static Guid ItemIndex_Property_GUID => new(0x92a053da, 0x2969, 0x4021, 0xbf, 0x27, 0x51, 0x4c, 0xfc, 0x2e, 0x4a, 0x69);
    
    public const uint ITSAT_DEFAULT_PRIORITY = 268435456;
    
    public const uint ITSAT_MAX_PRIORITY = 2147483647;
    
    public const uint ITSAT_MIN_PRIORITY = 0;
    
    public const uint ITSS_THREAD_TIMEOUT_NO_CHANGE = 4294967294;
    
    public const uint ITSSFLAG_COMPLETE_ON_DESTROY = 0;
    
    public const uint ITSSFLAG_FLAGS_MASK = 3;
    
    public const uint ITSSFLAG_KILL_ON_DESTROY = 1;
    
    public static Guid KnownFolderManager => new(0x4df0c730, 0xdf9d, 0x4ae3, 0x91, 0x53, 0xaa, 0x6b, 0x82, 0xe9, 0x79, 0x5a);
    
    public static Guid LocalThumbnailCache => new(0x50ef4544, 0xac9f, 0x4a8e, 0xb2, 0x1b, 0x8a, 0x26, 0x18, 0x0d, 0xb1, 0x3f);
    
    public static Guid MailRecipient => new(0x9e56be60, 0xc50f, 0x11cf, 0x9a, 0x2c, 0x00, 0xa0, 0xc9, 0x0a, 0x90, 0xce);
    
    public const uint MAX_COLUMN_DESC_LEN = 128;
    
    public const uint MAX_COLUMN_NAME_LEN = 80;
    
    public const uint MAX_SYNCMGR_ID = 64;
    
    public const uint MAX_SYNCMGR_NAME = 128;
    
    public const uint MAX_SYNCMGR_PROGRESSTEXT = 260;
    
    public const uint MAX_SYNCMGRHANDLERNAME = 32;
    
    public const uint MAX_SYNCMGRITEMNAME = 128;
    
    public const uint MAXFILELEN = 13;
    
    public static Guid MergedCategorizer => new(0x8e827c11, 0x33e7, 0x4bc1, 0xb2, 0x42, 0x8c, 0xd9, 0xa1, 0xc2, 0xb3, 0x04);
    
    public static Guid MSDAINITIALIZE => new(0x2206cdb0, 0x19c1, 0x11d1, 0x89, 0xe0, 0x00, 0xc0, 0x4f, 0xd7, 0xa8, 0x29);
    
    public const uint MUTZ_ACCEPT_WILDCARD_SCHEME = 128;
    
    public const uint MUTZ_DONT_UNESCAPE = 2048;
    
    public const uint MUTZ_DONT_USE_CACHE = 4096;
    
    public const uint MUTZ_ENFORCERESTRICTED = 256;
    
    public const uint MUTZ_FORCE_INTRANET_FLAGS = 8192;
    
    public const uint MUTZ_IGNORE_ZONE_MAPPINGS = 16384;
    
    public const uint MUTZ_ISFILE = 2;
    
    public const uint MUTZ_NOSAVEDFILECHECK = 1;
    
    public const uint MUTZ_REQUIRESAVEDFILECHECK = 1024;
    
    public const uint MUTZ_RESERVED = 512;
    
    public const string NAME_3GPP2File = @"3GPP2File";
    
    public const string NAME_3GPPFile = @"3GPPFile";
    
    public const string NAME_AACFile = @"AACFile";
    
    public const string NAME_AbstractActivity = @"AbstractActivity";
    
    public const string NAME_AbstractActivityOccurrence = @"AbstractActivityOccurrence";
    
    public const string NAME_AbstractAudioAlbum = @"AbstractAudioAlbum";
    
    public const string NAME_AbstractAudioPlaylist = @"AbstractAudioPlaylist";
    
    public const string NAME_AbstractAudioVideoAlbum = @"AbstractAudioVideoAlbum";
    
    public const string NAME_AbstractChapteredProduction = @"AbstractChapteredProduction";
    
    public const string NAME_AbstractContact = @"AbstractContact";
    
    public const string NAME_AbstractContactGroup = @"AbstractContactGroup";
    
    public const string NAME_AbstractDocument = @"AbstractDocument";
    
    public const string NAME_AbstractImageAlbum = @"AbstractImageAlbum";
    
    public const string NAME_AbstractMediacast = @"AbstractMediacast";
    
    public const string NAME_AbstractMessage = @"AbstractMessage";
    
    public const string NAME_AbstractMessageFolder = @"AbstractMessageFolder";
    
    public const string NAME_AbstractMultimediaAlbum = @"AbstractMultimediaAlbum";
    
    public const string NAME_AbstractNote = @"AbstractNote";
    
    public const string NAME_AbstractTask = @"AbstractTask";
    
    public const string NAME_AbstractVideoAlbum = @"AbstractVideoAlbum";
    
    public const string NAME_AbstractVideoPlaylist = @"AbstractVideoPlaylist";
    
    public const string NAME_AIFFFile = @"AIFFFile";
    
    public const string NAME_AMRFile = @"AMRFile";
    
    public const string NAME_AnchorResults = @"AnchorResults";
    
    public const string NAME_AnchorResults_Anchor = @"Anchor";
    
    public const string NAME_AnchorResults_AnchorState = @"AnchorState";
    
    public const string NAME_AnchorResults_ResultObjectID = @"ResultObjectID";
    
    public const string NAME_AnchorSyncKnowledge = @"AnchorSyncKnowledge";
    
    public const string NAME_AnchorSyncSvc = @"AnchorSync";
    
    public const string NAME_AnchorSyncSvc_BeginSync = @"BeginSync";
    
    public const string NAME_AnchorSyncSvc_CurrentAnchor = @"AnchorCurrentAnchor";
    
    public const string NAME_AnchorSyncSvc_EndSync = @"EndSync";
    
    public const string NAME_AnchorSyncSvc_FilterType = @"FilterType";
    
    public const string NAME_AnchorSyncSvc_GetChangesSinceAnchor = @"GetChangesSinceAnchor";
    
    public const string NAME_AnchorSyncSvc_KnowledgeObjectID = @"AnchorKnowledgeObjectID";
    
    public const string NAME_AnchorSyncSvc_LastSyncProxyID = @"AnchorLastSyncProxyID";
    
    public const string NAME_AnchorSyncSvc_LocalOnlyDelete = @"LocalOnlyDelete";
    
    public const string NAME_AnchorSyncSvc_ProviderVersion = @"AnchorProviderVersion";
    
    public const string NAME_AnchorSyncSvc_ReplicaID = @"AnchorReplicaID";
    
    public const string NAME_AnchorSyncSvc_SyncFormat = @"SyncFormat";
    
    public const string NAME_AnchorSyncSvc_VersionProps = @"AnchorVersionProps";
    
    public const string NAME_ASFFile = @"ASFFile";
    
    public const string NAME_Association = @"Association";
    
    public const string NAME_ASXPlaylist = @"ASXPlaylist";
    
    public const string NAME_ATSCTSFile = @"ATSCTSFile";
    
    public const string NAME_AudibleFile = @"AudibleFile";
    
    public const string NAME_AudioObj_AudioBitDepth = @"AudioBitDepth";
    
    public const string NAME_AudioObj_AudioBitRate = @"AudioBitRate";
    
    public const string NAME_AudioObj_AudioBlockAlignment = @"AudioBlockAlignment";
    
    public const string NAME_AudioObj_AudioFormatCode = @"AudioFormatCode";
    
    public const string NAME_AudioObj_Channels = @"Channels";
    
    public const string NAME_AudioObj_Lyrics = @"Lyrics";
    
    public const string NAME_AVCHDFile = @"AVCHDFile";
    
    public const string NAME_AVIFile = @"AVIFile";
    
    public const string NAME_BMPImage = @"BMPImage";
    
    public const string NAME_CalendarObj_Accepted = @"Accepted";
    
    public const string NAME_CalendarObj_BeginDateTime = @"BeginDateTime";
    
    public const string NAME_CalendarObj_BusyStatus = @"BusyStatus";
    
    public const string NAME_CalendarObj_Declined = @"Declined";
    
    public const string NAME_CalendarObj_EndDateTime = @"EndDateTime";
    
    public const string NAME_CalendarObj_Location = @"Location";
    
    public const string NAME_CalendarObj_PatternDuration = @"PatternDuration";
    
    public const string NAME_CalendarObj_PatternStartTime = @"PatternStartTime";
    
    public const string NAME_CalendarObj_ReminderOffset = @"ReminderOffset";
    
    public const string NAME_CalendarObj_Tentative = @"Tentative";
    
    public const string NAME_CalendarObj_TimeZone = @"TimeZone";
    
    public const string NAME_CalendarSvc = @"Calendar";
    
    public const string NAME_CalendarSvc_SyncWindowEnd = @"SyncWindowEnd";
    
    public const string NAME_CalendarSvc_SyncWindowStart = @"SyncWindowStart";
    
    public const string NAME_CIFFImage = @"CIFFImage";
    
    public const string NAME_ContactObj_AnniversaryDate = @"AnniversaryDate";
    
    public const string NAME_ContactObj_Assistant = @"Assistant";
    
    public const string NAME_ContactObj_Birthdate = @"Birthdate";
    
    public const string NAME_ContactObj_BusinessAddressCity = @"BusinessAddressCity";
    
    public const string NAME_ContactObj_BusinessAddressCountry = @"BusinessAddressCountry";
    
    public const string NAME_ContactObj_BusinessAddressFull = @"BusinessAddressFull";
    
    public const string NAME_ContactObj_BusinessAddressLine2 = @"BusinessAddressLine2";
    
    public const string NAME_ContactObj_BusinessAddressPostalCode = @"BusinessAddressPostalCode";
    
    public const string NAME_ContactObj_BusinessAddressRegion = @"BusinessAddressRegion";
    
    public const string NAME_ContactObj_BusinessAddressStreet = @"BusinessAddressStreet";
    
    public const string NAME_ContactObj_BusinessEmail = @"BusinessEmail";
    
    public const string NAME_ContactObj_BusinessEmail2 = @"BusinessEmail2";
    
    public const string NAME_ContactObj_BusinessFax = @"BusinessFax";
    
    public const string NAME_ContactObj_BusinessPhone = @"BusinessPhone";
    
    public const string NAME_ContactObj_BusinessPhone2 = @"BusinessPhone2";
    
    public const string NAME_ContactObj_BusinessWebAddress = @"BusinessWebAddress";
    
    public const string NAME_ContactObj_Children = @"Children";
    
    public const string NAME_ContactObj_Email = @"Email";
    
    public const string NAME_ContactObj_FamilyName = @"FamilyName";
    
    public const string NAME_ContactObj_Fax = @"Fax";
    
    public const string NAME_ContactObj_GivenName = @"GivenName";
    
    public const string NAME_ContactObj_IMAddress = @"IMAddress";
    
    public const string NAME_ContactObj_IMAddress2 = @"IMAddress2";
    
    public const string NAME_ContactObj_IMAddress3 = @"IMAddress3";
    
    public const string NAME_ContactObj_MiddleNames = @"MiddleNames";
    
    public const string NAME_ContactObj_MobilePhone = @"MobilePhone";
    
    public const string NAME_ContactObj_MobilePhone2 = @"MobilePhone2";
    
    public const string NAME_ContactObj_Organization = @"Organization";
    
    public const string NAME_ContactObj_OtherAddressCity = @"OtherAddressCity";
    
    public const string NAME_ContactObj_OtherAddressCountry = @"OtherAddressCountry";
    
    public const string NAME_ContactObj_OtherAddressFull = @"OtherAddressFull";
    
    public const string NAME_ContactObj_OtherAddressLine2 = @"OtherAddressLine2";
    
    public const string NAME_ContactObj_OtherAddressPostalCode = @"OtherAddressPostalCode";
    
    public const string NAME_ContactObj_OtherAddressRegion = @"OtherAddressRegion";
    
    public const string NAME_ContactObj_OtherAddressStreet = @"OtherAddressStreet";
    
    public const string NAME_ContactObj_OtherEmail = @"OtherEmail";
    
    public const string NAME_ContactObj_OtherPhone = @"OtherPhone";
    
    public const string NAME_ContactObj_Pager = @"Pager";
    
    public const string NAME_ContactObj_PersonalAddressCity = @"PersonalAddressCity";
    
    public const string NAME_ContactObj_PersonalAddressCountry = @"PersonalAddressCountry";
    
    public const string NAME_ContactObj_PersonalAddressFull = @"PersonalAddressFull";
    
    public const string NAME_ContactObj_PersonalAddressLine2 = @"PersonalAddressLine2";
    
    public const string NAME_ContactObj_PersonalAddressPostalCode = @"PersonalAddressPostalCode";
    
    public const string NAME_ContactObj_PersonalAddressRegion = @"PersonalAddressRegion";
    
    public const string NAME_ContactObj_PersonalAddressStreet = @"PersonalAddressStreet";
    
    public const string NAME_ContactObj_PersonalEmail = @"PersonalEmail";
    
    public const string NAME_ContactObj_PersonalEmail2 = @"PersonalEmail2";
    
    public const string NAME_ContactObj_PersonalFax = @"PersonalFax";
    
    public const string NAME_ContactObj_PersonalPhone = @"PersonalPhone";
    
    public const string NAME_ContactObj_PersonalPhone2 = @"PersonalPhone2";
    
    public const string NAME_ContactObj_PersonalWebAddress = @"PersonalWebAddress";
    
    public const string NAME_ContactObj_Phone = @"Phone";
    
    public const string NAME_ContactObj_PhoneticFamilyName = @"PhoneticFamilyName";
    
    public const string NAME_ContactObj_PhoneticGivenName = @"PhoneticGivenName";
    
    public const string NAME_ContactObj_PhoneticOrganization = @"PhoneticOrganization";
    
    public const string NAME_ContactObj_Ringtone = @"Ringtone";
    
    public const string NAME_ContactObj_Role = @"Role";
    
    public const string NAME_ContactObj_Spouse = @"Spouse";
    
    public const string NAME_ContactObj_Suffix = @"Suffix";
    
    public const string NAME_ContactObj_Title = @"Title";
    
    public const string NAME_ContactObj_WebAddress = @"WebAddress";
    
    public const string NAME_ContactsSvc = @"Contacts";
    
    public const string NAME_ContactSvc_SyncWithPhoneOnly = @"FilterType";
    
    public const string NAME_DeviceExecutable = @"DeviceExecutable";
    
    public const string NAME_DeviceMetadataCAB = @"DeviceMetadataCAB";
    
    public const string NAME_DeviceMetadataObj_ContentID = @"ContentID";
    
    public const string NAME_DeviceMetadataObj_DefaultCAB = @"DefaultCAB";
    
    public const string NAME_DeviceMetadataSvc = @"Metadata";
    
    public const string NAME_DeviceScript = @"DeviceScript";
    
    public const string NAME_DPOFDocument = @"DPOFDocument";
    
    public const string NAME_DVBTSFile = @"DVBTSFile";
    
    public const string NAME_ExcelDocument = @"ExcelDocument";
    
    public const string NAME_EXIFImage = @"EXIFImage";
    
    public const string NAME_FirmwareFile = @"FirmwareFile";
    
    public const string NAME_FLACFile = @"FLACFile";
    
    public const string NAME_FlashPixImage = @"FlashPixImage";
    
    public const string NAME_FullEnumSyncKnowledge = @"FullEnumSyncKnowledge";
    
    public const string NAME_FullEnumSyncSvc = @"FullEnumSync";
    
    public const string NAME_FullEnumSyncSvc_BeginSync = @"BeginSync";
    
    public const string NAME_FullEnumSyncSvc_EndSync = @"EndSync";
    
    public const string NAME_FullEnumSyncSvc_FilterType = @"FilterType";
    
    public const string NAME_FullEnumSyncSvc_KnowledgeObjectID = @"FullEnumKnowledgeObjectID";
    
    public const string NAME_FullEnumSyncSvc_LastSyncProxyID = @"FullEnumLastSyncProxyID";
    
    public const string NAME_FullEnumSyncSvc_LocalOnlyDelete = @"LocalOnlyDelete";
    
    public const string NAME_FullEnumSyncSvc_ProviderVersion = @"FullEnumProviderVersion";
    
    public const string NAME_FullEnumSyncSvc_ReplicaID = @"FullEnumReplicaID";
    
    public const string NAME_FullEnumSyncSvc_SyncFormat = @"SyncFormat";
    
    public const string NAME_FullEnumSyncSvc_VersionProps = @"FullEnumVersionProps";
    
    public const string NAME_GenericObj_AllowedFolderContents = @"AllowedFolderContents";
    
    public const string NAME_GenericObj_AssociationDesc = @"AssociationDesc";
    
    public const string NAME_GenericObj_AssociationType = @"AssociationType";
    
    public const string NAME_GenericObj_Copyright = @"Copyright";
    
    public const string NAME_GenericObj_Corrupt = @"Corrupt";
    
    public const string NAME_GenericObj_DateAccessed = @"DateAccessed";
    
    public const string NAME_GenericObj_DateAdded = @"DateAdded";
    
    public const string NAME_GenericObj_DateAuthored = @"DateAuthored";
    
    public const string NAME_GenericObj_DateCreated = @"DateCreated";
    
    public const string NAME_GenericObj_DateModified = @"DateModified";
    
    public const string NAME_GenericObj_DateRevised = @"DateRevised";
    
    public const string NAME_GenericObj_Description = @"Description";
    
    public const string NAME_GenericObj_DRMStatus = @"DRMStatus";
    
    public const string NAME_GenericObj_Hidden = @"Hidden";
    
    public const string NAME_GenericObj_Keywords = @"Keywords";
    
    public const string NAME_GenericObj_LanguageLocale = @"LanguageLocale";
    
    public const string NAME_GenericObj_Name = @"Name";
    
    public const string NAME_GenericObj_NonConsumable = @"NonConsumable";
    
    public const string NAME_GenericObj_ObjectFileName = @"ObjectFileName";
    
    public const string NAME_GenericObj_ObjectFormat = @"ObjectFormat";
    
    public const string NAME_GenericObj_ObjectID = @"ObjectID";
    
    public const string NAME_GenericObj_ObjectSize = @"ObjectSize";
    
    public const string NAME_GenericObj_ParentID = @"ParentID";
    
    public const string NAME_GenericObj_PersistentUID = @"PersistentUID";
    
    public const string NAME_GenericObj_PropertyBag = @"PropertyBag";
    
    public const string NAME_GenericObj_ProtectionStatus = @"ProtectionStatus";
    
    public const string NAME_GenericObj_ReferenceParentID = @"ReferenceParentID";
    
    public const string NAME_GenericObj_StorageID = @"StorageID";
    
    public const string NAME_GenericObj_SubDescription = @"SubDescription";
    
    public const string NAME_GenericObj_SyncID = @"SyncID";
    
    public const string NAME_GenericObj_SystemObject = @"SystemObject";
    
    public const string NAME_GenericObj_TimeToLive = @"TimeToLive";
    
    public const string NAME_GIFImage = @"GIFImage";
    
    public const string NAME_HDPhotoImage = @"HDPhotoImage";
    
    public const string NAME_HintsSvc = @"Hints";
    
    public const string NAME_HTMLDocument = @"HTMLDocument";
    
    public const string NAME_ICalendarActivity = @"ICalendar";
    
    public const string NAME_ImageObj_Aperature = @"Aperature";
    
    public const string NAME_ImageObj_Exposure = @"Exposure";
    
    public const string NAME_ImageObj_ImageBitDepth = @"ImageBitDepth";
    
    public const string NAME_ImageObj_IsColorCorrected = @"IsColorCorrected";
    
    public const string NAME_ImageObj_IsCropped = @"IsCropped";
    
    public const string NAME_ImageObj_ISOSpeed = @"ISOSpeed";
    
    public const string NAME_JFIFImage = @"JFIFImage";
    
    public const string NAME_JP2Image = @"JP2Image";
    
    public const string NAME_JPEGXRImage = @"JPEGXRImage";
    
    public const string NAME_JPXImage = @"JPXImage";
    
    public const string NAME_M3UPlaylist = @"M3UPlaylist";
    
    public const string NAME_MediaObj_AlbumArtist = @"AlbumArtist";
    
    public const string NAME_MediaObj_AlbumName = @"AlbumName";
    
    public const string NAME_MediaObj_Artist = @"Artist";
    
    public const string NAME_MediaObj_AudioEncodingProfile = @"AudioEncodingProfile";
    
    public const string NAME_MediaObj_BitRateType = @"BitRateType";
    
    public const string NAME_MediaObj_BookmarkByte = @"BookmarkByte";
    
    public const string NAME_MediaObj_BookmarkObject = @"BookmarkObject";
    
    public const string NAME_MediaObj_BookmarkTime = @"BookmarkTime";
    
    public const string NAME_MediaObj_BufferSize = @"BufferSize";
    
    public const string NAME_MediaObj_Composer = @"Composer";
    
    public const string NAME_MediaObj_Credits = @"Credits";
    
    public const string NAME_MediaObj_DateOriginalRelease = @"DateOriginalRelease";
    
    public const string NAME_MediaObj_Duration = @"Duration";
    
    public const string NAME_MediaObj_Editor = @"Editor";
    
    public const string NAME_MediaObj_EffectiveRating = @"EffectiveRating";
    
    public const string NAME_MediaObj_EncodingProfile = @"EncodingProfile";
    
    public const string NAME_MediaObj_EncodingQuality = @"EncodingQuality";
    
    public const string NAME_MediaObj_Genre = @"Genre";
    
    public const string NAME_MediaObj_GeographicOrigin = @"GeographicOrigin";
    
    public const string NAME_MediaObj_Height = @"Height";
    
    public const string NAME_MediaObj_MediaType = @"MediaType";
    
    public const string NAME_MediaObj_MediaUID = @"MediaUID";
    
    public const string NAME_MediaObj_Mood = @"Mood";
    
    public const string NAME_MediaObj_Owner = @"Owner";
    
    public const string NAME_MediaObj_ParentalRating = @"ParentalRating";
    
    public const string NAME_MediaObj_Producer = @"Producer";
    
    public const string NAME_MediaObj_SampleRate = @"SampleRate";
    
    public const string NAME_MediaObj_SkipCount = @"SkipCount";
    
    public const string NAME_MediaObj_SubscriptionContentID = @"SubscriptionContentID";
    
    public const string NAME_MediaObj_Subtitle = @"Subtitle";
    
    public const string NAME_MediaObj_TotalBitRate = @"TotalBitRate";
    
    public const string NAME_MediaObj_Track = @"Track";
    
    public const string NAME_MediaObj_URLLink = @"URLLink";
    
    public const string NAME_MediaObj_URLSource = @"URLSource";
    
    public const string NAME_MediaObj_UseCount = @"UseCount";
    
    public const string NAME_MediaObj_UserRating = @"UserRating";
    
    public const string NAME_MediaObj_WebMaster = @"WebMaster";
    
    public const string NAME_MediaObj_Width = @"Width";
    
    public const string NAME_MessageObj_BCC = @"BCC";
    
    public const string NAME_MessageObj_Body = @"Body";
    
    public const string NAME_MessageObj_Category = @"Category";
    
    public const string NAME_MessageObj_CC = @"CC";
    
    public const string NAME_MessageObj_PatternDayOfMonth = @"PatternDayOfMonth";
    
    public const string NAME_MessageObj_PatternDayOfWeek = @"PatternDayOfWeek";
    
    public const string NAME_MessageObj_PatternDeleteDates = @"PatternDeleteDates";
    
    public const string NAME_MessageObj_PatternInstance = @"PatternInstance";
    
    public const string NAME_MessageObj_PatternMonthOfYear = @"PatternMonthOfYear";
    
    public const string NAME_MessageObj_PatternOriginalDateTime = @"PatternOriginalDateTime";
    
    public const string NAME_MessageObj_PatternPeriod = @"PatternPeriod";
    
    public const string NAME_MessageObj_PatternType = @"PatternType";
    
    public const string NAME_MessageObj_PatternValidEndDate = @"PatternValidEndDate";
    
    public const string NAME_MessageObj_PatternValidStartDate = @"PatternValidStartDate";
    
    public const string NAME_MessageObj_Priority = @"Priority";
    
    public const string NAME_MessageObj_Read = @"Read";
    
    public const string NAME_MessageObj_ReceivedTime = @"ReceivedTime";
    
    public const string NAME_MessageObj_Sender = @"Sender";
    
    public const string NAME_MessageObj_Subject = @"Subject";
    
    public const string NAME_MessageObj_To = @"To";
    
    public const string NAME_MessageSvc = @"Message";
    
    public const string NAME_MHTDocument = @"MHTDocument";
    
    public const string NAME_MP3File = @"MP3File";
    
    public const string NAME_MPEG2File = @"MPEG2File";
    
    public const string NAME_MPEG4File = @"MPEG4File";
    
    public const string NAME_MPEGFile = @"MPEGFile";
    
    public const string NAME_MPLPlaylist = @"MPLPlaylist";
    
    public const string NAME_NotesSvc = @"Notes";
    
    public const string NAME_OGGFile = @"OGGFile";
    
    public const string NAME_PCDImage = @"PCDImage";
    
    public const string NAME_PICTImage = @"PICTImage";
    
    public const string NAME_PNGImage = @"PNGImage";
    
    public const string NAME_PowerPointDocument = @"PowerPointDocument";
    
    public const string NAME_PSLPlaylist = @"PSLPlaylist";
    
    public const string NAME_QCELPFile = @"QCELPFile";
    
    public const string NAME_RingtonesSvc = @"Ringtones";
    
    public const string NAME_RingtonesSvc_DefaultRingtone = @"DefaultRingtone";
    
    public const string NAME_Services_ServiceDisplayName = @"ServiceDisplayName";
    
    public const string NAME_Services_ServiceIcon = @"ServiceIcon";
    
    public const string NAME_Services_ServiceLocale = @"ServiceLocale";
    
    public const string NAME_StatusSvc = @"Status";
    
    public const string NAME_StatusSvc_BatteryLife = @"BatteryLife";
    
    public const string NAME_StatusSvc_ChargingState = @"ChargingState";
    
    public const string NAME_StatusSvc_MissedCalls = @"MissedCalls";
    
    public const string NAME_StatusSvc_NetworkName = @"NetworkName";
    
    public const string NAME_StatusSvc_NetworkType = @"NetworkType";
    
    public const string NAME_StatusSvc_NewPictures = @"NewPictures";
    
    public const string NAME_StatusSvc_Roaming = @"Roaming";
    
    public const string NAME_StatusSvc_SignalStrength = @"SignalStrength";
    
    public const string NAME_StatusSvc_StorageCapacity = @"StorageCapacity";
    
    public const string NAME_StatusSvc_StorageFreeSpace = @"StorageFreeSpace";
    
    public const string NAME_StatusSvc_TextMessages = @"TextMessages";
    
    public const string NAME_StatusSvc_VoiceMail = @"VoiceMail";
    
    public const string NAME_SyncObj_LastAuthorProxyID = @"LastAuthorProxyID";
    
    public const string NAME_SyncSvc_BeginSync = @"BeginSync";
    
    public const string NAME_SyncSvc_EndSync = @"EndSync";
    
    public const string NAME_SyncSvc_FilterType = @"FilterType";
    
    public const string NAME_SyncSvc_LocalOnlyDelete = @"LocalOnlyDelete";
    
    public const string NAME_SyncSvc_SyncFormat = @"SyncFormat";
    
    public const string NAME_SyncSvc_SyncObjectReferences = @"SyncObjectReferences";
    
    public const string NAME_TaskObj_BeginDate = @"BeginDate";
    
    public const string NAME_TaskObj_Complete = @"Complete";
    
    public const string NAME_TaskObj_EndDate = @"EndDate";
    
    public const string NAME_TaskObj_ReminderDateTime = @"ReminderDateTime";
    
    public const string NAME_TasksSvc = @"Tasks";
    
    public const string NAME_TasksSvc_SyncActiveOnly = @"FilterType";
    
    public const string NAME_TextDocument = @"TextDocument";
    
    public const string NAME_TIFFEPImage = @"TIFFEPImage";
    
    public const string NAME_TIFFImage = @"TIFFImage";
    
    public const string NAME_TIFFITImage = @"TIFFITImage";
    
    public const string NAME_Undefined = @"Undefined";
    
    public const string NAME_UndefinedAudio = @"UndefinedAudio";
    
    public const string NAME_UndefinedCollection = @"UndefinedCollection";
    
    public const string NAME_UndefinedDocument = @"UndefinedDocument";
    
    public const string NAME_UndefinedVideo = @"UndefinedVideo";
    
    public const string NAME_UnknownImage = @"UnknownImage";
    
    public const string NAME_VCalendar1Activity = @"VCalendar1";
    
    public const string NAME_VCard2Contact = @"VCard2Contact";
    
    public const string NAME_VCard3Contact = @"VCard3Contact";
    
    public const string NAME_VideoObj_KeyFrameDistance = @"KeyFrameDistance";
    
    public const string NAME_VideoObj_ScanType = @"ScanType";
    
    public const string NAME_VideoObj_Source = @"Source";
    
    public const string NAME_VideoObj_VideoBitRate = @"VideoBitRate";
    
    public const string NAME_VideoObj_VideoFormatCode = @"VideoFormatCode";
    
    public const string NAME_VideoObj_VideoFrameRate = @"VideoFrameRate";
    
    public const string NAME_WAVFile = @"WAVFile";
    
    public const string NAME_WBMPImage = @"WBMPImage";
    
    public const string NAME_WMAFile = @"WMAFile";
    
    public const string NAME_WMVFile = @"WMVFile";
    
    public const string NAME_WordDocument = @"WordDocument";
    
    public const string NAME_WPLPlaylist = @"WPLPlaylist";
    
    public const string NAME_XMLDocument = @"XMLDocument";
    
    public static Guid NamespaceTreeControl => new(0xae054212, 0x3535, 0x4430, 0x83, 0xed, 0xd5, 0x01, 0xaa, 0x66, 0x80, 0xe6);
    
    public static Guid NamespaceWalker => new(0x72eb61e0, 0x8672, 0x4303, 0x91, 0x75, 0xf2, 0xe4, 0xc6, 0x8b, 0x2e, 0x7c);
    
    public const uint NCM_DISPLAYERRORTIP = 1028;
    
    public const uint NCM_GETADDRESS = 1025;
    
    public const uint NCM_GETALLOWTYPE = 1027;
    
    public const uint NCM_SETALLOWTYPE = 1026;
    
    public static Guid NetworkConnections => new(0x7007acc7, 0x3202, 0x11d1, 0xaa, 0xd2, 0x00, 0x80, 0x5f, 0xc1, 0x27, 0x0e);
    
    public static Guid NetworkExplorerFolder => new(0xf02c1a0d, 0xbe21, 0x4350, 0x88, 0xb0, 0x73, 0x67, 0xfc, 0x96, 0xef, 0x3c);
    
    public static Guid NetworkPlaces => new(0x208d2c60, 0x3aea, 0x1069, 0xa2, 0xd7, 0x08, 0x00, 0x2b, 0x30, 0x30, 0x9d);
    
    public const uint NIN_BALLOONHIDE = 1027;
    
    public const uint NIN_BALLOONSHOW = 1026;
    
    public const uint NIN_BALLOONTIMEOUT = 1028;
    
    public const uint NIN_BALLOONUSERCLICK = 1029;
    
    public const uint NIN_POPUPCLOSE = 1031;
    
    public const uint NIN_POPUPOPEN = 1030;
    
    public const uint NIN_SELECT = 1024;
    
    public const uint NINF_KEY = 1;
    
    public const uint NOTIFYICON_VERSION = 3;
    
    public const uint NOTIFYICON_VERSION_4 = 4;
    
    public static Guid NPCredentialProvider => new(0x3dd6bec0, 0x8193, 0x4ffe, 0xae, 0x25, 0xe0, 0x8e, 0x39, 0xea, 0x40, 0x63);
    
    public const int NSTCDHPOS_ONTOP = -1;
    
    public const uint NT_CONSOLE_PROPS_SIG = 2684354562;
    
    public const uint NT_FE_CONSOLE_PROPS_SIG = 2684354564;
    
    public const uint NUM_POINTS = 3;
    
    public const uint OF_CAP_CANCLOSE = 2;
    
    public const uint OF_CAP_CANSWITCHTO = 1;
    
    public const uint OFASI_EDIT = 1;
    
    public const uint OFASI_OPENDESKTOP = 2;
    
    public const uint OFFLINE_STATUS_INCOMPLETE = 4;
    
    public const uint OFFLINE_STATUS_LOCAL = 1;
    
    public const uint OFFLINE_STATUS_REMOTE = 2;
    
    public const uint OI_ASYNC = 4294962926;
    
    public const uint OI_DEFAULT = 0;
    
    public static Guid OnexCredentialProvider => new(0x07aa0886, 0xcc8d, 0x4e19, 0xa4, 0x10, 0x1c, 0x75, 0xaf, 0x68, 0x6e, 0x62);
    
    public static Guid OnexPlapSmartcardCredentialProvider => new(0x33c86cd6, 0x705f, 0x4ba1, 0x9a, 0xdb, 0x67, 0x07, 0x0b, 0x83, 0x77, 0x75);
    
    public static Guid OpenControlPanel => new(0x06622d85, 0x6856, 0x4460, 0x8d, 0xe1, 0xa8, 0x19, 0x21, 0xb4, 0x1c, 0x4b);
    
    public const uint OPENPROPS_INHIBITPIF = 32768;
    
    public const uint OPENPROPS_NONE = 0;
    
    public static Guid PackageDebugSettings => new(0xb1aec16f, 0x2383, 0x4852, 0xb0, 0xe9, 0x8f, 0x0b, 0x1d, 0xc6, 0x6b, 0x4d);
    
    public const uint PANE_NAVIGATION = 5;
    
    public const uint PANE_NONE = uint.MaxValue;
    
    public const uint PANE_OFFLINE = 2;
    
    public const uint PANE_PRINTER = 3;
    
    public const uint PANE_PRIVACY = 7;
    
    public const uint PANE_PROGRESS = 6;
    
    public const uint PANE_SSL = 4;
    
    public const uint PANE_ZONE = 1;
    
    public static Guid PasswordCredentialProvider => new(0x60b78e88, 0xead8, 0x445c, 0x9c, 0xfd, 0x0b, 0x87, 0xf7, 0x4e, 0xa6, 0xcd);
    
    public const uint PATHCCH_MAX_CCH = 32768;
    
    public const uint PDTIMER_PAUSE = 2;
    
    public const uint PDTIMER_RESET = 1;
    
    public const uint PDTIMER_RESUME = 3;
    
    public const uint PERCEIVEDFLAG_GDIPLUS = 16;
    
    public const uint PERCEIVEDFLAG_HARDCODED = 2;
    
    public const uint PERCEIVEDFLAG_NATIVESUPPORT = 4;
    
    public const uint PERCEIVEDFLAG_SOFTCODED = 1;
    
    public const uint PERCEIVEDFLAG_UNDEFINED = 0;
    
    public const uint PERCEIVEDFLAG_WMSDK = 32;
    
    public const uint PERCEIVEDFLAG_ZIPFOLDER = 64;
    
    public const uint PID_COMPUTERNAME = 5;
    
    public const uint PID_CONTROLPANEL_CATEGORY = 2;
    
    public const uint PID_DESCRIPTIONID = 2;
    
    public const uint PID_DISPLACED_DATE = 3;
    
    public const uint PID_DISPLACED_FROM = 2;
    
    public const uint PID_DISPLAY_PROPERTIES = 0;
    
    public const uint PID_FINDDATA = 0;
    
    public const uint PID_HTMLINFOTIPFILE = 5;
    
    public const uint PID_INTROTEXT = 1;
    
    public const uint PID_LINK_TARGET = 2;
    
    public const uint PID_LINK_TARGET_TYPE = 3;
    
    public const uint PID_MISC_ACCESSCOUNT = 3;
    
    public const uint PID_MISC_OWNER = 4;
    
    public const uint PID_MISC_PICS = 6;
    
    public const uint PID_MISC_STATUS = 2;
    
    public const uint PID_NETRESOURCE = 1;
    
    public const uint PID_NETWORKLOCATION = 4;
    
    public const uint PID_QUERY_RANK = 2;
    
    public const uint PID_SHARE_CSC_STATUS = 2;
    
    public const uint PID_SYNC_COPY_IN = 2;
    
    public const uint PID_VOLUME_CAPACITY = 3;
    
    public const uint PID_VOLUME_FILESYSTEM = 4;
    
    public const uint PID_VOLUME_FREE = 2;
    
    public const uint PID_WHICHFOLDER = 3;
    
    public const uint PIDASI_AVG_DATA_RATE = 4;
    
    public const uint PIDASI_CHANNEL_COUNT = 7;
    
    public const uint PIDASI_COMPRESSION = 10;
    
    public const uint PIDASI_FORMAT = 2;
    
    public const uint PIDASI_SAMPLE_RATE = 5;
    
    public const uint PIDASI_SAMPLE_SIZE = 6;
    
    public const uint PIDASI_STREAM_NAME = 9;
    
    public const uint PIDASI_STREAM_NUMBER = 8;
    
    public const uint PIDASI_TIMELENGTH = 3;
    
    public const uint PIDDRSI_DESCRIPTION = 3;
    
    public const uint PIDDRSI_PLAYCOUNT = 4;
    
    public const uint PIDDRSI_PLAYEXPIRES = 6;
    
    public const uint PIDDRSI_PLAYSTARTS = 5;
    
    public const uint PIDDRSI_PROTECTED = 2;
    
    public const uint PIDSI_ALBUM = 4;
    
    public const uint PIDSI_ARTIST = 2;
    
    public const uint PIDSI_COMMENT = 6;
    
    public const uint PIDSI_GENRE = 11;
    
    public const uint PIDSI_LYRICS = 12;
    
    public const uint PIDSI_SONGTITLE = 3;
    
    public const uint PIDSI_TRACK = 7;
    
    public const uint PIDSI_YEAR = 5;
    
    public const uint PIDVSI_COMPRESSION = 10;
    
    public const uint PIDVSI_DATA_RATE = 8;
    
    public const uint PIDVSI_FRAME_COUNT = 5;
    
    public const uint PIDVSI_FRAME_HEIGHT = 4;
    
    public const uint PIDVSI_FRAME_RATE = 6;
    
    public const uint PIDVSI_FRAME_WIDTH = 3;
    
    public const uint PIDVSI_SAMPLE_SIZE = 9;
    
    public const uint PIDVSI_STREAM_NAME = 2;
    
    public const uint PIDVSI_STREAM_NUMBER = 11;
    
    public const uint PIDVSI_TIMELENGTH = 7;
    
    public const uint PIFDEFFILESIZE = 80;
    
    public const uint PIFDEFPATHSIZE = 64;
    
    public const uint PIFMAXFILEPATH = 260;
    
    public const uint PIFNAMESIZE = 30;
    
    public const uint PIFPARAMSSIZE = 64;
    
    public const uint PIFSHDATASIZE = 64;
    
    public const uint PIFSHPROGSIZE = 64;
    
    public const uint PIFSTARTLOCSIZE = 63;
    
    public static Guid PINLogonCredentialProvider => new(0xcb82ea12, 0x9f71, 0x446d, 0x89, 0xe1, 0x8d, 0x09, 0x24, 0xe1, 0x25, 0x6e);
    
    public const uint PKEY_PIDSTR_MAX = 10;
    
    public const uint PLATFORM_BROWSERONLY = 1;
    
    public const uint PLATFORM_IE3 = 1;
    
    public const uint PLATFORM_INTEGRATED = 2;
    
    public const uint PLATFORM_UNKNOWN = 0;
    
    public const uint PMSF_DONT_STRIP_SPACES = 65536;
    
    public const uint PMSF_MULTIPLE = 1;
    
    public const uint PMSF_NORMAL = 0;
    
    public const uint PO_DELETE = 19;
    
    public const uint PO_PORTCHANGE = 32;
    
    public const uint PO_REN_PORT = 52;
    
    public const uint PO_RENAME = 20;
    
    public const string PORTABLE_DEVICE_DRM_SCHEME_PDDRM = @"PDDRM";
    
    public const string PORTABLE_DEVICE_DRM_SCHEME_WMDRM10_PD = @"WMDRM10-PD";
    
    public const string PORTABLE_DEVICE_ICON = @"Icons";
    
    public const string PORTABLE_DEVICE_IS_MASS_STORAGE = @"PortableDeviceIsMassStorage";
    
    public const string PORTABLE_DEVICE_NAMESPACE_EXCLUDE_FROM_SHELL = @"PortableDeviceNameSpaceExcludeFromShell";
    
    public const string PORTABLE_DEVICE_NAMESPACE_THUMBNAIL_CONTENT_TYPES = @"PortableDeviceNameSpaceThumbnailContentTypes";
    
    public const string PORTABLE_DEVICE_NAMESPACE_TIMEOUT = @"PortableDeviceNameSpaceTimeout";
    
    public const string PORTABLE_DEVICE_TYPE = @"PortableDeviceType";
    
    public static Guid PortableDevice => new(0x728a21c5, 0x3d9e, 0x48d7, 0x98, 0x10, 0x86, 0x48, 0x48, 0xf0, 0xf4, 0x04);
    
    public static Guid PortableDeviceDispatchFactory => new(0x43232233, 0x8338, 0x4658, 0xae, 0x01, 0x0b, 0x4a, 0xe8, 0x30, 0xb6, 0xb0);
    
    public static Guid PortableDeviceFTM => new(0xf7c0039a, 0x4762, 0x488a, 0xb4, 0xb3, 0x76, 0x0e, 0xf9, 0xa1, 0xba, 0x9b);
    
    public static Guid PortableDeviceKeyCollection => new(0xde2d022d, 0x2480, 0x43be, 0x97, 0xf0, 0xd1, 0xfa, 0x2c, 0xf9, 0x8f, 0x4f);
    
    public static Guid PortableDeviceManager => new(0x0af10cec, 0x2ecd, 0x4b92, 0x95, 0x81, 0x34, 0xf6, 0xae, 0x06, 0x37, 0xf3);
    
    public static Guid PortableDevicePropVariantCollection => new(0x08a99e2f, 0x6d6d, 0x4b80, 0xaf, 0x5a, 0xba, 0xf2, 0xbc, 0xbe, 0x4c, 0xb9);
    
    public static Guid PortableDeviceService => new(0xef5db4c2, 0x9312, 0x422c, 0x91, 0x52, 0x41, 0x1c, 0xd9, 0xc4, 0xdd, 0x84);
    
    public static Guid PortableDeviceServiceFTM => new(0x1649b154, 0xc794, 0x497a, 0x9b, 0x03, 0xf3, 0xf0, 0x12, 0x13, 0x02, 0xf3);
    
    public static Guid PortableDeviceValues => new(0x0c15d503, 0xd017, 0x47ce, 0x90, 0x16, 0x7b, 0x3f, 0x97, 0x87, 0x21, 0xcc);
    
    public static Guid PortableDeviceValuesCollection => new(0x3882134d, 0x14cf, 0x4220, 0x9c, 0xb4, 0x43, 0x5f, 0x86, 0xd8, 0x3f, 0x60);
    
    public static Guid PortableDeviceWebControl => new(0x186dd02c, 0x2dec, 0x41b5, 0xa7, 0xd4, 0xb5, 0x90, 0x56, 0xfa, 0xde, 0x51);
    
    public const uint PPCF_ADDARGUMENTS = 3;
    
    public const uint PPCF_ADDQUOTES = 1;
    
    public const uint PPCF_FORCEQUALIFY = 64;
    
    public const uint PPCF_LONGESTPOSSIBLE = 128;
    
    public const uint PPCF_NODIRECTORIES = 16;
    
    public static Guid PreviousVersions => new(0x596ab062, 0xb4d2, 0x4215, 0x9f, 0x74, 0xe9, 0x10, 0x9b, 0x0a, 0x81, 0x53);
    
    public const uint PRINT_PROP_FORCE_NAME = 1;
    
    public const uint PRINTACTION_DOCUMENTDEFAULTS = 6;
    
    public const uint PRINTACTION_NETINSTALL = 2;
    
    public const uint PRINTACTION_NETINSTALLLINK = 3;
    
    public const uint PRINTACTION_OPEN = 0;
    
    public const uint PRINTACTION_OPENNETPRN = 5;
    
    public const uint PRINTACTION_PROPERTIES = 1;
    
    public const uint PRINTACTION_SERVERPROPERTIES = 7;
    
    public const uint PRINTACTION_TESTPAGE = 4;
    
    public const uint PROGDLG_AUTOTIME = 2;
    
    public const uint PROGDLG_MARQUEEPROGRESS = 32;
    
    public const uint PROGDLG_MODAL = 1;
    
    public const uint PROGDLG_NOCANCEL = 64;
    
    public const uint PROGDLG_NOMINIMIZE = 8;
    
    public const uint PROGDLG_NOPROGRESSBAR = 16;
    
    public const uint PROGDLG_NORMAL = 0;
    
    public const uint PROGDLG_NOTIME = 4;
    
    public const string PROP_CONTRACT_DELEGATE = @"ContractDelegate";
    
    public static Guid PropertiesUI => new(0xd912f8cf, 0x0396, 0x4915, 0x88, 0x4e, 0xfb, 0x42, 0x5d, 0x32, 0x94, 0x3b);
    
    public static Guid PropertySystem => new(0xb8967f85, 0x58ae, 0x4f46, 0x9f, 0xb2, 0x5d, 0x79, 0x04, 0x79, 0x8f, 0x4b);
    
    public const string PROPSTR_EXTENSIONCOMPLETIONSTATE = @"ExtensionCompletionState";
    
    public static Guid PSGUID_AUDIO => new(0x64440490, 0x4c8b, 0x11d1, 0x8b, 0x70, 0x08, 0x00, 0x36, 0xb1, 0x1a, 0x03);
    
    public static Guid PSGUID_BRIEFCASE => new(0x328d8b21, 0x7729, 0x4bfc, 0x95, 0x4c, 0x90, 0x2b, 0x32, 0x9d, 0x56, 0xb0);
    
    public static Guid PSGUID_CONTROLPANEL => new(0x305ca226, 0xd286, 0x468e, 0xb8, 0x48, 0x2b, 0x2e, 0x8e, 0x69, 0x7b, 0x74);
    
    public static Guid PSGUID_CUSTOMIMAGEPROPERTIES => new(0x7ecd8b0e, 0xc136, 0x4a9b, 0x94, 0x11, 0x4e, 0xbd, 0x66, 0x73, 0xcc, 0xc3);
    
    public static Guid PSGUID_DISPLACED => new(0x9b174b33, 0x40ff, 0x11d2, 0xa2, 0x7e, 0x00, 0xc0, 0x4f, 0xc3, 0x08, 0x71);
    
    public static Guid PSGUID_DOCUMENTSUMMARYINFORMATION => new(0xd5cdd502, 0x2e9c, 0x101b, 0x93, 0x97, 0x08, 0x00, 0x2b, 0x2c, 0xf9, 0xae);
    
    public static Guid PSGUID_DRM => new(0xaeac19e4, 0x89ae, 0x4508, 0xb9, 0xb7, 0xbb, 0x86, 0x7a, 0xbe, 0xe2, 0xed);
    
    public static Guid PSGUID_IMAGEPROPERTIES => new(0x14b81da1, 0x0135, 0x4d31, 0x96, 0xd9, 0x6c, 0xbf, 0xc9, 0x67, 0x1a, 0x99);
    
    public static Guid PSGUID_IMAGESUMMARYINFORMATION => new(0x6444048f, 0x4c8b, 0x11d1, 0x8b, 0x70, 0x08, 0x00, 0x36, 0xb1, 0x1a, 0x03);
    
    public static Guid PSGUID_LIBRARYPROPERTIES => new(0x5d76b67f, 0x9b3d, 0x44bb, 0xb6, 0xae, 0x25, 0xda, 0x4f, 0x63, 0x8a, 0x67);
    
    public static Guid PSGUID_LINK => new(0xb9b4b3fc, 0x2b51, 0x4a42, 0xb5, 0xd8, 0x32, 0x41, 0x46, 0xaf, 0xcf, 0x25);
    
    public static Guid PSGUID_MEDIAFILESUMMARYINFORMATION => new(0x64440492, 0x4c8b, 0x11d1, 0x8b, 0x70, 0x08, 0x00, 0x36, 0xb1, 0x1a, 0x03);
    
    public static Guid PSGUID_MISC => new(0x9b174b34, 0x40ff, 0x11d2, 0xa2, 0x7e, 0x00, 0xc0, 0x4f, 0xc3, 0x08, 0x71);
    
    public static Guid PSGUID_MUSIC => new(0x56a3372e, 0xce9c, 0x11d2, 0x9f, 0x0e, 0x00, 0x60, 0x97, 0xc6, 0x86, 0xf6);
    
    public static Guid PSGUID_QUERY_D => new(0x49691c90, 0x7e17, 0x101a, 0xa9, 0x1c, 0x08, 0x00, 0x2b, 0x2e, 0xcd, 0xa9);
    
    public static Guid PSGUID_SHARE => new(0xd8c3986f, 0x813b, 0x449c, 0x84, 0x5d, 0x87, 0xb9, 0x5d, 0x67, 0x4a, 0xde);
    
    public static Guid PSGUID_SHELLDETAILS => new(0x28636aa6, 0x953d, 0x11d2, 0xb5, 0xd6, 0x00, 0xc0, 0x4f, 0xd9, 0x18, 0xd0);
    
    public static Guid PSGUID_SUMMARYINFORMATION => new(0xf29f85e0, 0x4ff9, 0x1068, 0xab, 0x91, 0x08, 0x00, 0x2b, 0x27, 0xb3, 0xd9);
    
    public static Guid PSGUID_VIDEO => new(0x64440491, 0x4c8b, 0x11d1, 0x8b, 0x70, 0x08, 0x00, 0x36, 0xb1, 0x1a, 0x03);
    
    public static Guid PSGUID_VOLUME => new(0x9b174b35, 0x40ff, 0x11d2, 0xa2, 0x7e, 0x00, 0xc0, 0x4f, 0xc3, 0x08, 0x71);
    
    public static Guid PSGUID_WEBVIEW => new(0xf2275480, 0xf782, 0x4291, 0xbd, 0x94, 0xf1, 0x36, 0x93, 0x51, 0x3a, 0xec);
    
    public static Guid PublishDropTarget => new(0xcc6eeffb, 0x43f6, 0x46c5, 0x96, 0x19, 0x51, 0xd5, 0x71, 0x96, 0x7f, 0x7d);
    
    public static Guid PublishingWizard => new(0x6b33163c, 0x76a5, 0x4b6c, 0xbf, 0x21, 0x45, 0xde, 0x9c, 0xd5, 0x03, 0xa1);
    
    public const uint QCMINFO_PLACE_AFTER = 1;
    
    public const uint QCMINFO_PLACE_BEFORE = 0;
    
    public static Guid QueryCancelAutoPlay => new(0x331f1768, 0x05a9, 0x4ddd, 0xb8, 0x6e, 0xda, 0xe3, 0x4d, 0xdc, 0x99, 0x8a);
    
    public const uint RANGEMAX_MessageObj_PatternDayOfMonth = 31;
    
    public const uint RANGEMAX_MessageObj_PatternMonthOfYear = 12;
    
    public const uint RANGEMAX_StatusSvc_BatteryLife = 100;
    
    public const uint RANGEMAX_StatusSvc_MissedCalls = 255;
    
    public const uint RANGEMAX_StatusSvc_NewPictures = 65535;
    
    public const uint RANGEMAX_StatusSvc_SignalStrength = 4;
    
    public const uint RANGEMAX_StatusSvc_TextMessages = 255;
    
    public const uint RANGEMAX_StatusSvc_VoiceMail = 255;
    
    public const uint RANGEMIN_MessageObj_PatternDayOfMonth = 1;
    
    public const uint RANGEMIN_MessageObj_PatternMonthOfYear = 1;
    
    public const uint RANGEMIN_StatusSvc_BatteryLife = 0;
    
    public const uint RANGEMIN_StatusSvc_SignalStrength = 0;
    
    public const uint RANGESTEP_MessageObj_PatternDayOfMonth = 1;
    
    public const uint RANGESTEP_MessageObj_PatternMonthOfYear = 1;
    
    public const uint RANGESTEP_StatusSvc_BatteryLife = 1;
    
    public const uint RANGESTEP_StatusSvc_SignalStrength = 1;
    
    public static Guid RASProvider => new(0x5537e283, 0xb1e7, 0x4ef8, 0x9c, 0x6e, 0x7a, 0xb0, 0xaf, 0xe5, 0x05, 0x6d);
    
    public const uint RESOURCEDISPLAYTYPE_DIRECTORY = 9;
    
    public const uint RESOURCEDISPLAYTYPE_NDSCONTAINER = 11;
    
    public const uint RESOURCEDISPLAYTYPE_NETWORK = 6;
    
    public const uint RESOURCEDISPLAYTYPE_ROOT = 7;
    
    public const uint RESOURCEDISPLAYTYPE_SHAREADMIN = 8;
    
    public const uint SBSP_ABSOLUTE = 0;
    
    public const uint SBSP_ACTIVATE_NOFOCUS = 524288;
    
    public const uint SBSP_ALLOW_AUTONAVIGATE = 65536;
    
    public const uint SBSP_CALLERUNTRUSTED = 8388608;
    
    public const uint SBSP_CREATENOHISTORY = 1048576;
    
    public const uint SBSP_DEFBROWSER = 0;
    
    public const uint SBSP_DEFMODE = 0;
    
    public const uint SBSP_EXPLOREMODE = 32;
    
    public const uint SBSP_FEEDNAVIGATION = 536870912;
    
    public const uint SBSP_HELPMODE = 64;
    
    public const uint SBSP_INITIATEDBYHLINKFRAME = 2147483648;
    
    public const uint SBSP_KEEPSAMETEMPLATE = 131072;
    
    public const uint SBSP_KEEPWORDWHEELTEXT = 262144;
    
    public const uint SBSP_NAVIGATEBACK = 16384;
    
    public const uint SBSP_NAVIGATEFORWARD = 32768;
    
    public const uint SBSP_NEWBROWSER = 2;
    
    public const uint SBSP_NOAUTOSELECT = 67108864;
    
    public const uint SBSP_NOTRANSFERHIST = 128;
    
    public const uint SBSP_OPENMODE = 16;
    
    public const uint SBSP_PARENT = 8192;
    
    public const uint SBSP_PLAYNOSOUND = 2097152;
    
    public const uint SBSP_REDIRECT = 1073741824;
    
    public const uint SBSP_RELATIVE = 4096;
    
    public const uint SBSP_SAMEBROWSER = 1;
    
    public const uint SBSP_TRUSTEDFORACTIVEX = 268435456;
    
    public const uint SBSP_TRUSTFIRSTDOWNLOAD = 16777216;
    
    public const uint SBSP_UNTRUSTEDFORDOWNLOAD = 33554432;
    
    public const uint SBSP_WRITENOHISTORY = 134217728;
    
    public static Guid ScheduledTasks => new(0xd6277990, 0x4c6a, 0x11cf, 0x8d, 0x87, 0x00, 0xaa, 0x00, 0x60, 0xf5, 0xbf);
    
    public const uint SCHEME_CREATE = 128;
    
    public const uint SCHEME_DISPLAY = 1;
    
    public const uint SCHEME_DONOTUSE = 64;
    
    public const uint SCHEME_EDIT = 2;
    
    public const uint SCHEME_GLOBAL = 8;
    
    public const uint SCHEME_LOCAL = 4;
    
    public const uint SCHEME_REFRESH = 16;
    
    public const uint SCHEME_UPDATE = 32;
    
    public const uint SCRM_VERIFYPW = 32768;
    
    public const uint SE_ERR_ACCESSDENIED = 5;
    
    public const uint SE_ERR_ASSOCINCOMPLETE = 27;
    
    public const uint SE_ERR_DDEBUSY = 30;
    
    public const uint SE_ERR_DDEFAIL = 29;
    
    public const uint SE_ERR_DDETIMEOUT = 28;
    
    public const uint SE_ERR_DLLNOTFOUND = 32;
    
    public const uint SE_ERR_FNF = 2;
    
    public const uint SE_ERR_NOASSOC = 31;
    
    public const uint SE_ERR_OOM = 8;
    
    public const uint SE_ERR_PNF = 3;
    
    public const uint SE_ERR_SHARE = 26;
    
    public static Guid SearchFolderItemFactory => new(0x14010e02, 0xbbbd, 0x41f0, 0x88, 0xe3, 0xed, 0xa3, 0x71, 0x21, 0x65, 0x84);
    
    public const uint SEE_MASK_ASYNCOK = 1048576;
    
    public const uint SEE_MASK_CLASSKEY = 3;
    
    public const uint SEE_MASK_CLASSNAME = 1;
    
    public const uint SEE_MASK_CONNECTNETDRV = 128;
    
    public const uint SEE_MASK_DEFAULT = 0;
    
    public const uint SEE_MASK_DOENVSUBST = 512;
    
    public const uint SEE_MASK_FLAG_DDEWAIT = 256;
    
    public const uint SEE_MASK_FLAG_HINST_IS_SITE = 134217728;
    
    public const uint SEE_MASK_FLAG_LOG_USAGE = 67108864;
    
    public const uint SEE_MASK_FLAG_NO_UI = 1024;
    
    public const uint SEE_MASK_HMONITOR = 2097152;
    
    public const uint SEE_MASK_HOTKEY = 32;
    
    public const uint SEE_MASK_ICON = 16;
    
    public const uint SEE_MASK_IDLIST = 4;
    
    public const uint SEE_MASK_INVOKEIDLIST = 12;
    
    public const uint SEE_MASK_NO_CONSOLE = 32768;
    
    public const uint SEE_MASK_NOASYNC = 256;
    
    public const uint SEE_MASK_NOCLOSEPROCESS = 64;
    
    public const uint SEE_MASK_NOQUERYCLASSSTORE = 16777216;
    
    public const uint SEE_MASK_NOZONECHECKS = 8388608;
    
    public const uint SEE_MASK_UNICODE = 16384;
    
    public const uint SEE_MASK_WAITFORINPUTIDLE = 33554432;
    
    public static Guid SelectedItemCount_Property_GUID => new(0x8fe316d2, 0x0e52, 0x460a, 0x9c, 0x1e, 0x48, 0xf2, 0x73, 0xd4, 0x70, 0xa3);
    
    public const uint SETPROPS_NONE = 0;
    
    public const int SFBID_PIDLCHANGED = 0;
    
    public const uint SFVM_ADDOBJECT = 3;
    
    public const uint SFVM_GETSELECTEDOBJECTS = 9;
    
    public const uint SFVM_REARRANGE = 1;
    
    public const uint SFVM_REMOVEOBJECT = 6;
    
    public const uint SFVM_SETCLIPBOARD = 16;
    
    public const uint SFVM_SETITEMPOS = 14;
    
    public const uint SFVM_SETPOINTS = 23;
    
    public const uint SFVM_UPDATEOBJECT = 7;
    
    public const uint SFVSOC_INVALIDATE_ALL = 1;
    
    public const uint SFVSOC_NOSCROLL = 2;
    
    public const uint SHA_HASH_LEN = 20;
    
    public const uint shaFixup = 1;
    
    public const uint SHARE_CURRENT_USES_PARMNUM = 7;
    
    public const uint SHARE_FILE_SD_PARMNUM = 501;
    
    public const uint SHARE_MAX_USES_PARMNUM = 6;
    
    public const uint SHARE_NETNAME_PARMNUM = 1;
    
    public const uint SHARE_PASSWD_PARMNUM = 9;
    
    public const uint SHARE_PATH_PARMNUM = 8;
    
    public const uint SHARE_PERMISSIONS_PARMNUM = 5;
    
    public const uint SHARE_QOS_POLICY_PARMNUM = 504;
    
    public const uint SHARE_REMARK_PARMNUM = 4;
    
    public const uint SHARE_SERVER_PARMNUM = 503;
    
    public const uint SHARE_TYPE_PARMNUM = 3;
    
    public static Guid SharedBitmap => new(0x4db26476, 0x6787, 0x4046, 0xb8, 0x36, 0xe8, 0x41, 0x2a, 0x9e, 0x8a, 0x27);
    
    public const uint SHAREDSECRET = 64;
    
    public const string SHAREVISTRING = @"commdlg_ShareViolation";
    
    public const string SHAREVISTRINGA = @"commdlg_ShareViolation";
    
    public const string SHAREVISTRINGW = @"commdlg_ShareViolation";
    
    public static Guid SharingConfigurationManager => new(0x49f371e1, 0x8c5c, 0x4d9c, 0x9a, 0x3b, 0x54, 0xa6, 0x82, 0x7f, 0x51, 0x3c);
    
    public const uint SHARINGSTATUS_NOTSHARED = 0;
    
    public const uint SHARINGSTATUS_PRIVATE = 2;
    
    public const uint SHARINGSTATUS_SHARED = 1;
    
    public static Guid SharpenEffectGuid => new(0x63cbf3ee, 0xc526, 0x402c, 0x8f, 0x71, 0x62, 0xc5, 0x40, 0xbf, 0x51, 0x42);
    
    public const uint SHCDF_UPDATEITEM = 1;
    
    public const int SHCIDS_ALLFIELDS = int.MinValue;
    
    public const int SHCIDS_BITMASK = -65536;
    
    public const int SHCIDS_CANONICALONLY = 268435456;
    
    public const int SHCIDS_COLUMNMASK = 65535;
    
    public const int SHCNEE_MSI_CHANGE = 4;
    
    public const int SHCNEE_MSI_UNINSTALL = 5;
    
    public const int SHCNEE_ORDERCHANGED = 2;
    
    public static Guid Shell => new(0x13709620, 0xc279, 0x11ce, 0xa4, 0x9e, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00);
    
    public static Guid ShellBrowserWindow => new(0xc08afd90, 0xf2a1, 0x11d1, 0x84, 0x55, 0x00, 0xa0, 0xc9, 0x1f, 0x38, 0x80);
    
    public static Guid ShellDesktop => new(0x00021400, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid ShellDispatchInproc => new(0x0a89a860, 0xd7b1, 0x11ce, 0x83, 0x50, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00);
    
    public const string SHELLEX_WIAUIEXTENSION_NAME = @"WiaDialogExtensionHandlers";
    
    public static Guid ShellFolderItem => new(0x2fe352ea, 0xfd1f, 0x11d2, 0xb1, 0xf4, 0x00, 0xc0, 0x4f, 0x8e, 0xeb, 0x3e);
    
    public static Guid ShellFolderView => new(0x62112aa1, 0xebe4, 0x11cf, 0xa5, 0xfb, 0x00, 0x20, 0xaf, 0xe7, 0x29, 0x2d);
    
    public static Guid ShellFolderViewOC => new(0x9ba05971, 0xf6a8, 0x11cf, 0xa4, 0x42, 0x00, 0xa0, 0xc9, 0x0a, 0x8f, 0x39);
    
    public static Guid ShellFSFolder => new(0xf3364ba0, 0x65b9, 0x11ce, 0xa9, 0xba, 0x00, 0xaa, 0x00, 0x4a, 0xe8, 0x37);
    
    public static Guid ShellImageDataFactory => new(0x66e4e4fb, 0xf385, 0x4dd0, 0x8d, 0x74, 0xa2, 0xef, 0xd1, 0xbc, 0x61, 0x78);
    
    public static Guid ShellItem => new(0x9ac9fbe1, 0xe0a2, 0x4ad6, 0xb4, 0xee, 0xe2, 0x12, 0x01, 0x3e, 0xa9, 0x17);
    
    public static Guid ShellLibrary => new(0xd9b3211d, 0xe57f, 0x4426, 0xaa, 0xef, 0x30, 0xa8, 0x06, 0xad, 0xd3, 0x97);
    
    public static Guid ShellLink => new(0x00021401, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
    
    public static Guid ShellLinkObject => new(0x11219420, 0x1768, 0x11d1, 0x95, 0xbe, 0x00, 0x60, 0x97, 0x97, 0xea, 0x4f);
    
    public static Guid ShellNameSpace => new(0x55136805, 0xb2de, 0x11d1, 0xb9, 0xf2, 0x00, 0xa0, 0xc9, 0x8b, 0xc5, 0x47);
    
    public const uint SHELLSTATEVERSION_IE4 = 9;
    
    public const uint SHELLSTATEVERSION_WIN2K = 10;
    
    public static Guid ShellUIHelper => new(0x64ab4bb7, 0x111e, 0x11d1, 0x8f, 0x79, 0x00, 0xc0, 0x4f, 0xc2, 0xfb, 0xe1);
    
    public static Guid ShellWindows => new(0x9ba05972, 0xf6a8, 0x11cf, 0xa4, 0x42, 0x00, 0xa0, 0xc9, 0x0a, 0x8f, 0x39);
    
    public const uint SHERB_NOCONFIRMATION = 1;
    
    public const uint SHERB_NOPROGRESSUI = 2;
    
    public const uint SHERB_NOSOUND = 4;
    
    public const uint SHFT_INVALID = 15;
    
    public const ulong SHGNLI_NOLNK = 8;
    
    public const ulong SHGNLI_NOLOCNAME = 16;
    
    public const ulong SHGNLI_NOUNIQUE = 4;
    
    public const ulong SHGNLI_PIDL = 1;
    
    public const ulong SHGNLI_PREFIXNAME = 2;
    
    public const ulong SHGNLI_USEURLEXT = 32;
    
    public const uint SHGVSPB_ALLFOLDERS = 8;
    
    public const uint SHGVSPB_ALLUSERS = 2;
    
    public const uint SHGVSPB_INHERIT = 16;
    
    public const uint SHGVSPB_NOAUTODEFAULTS = 2147483648;
    
    public const uint SHGVSPB_PERFOLDER = 4;
    
    public const uint SHGVSPB_PERUSER = 1;
    
    public const uint SHGVSPB_ROAM = 32;
    
    public const uint SHHLNF_NOAUTOSELECT = 67108864;
    
    public const uint SHHLNF_WRITENOHISTORY = 134217728;
    
    public const uint SHI_USES_UNLIMITED = uint.MaxValue;
    
    public const uint SHI1_NUM_ELEMENTS = 4;
    
    public const uint SHI1005_FLAGS_ACCESS_BASED_DIRECTORY_ENUM = 2048;
    
    public const uint SHI1005_FLAGS_ALLOW_NAMESPACE_CACHING = 1024;
    
    public const uint SHI1005_FLAGS_CLUSTER_MANAGED = 524288;
    
    public const uint SHI1005_FLAGS_COMPRESS_DATA = 1048576;
    
    public const uint SHI1005_FLAGS_DFS = 1;
    
    public const uint SHI1005_FLAGS_DFS_ROOT = 2;
    
    public const uint SHI1005_FLAGS_DISABLE_CLIENT_BUFFERING = 131072;
    
    public const uint SHI1005_FLAGS_DISABLE_DIRECTORY_HANDLE_LEASING = 4194304;
    
    public const uint SHI1005_FLAGS_ENABLE_CA = 16384;
    
    public const uint SHI1005_FLAGS_ENABLE_HASH = 8192;
    
    public const uint SHI1005_FLAGS_ENCRYPT_DATA = 32768;
    
    public const uint SHI1005_FLAGS_FORCE_LEVELII_OPLOCK = 4096;
    
    public const uint SHI1005_FLAGS_FORCE_SHARED_DELETE = 512;
    
    public const uint SHI1005_FLAGS_IDENTITY_REMOTING = 262144;
    
    public const uint SHI1005_FLAGS_ISOLATED_TRANSPORT = 2097152;
    
    public const uint SHI1005_FLAGS_RESERVED = 65536;
    
    public const uint SHI1005_FLAGS_RESTRICT_EXCLUSIVE_OPENS = 256;
    
    public const uint SHI2_NUM_ELEMENTS = 10;
    
    public const uint SHIFT_PRESSED = 16;
    
    public const uint SHIL_EXTRALARGE = 2;
    
    public const uint SHIL_JUMBO = 4;
    
    public const uint SHIL_LARGE = 0;
    
    public const uint SHIL_LAST = 4;
    
    public const uint SHIL_SMALL = 1;
    
    public const uint SHIL_SYSSMALL = 3;
    
    public const uint SHIMGDEC_DEFAULT = 0;
    
    public const uint SHIMGDEC_LOADFULL = 2;
    
    public const uint SHIMGDEC_THUMBNAIL = 1;
    
    public const string SHIMGKEY_QUALITY = @"Compression";
    
    public const string SHIMGKEY_RAWFORMAT = @"RawDataFormat";
    
    public const uint SHIMSTCAPFLAG_LOCKABLE = 1;
    
    public const uint SHIMSTCAPFLAG_PURGEABLE = 2;
    
    public const uint SHORTPATH_CACHE_ENTRY = 512;
    
    public const uint SHOW_FULLSCREEN = 3;
    
    public const uint SHOW_ICONWINDOW = 2;
    
    public const uint SHOW_OPENNOACTIVATE = 4;
    
    public const uint SHOW_OPENWINDOW = 1;
    
    public const uint SHOWIMEPAD_CATEGORY = 1;
    
    public const uint SHOWIMEPAD_DEFAULT = 0;
    
    public const uint SHOWIMEPAD_GUID = 2;
    
    public static Guid ShowInputPaneAnimationCoordinator => new(0x1f046abf, 0x3202, 0x4dc1, 0x8c, 0xb5, 0x3c, 0x67, 0x61, 0x7c, 0xe1, 0xfa);
    
    public const uint SHPPFW_ASKDIRCREATE = 2;
    
    public const uint SHPPFW_DIRCREATE = 1;
    
    public const uint SHPPFW_IGNOREFILENAME = 4;
    
    public const uint SHPPFW_MEDIACHECKONLY = 16;
    
    public const uint SHPPFW_NONE = 0;
    
    public const uint SHPPFW_NOWRITECHECK = 8;
    
    public const uint SHPWHF_ANYLOCATION = 256;
    
    public const uint SHPWHF_NOFILESELECTOR = 4;
    
    public const uint SHPWHF_NONETPLACECREATE = 2;
    
    public const uint SHPWHF_NORECOMPRESS = 1;
    
    public const uint SHPWHF_USEMRU = 8;
    
    public const uint SHPWHF_VALIDATEVIAWEBFOLDERS = 65536;
    
    public const uint SHPWLEN = 8;
    
    public const uint SHREGSET_FORCE_HKCU = 2;
    
    public const uint SHREGSET_FORCE_HKLM = 8;
    
    public const uint SHREGSET_HKCU = 1;
    
    public const uint SHREGSET_HKLM = 4;
    
    public const uint SHUFFLE_FILE_FLAG_SKIP_INITIALIZING_NEW_CLUSTERS = 1;
    
    public static Guid SID_CommandsPropertyBag => new(0x6e043250, 0x4416, 0x485c, 0xb1, 0x43, 0xe6, 0x2a, 0x76, 0x0d, 0x9f, 0xe5);
    
    public static Guid SID_CtxQueryAssociations => new(0xfaadfc40, 0xb777, 0x4b69, 0xaa, 0x81, 0x77, 0x03, 0x5e, 0xf0, 0xe6, 0xe8);
    
    public static Guid SID_DefView => new(0x6d12fe80, 0x7911, 0x11cf, 0x95, 0x34, 0x00, 0x00, 0xc0, 0x5b, 0xae, 0x0b);
    
    public static Guid SID_LaunchSourceAppUserModelId => new(0x2ce78010, 0x74db, 0x48bc, 0x9c, 0x6a, 0x10, 0xf3, 0x72, 0x49, 0x57, 0x23);
    
    public static Guid SID_LaunchSourceViewSizePreference => new(0x80605492, 0x67d9, 0x414f, 0xaf, 0x89, 0xa1, 0xcd, 0xf1, 0x24, 0x2b, 0xc1);
    
    public static Guid SID_LaunchTargetViewSizePreference => new(0x26db2472, 0xb7b7, 0x406b, 0x97, 0x02, 0x73, 0x0a, 0x4e, 0x20, 0xd3, 0xbf);
    
    public static Guid SID_MenuShellFolder => new(0xa6c17eb4, 0x2d65, 0x11d2, 0x83, 0x8f, 0x00, 0xc0, 0x4f, 0xd9, 0x18, 0xd0);
    
    public static Guid SID_SCommandBarState => new(0xb99eaa5c, 0x3850, 0x4400, 0xbc, 0x33, 0x2c, 0xe5, 0x34, 0x04, 0x8b, 0xf8);
    
    public static Guid SID_SCommDlgBrowser => new(0x80f30233, 0xb7df, 0x11d2, 0xa3, 0x3b, 0x00, 0x60, 0x97, 0xdf, 0x5b, 0xd4);
    
    public static Guid SID_SGetViewFromViewDual => new(0x889a935d, 0x971e, 0x4b12, 0xb9, 0x0c, 0x24, 0xdf, 0xc9, 0xe1, 0xe5, 0xe8);
    
    public static Guid SID_ShellExecuteNamedPropertyStore => new(0xeb84ada2, 0x00ff, 0x4992, 0x83, 0x24, 0xed, 0x5c, 0xe0, 0x61, 0xcb, 0x29);
    
    public static Guid SID_SInPlaceBrowser => new(0x1d2ae02b, 0x3655, 0x46cc, 0xb6, 0x3a, 0x28, 0x59, 0x88, 0x15, 0x3b, 0xca);
    
    public static Guid SID_SMenuBandBKContextMenu => new(0x164bbd86, 0x1d0d, 0x4de0, 0x9a, 0x3b, 0xd9, 0x72, 0x96, 0x47, 0xc2, 0xb8);
    
    public static Guid SID_SMenuBandBottom => new(0x743ca664, 0x0deb, 0x11d1, 0x98, 0x25, 0x00, 0xc0, 0x4f, 0xd9, 0x19, 0x72);
    
    public static Guid SID_SMenuBandBottomSelected => new(0x165ebaf4, 0x6d51, 0x11d2, 0x83, 0xad, 0x00, 0xc0, 0x4f, 0xd9, 0x18, 0xd0);
    
    public static Guid SID_SMenuBandChild => new(0xed9cc020, 0x08b9, 0x11d1, 0x98, 0x23, 0x00, 0xc0, 0x4f, 0xd9, 0x19, 0x72);
    
    public static Guid SID_SMenuBandContextMenuModifier => new(0x39545874, 0x7162, 0x465e, 0xb7, 0x83, 0x2a, 0xa1, 0x87, 0x4f, 0xef, 0x81);
    
    public static Guid SID_SMenuBandParent => new(0x8c278eec, 0x3eab, 0x11d1, 0x8c, 0xb0, 0x00, 0xc0, 0x4f, 0xd9, 0x18, 0xd0);
    
    public static Guid SID_SMenuBandTop => new(0x9493a810, 0xec38, 0x11d0, 0xbc, 0x46, 0x00, 0xaa, 0x00, 0x6c, 0xe2, 0xf5);
    
    public static Guid SID_SMenuPopup => new(0xd1e7afeb, 0x6a2e, 0x11d0, 0x8c, 0x78, 0x00, 0xc0, 0x4f, 0xd9, 0x18, 0xb4);
    
    public static Guid SID_SSearchBoxInfo => new(0x142daa61, 0x516b, 0x4713, 0xb4, 0x9c, 0xfb, 0x98, 0x5e, 0xf8, 0x29, 0x98);
    
    public static Guid SID_STopLevelBrowser => new(0x4c96be40, 0x915c, 0x11cf, 0x99, 0xd3, 0x00, 0xaa, 0x00, 0x4a, 0xe8, 0x37);
    
    public static Guid SID_STopWindow => new(0x49e1b500, 0x4636, 0x11d3, 0x97, 0xf7, 0x00, 0xc0, 0x4f, 0x45, 0xd0, 0xb3);
    
    public static Guid SID_URLExecutionContext => new(0xfb5f8ebc, 0xbbb6, 0x4d10, 0xa4, 0x61, 0x77, 0x72, 0x91, 0xa0, 0x90, 0x30);
    
    public static Guid SimpleConflictPresenter => new(0x7a0f6ab7, 0xed84, 0x46b6, 0xb4, 0x7e, 0x02, 0xaa, 0x15, 0x9a, 0x15, 0x2b);
    
    public const uint SIOM_ICONINDEX = 2;
    
    public const uint SIOM_OVERLAYINDEX = 1;
    
    public const uint SIOM_RESERVED_DEFAULT = 3;
    
    public const uint SIOM_RESERVED_LINK = 1;
    
    public const uint SIOM_RESERVED_SHARED = 0;
    
    public const uint SIOM_RESERVED_SLOWFILE = 2;
    
    public static Guid SizeCategorizer => new(0x55d7b852, 0xf6d1, 0x42f2, 0xaa, 0x75, 0x87, 0x28, 0xa1, 0xb2, 0xd2, 0x64);
    
    public const uint SMAE_CONTRACTED = 2;
    
    public const uint SMAE_EXPANDED = 1;
    
    public const uint SMAE_USER = 4;
    
    public const uint SMAE_VALID = 7;
    
    public static Guid SmartcardCredentialProvider => new(0x8fd7e19c, 0x3bf7, 0x489b, 0xa7, 0x2c, 0x84, 0x6a, 0xb3, 0x67, 0x8c, 0x96);
    
    public static Guid SmartcardPinProvider => new(0x94596c7e, 0x3744, 0x41ce, 0x89, 0x3e, 0xbb, 0xf0, 0x91, 0x22, 0xf7, 0x6a);
    
    public static Guid SmartcardReaderSelectionProvider => new(0x1b283861, 0x754f, 0x4022, 0xad, 0x47, 0xa5, 0xea, 0xaa, 0x61, 0x88, 0x94);
    
    public static Guid SmartcardWinRTProvider => new(0x1ee7337f, 0x85ac, 0x45e2, 0xa2, 0x3c, 0x37, 0xc7, 0x53, 0x20, 0x97, 0x69);
    
    public const uint SMC_AUTOEXPANDCHANGE = 66;
    
    public const uint SMC_CHEVRONEXPAND = 25;
    
    public const uint SMC_CHEVRONGETTIP = 47;
    
    public const uint SMC_CREATE = 2;
    
    public const uint SMC_DEFAULTICON = 22;
    
    public const uint SMC_DEMOTE = 17;
    
    public const uint SMC_DISPLAYCHEVRONTIP = 42;
    
    public const uint SMC_EXITMENU = 3;
    
    public const uint SMC_GETAUTOEXPANDSTATE = 65;
    
    public const uint SMC_GETBKCONTEXTMENU = 68;
    
    public const uint SMC_GETCONTEXTMENUMODIFIER = 67;
    
    public const uint SMC_GETINFO = 5;
    
    public const uint SMC_GETOBJECT = 7;
    
    public const uint SMC_GETSFINFO = 6;
    
    public const uint SMC_GETSFOBJECT = 8;
    
    public const uint SMC_INITMENU = 1;
    
    public const uint SMC_NEWITEM = 23;
    
    public const uint SMC_OPEN = 69;
    
    public const uint SMC_PROMOTE = 18;
    
    public const uint SMC_REFRESH = 16;
    
    public const uint SMC_SETSFOBJECT = 45;
    
    public const uint SMC_SFDDRESTRICTED = 48;
    
    public const uint SMC_SFEXEC = 9;
    
    public const uint SMC_SFEXEC_MIDDLE = 49;
    
    public const uint SMC_SFSELECTITEM = 10;
    
    public const uint SMC_SHCHANGENOTIFY = 46;
    
    public const uint SMDM_HMENU = 2;
    
    public const uint SMDM_SHELLFOLDER = 1;
    
    public const uint SMDM_TOOLBAR = 4;
    
    public const uint SMINIT_AUTOEXPAND = 256;
    
    public const uint SMINIT_AUTOTOOLTIP = 512;
    
    public const uint SMINIT_CACHED = 16;
    
    public const uint SMINIT_DEFAULT = 0;
    
    public const uint SMINIT_DROPONCONTAINER = 1024;
    
    public const uint SMINIT_HORIZONTAL = 536870912;
    
    public const uint SMINIT_RESTRICT_DRAGDROP = 2;
    
    public const uint SMINIT_TOPLEVEL = 4;
    
    public const uint SMINIT_VERTICAL = 268435456;
    
    public const uint SMINV_ID = 8;
    
    public const uint SMINV_REFRESH = 1;
    
    public const uint SMSET_BOTTOM = 536870912;
    
    public const uint SMSET_DONTOWN = 1;
    
    public const uint SMSET_TOP = 268435456;
    
    public const uint SPMODE_BROWSER = 8;
    
    public const uint SPMODE_DBMON = 8192;
    
    public const uint SPMODE_DEBUGBREAK = 512;
    
    public const uint SPMODE_DEBUGOUT = 2;
    
    public const uint SPMODE_EVENT = 32;
    
    public const uint SPMODE_EVENTTRACE = 32768;
    
    public const uint SPMODE_FLUSH = 16;
    
    public const uint SPMODE_FORMATTEXT = 128;
    
    public const uint SPMODE_MEMWATCH = 4096;
    
    public const uint SPMODE_MSGTRACE = 1024;
    
    public const uint SPMODE_MSVM = 64;
    
    public const uint SPMODE_MULTISTOP = 16384;
    
    public const uint SPMODE_PERFTAGS = 2048;
    
    public const uint SPMODE_PROFILE = 256;
    
    public const uint SPMODE_SHELL = 1;
    
    public const uint SPMODE_TEST = 4;
    
    public const uint SRRF_NOEXPAND = 268435456;
    
    public const uint SRRF_NOVIRT = 1073741824;
    
    public const uint SRRF_RM_ANY = 0;
    
    public const uint SRRF_RM_NORMAL = 65536;
    
    public const uint SRRF_RM_SAFE = 131072;
    
    public const uint SRRF_RM_SAFENETWORK = 262144;
    
    public const uint SRRF_RT_ANY = 65535;
    
    public const uint SRRF_RT_REG_BINARY = 8;
    
    public const uint SRRF_RT_REG_DWORD = 16;
    
    public const uint SRRF_RT_REG_EXPAND_SZ = 4;
    
    public const uint SRRF_RT_REG_MULTI_SZ = 32;
    
    public const uint SRRF_RT_REG_NONE = 1;
    
    public const uint SRRF_RT_REG_QWORD = 64;
    
    public const uint SRRF_RT_REG_SZ = 2;
    
    public const uint SRRF_ZEROONFAILURE = 536870912;
    
    public const uint SSM_CLEAR = 0;
    
    public const uint SSM_REFRESH = 2;
    
    public const uint SSM_SET = 1;
    
    public const uint SSM_UPDATE = 4;
    
    public static Guid StartMenuPin => new(0xa2a9545d, 0xa0c2, 0x42b4, 0x97, 0x08, 0xa0, 0xb2, 0xba, 0xdd, 0x77, 0xc8);
    
    public const int STIF_DEFAULT = 0;
    
    public const int STIF_SUPPORT_HEX = 1;
    
    public static Guid StorageProviderBanners => new(0x7ccdf9f4, 0xe576, 0x455a, 0x8b, 0xc7, 0xf6, 0xec, 0x68, 0xd6, 0xf0, 0x63);
    
    public const string STR_AVOID_DRIVE_RESTRICTION_POLICY = @"Avoid Drive Restriction Policy";
    
    public const string STR_BIND_DELEGATE_CREATE_OBJECT = @"Delegate Object Creation";
    
    public const string STR_BIND_FOLDER_ENUM_MODE = @"Folder Enum Mode";
    
    public const string STR_BIND_FOLDERS_READ_ONLY = @"Folders As Read Only";
    
    public const string STR_BIND_FORCE_FOLDER_SHORTCUT_RESOLVE = @"Force Folder Shortcut Resolve";
    
    public const string STR_DONT_PARSE_RELATIVE = @"Don't Parse Relative";
    
    public const string STR_DONT_RESOLVE_LINK = @"Don't Resolve Link";
    
    public const string STR_ENUM_ITEMS_FLAGS = @"SHCONTF";
    
    public const string STR_FILE_SYS_BIND_DATA = @"File System Bind Data";
    
    public const string STR_FILE_SYS_BIND_DATA_WIN7_FORMAT = @"Win7FileSystemIdList";
    
    public const string STR_GET_ASYNC_HANDLER = @"GetAsyncHandler";
    
    public const string STR_GPS_BESTEFFORT = @"GPS_BESTEFFORT";
    
    public const string STR_GPS_DELAYCREATION = @"GPS_DELAYCREATION";
    
    public const string STR_GPS_FASTPROPERTIESONLY = @"GPS_FASTPROPERTIESONLY";
    
    public const string STR_GPS_HANDLERPROPERTIESONLY = @"GPS_HANDLERPROPERTIESONLY";
    
    public const string STR_GPS_NO_OPLOCK = @"GPS_NO_OPLOCK";
    
    public const string STR_GPS_OPENSLOWITEM = @"GPS_OPENSLOWITEM";
    
    public const string STR_INTERNAL_NAVIGATE = @"Internal Navigation";
    
    public const string STR_INTERNETFOLDER_PARSE_ONLY_URLMON_BINDABLE = @"Validate URL";
    
    public const string STR_ITEM_CACHE_CONTEXT = @"ItemCacheContext";
    
    public const string STR_MYDOCS_CLSID = @"{450D8FBA-AD25-11D0-98A8-0800361B1103}";
    
    public const string STR_NO_VALIDATE_FILENAME_CHARS = @"NoValidateFilenameChars";
    
    public const string STR_PARSE_ALLOW_INTERNET_SHELL_FOLDERS = @"Allow binding to Internet shell folder handlers and negate STR_PARSE_PREFER_WEB_BROWSING";
    
    public const string STR_PARSE_AND_CREATE_ITEM = @"ParseAndCreateItem";
    
    public const string STR_PARSE_DONT_REQUIRE_VALIDATED_URLS = @"Do not require validated URLs";
    
    public const string STR_PARSE_EXPLICIT_ASSOCIATION_SUCCESSFUL = @"ExplicitAssociationSuccessful";
    
    public const string STR_PARSE_PARTIAL_IDLIST = @"ParseOriginalItem";
    
    public const string STR_PARSE_PREFER_FOLDER_BROWSING = @"Parse Prefer Folder Browsing";
    
    public const string STR_PARSE_PREFER_WEB_BROWSING = @"Do not bind to Internet shell folder handlers";
    
    public const string STR_PARSE_PROPERTYSTORE = @"DelegateNamedProperties";
    
    public const string STR_PARSE_SHELL_PROTOCOL_TO_FILE_OBJECTS = @"Parse Shell Protocol To File Objects";
    
    public const string STR_PARSE_SHOW_NET_DIAGNOSTICS_UI = @"Show network diagnostics UI";
    
    public const string STR_PARSE_SKIP_NET_CACHE = @"Skip Net Resource Cache";
    
    public const string STR_PARSE_TRANSLATE_ALIASES = @"Parse Translate Aliases";
    
    public const string STR_PARSE_WITH_EXPLICIT_ASSOCAPP = @"ExplicitAssociationApp";
    
    public const string STR_PARSE_WITH_EXPLICIT_PROGID = @"ExplicitProgid";
    
    public const string STR_PARSE_WITH_PROPERTIES = @"ParseWithProperties";
    
    public const string STR_PROPERTYBAG_PARAM = @"SHBindCtxPropertyBag";
    
    public const string STR_REFERRER_IDENTIFIER = @"Referrer Identifier";
    
    public const string STR_SKIP_BINDING_CLSID = @"Skip Binding CLSID";
    
    public const string STR_STORAGEITEM_CREATION_FLAGS = @"SHGETSTORAGEITEM";
    
    public const string STR_TAB_REUSE_IDENTIFIER = @"Tab Reuse Identifier";
    
    public const string STR_TRACK_CLSID = @"Track the CLSID";
    
    public const string STR_WPDNSE_FAST_ENUM = @"WPDNSE Fast Enum";
    
    public const string STR_WPDNSE_SIMPLE_ITEM = @"WPDNSE SimpleItem";
    
    public static Guid SuspensionDependencyManager => new(0x6b273fc5, 0x61fd, 0x4918, 0x95, 0xa2, 0xc3, 0xb5, 0xe9, 0xd7, 0xf5, 0x81);
    
    public static Guid SyncMgr => new(0x6295df27, 0x35ee, 0x11d1, 0x87, 0x07, 0x00, 0xc0, 0x4f, 0xd9, 0x33, 0x27);
    
    public static Guid SYNCMGR_OBJECTID_BrowseContent => new(0x57cbb584, 0xe9b4, 0x47ae, 0xa1, 0x20, 0xc4, 0xdf, 0x33, 0x35, 0xde, 0xe2);
    
    public static Guid SYNCMGR_OBJECTID_ConflictStore => new(0xd78181f4, 0x2389, 0x47e4, 0xa9, 0x60, 0x60, 0xbc, 0xc2, 0xed, 0x93, 0x0b);
    
    public static Guid SYNCMGR_OBJECTID_EventLinkClick => new(0x2203bdc1, 0x1af1, 0x4082, 0x8c, 0x30, 0x28, 0x39, 0x9f, 0x41, 0x38, 0x4c);
    
    public static Guid SYNCMGR_OBJECTID_EventStore => new(0x4bef34b9, 0xa786, 0x4075, 0xba, 0x88, 0x0c, 0x2b, 0x9d, 0x89, 0xa9, 0x8f);
    
    public static Guid SYNCMGR_OBJECTID_Icon => new(0x6dbc85c3, 0x5d07, 0x4c72, 0xa7, 0x77, 0x7f, 0xec, 0x78, 0x07, 0x2c, 0x06);
    
    public static Guid SYNCMGR_OBJECTID_QueryBeforeActivate => new(0xd882d80b, 0xe7aa, 0x49ed, 0x86, 0xb7, 0xe6, 0xe1, 0xf7, 0x14, 0xcd, 0xfe);
    
    public static Guid SYNCMGR_OBJECTID_QueryBeforeDeactivate => new(0xa0efc282, 0x60e0, 0x460e, 0x93, 0x74, 0xea, 0x88, 0x51, 0x3c, 0xfc, 0x80);
    
    public static Guid SYNCMGR_OBJECTID_QueryBeforeDelete => new(0xf76c3397, 0xafb3, 0x45d7, 0xa5, 0x9f, 0x5a, 0x49, 0xe9, 0x05, 0x43, 0x7e);
    
    public static Guid SYNCMGR_OBJECTID_QueryBeforeDisable => new(0xbb5f64aa, 0xf004, 0x4eb5, 0x8e, 0x4d, 0x26, 0x75, 0x19, 0x66, 0x34, 0x4c);
    
    public static Guid SYNCMGR_OBJECTID_QueryBeforeEnable => new(0x04cbf7f0, 0x5beb, 0x4de1, 0xbc, 0x90, 0x90, 0x83, 0x45, 0xc4, 0x80, 0xf6);
    
    public static Guid SYNCMGR_OBJECTID_ShowSchedule => new(0xedc6f3e3, 0x8441, 0x4109, 0xad, 0xf3, 0x6c, 0x1c, 0xa0, 0xb7, 0xde, 0x47);
    
    public static Guid SyncMgrClient => new(0x1202db60, 0x1dac, 0x42c5, 0xae, 0xd5, 0x1a, 0xbd, 0xd4, 0x32, 0x24, 0x8e);
    
    public static Guid SyncMgrControl => new(0x1a1f4206, 0x0688, 0x4e7f, 0xbe, 0x03, 0xd8, 0x2e, 0xc6, 0x9d, 0xf9, 0xa5);
    
    public static Guid SyncMgrFolder => new(0x9c73f5e5, 0x7ae7, 0x4e32, 0xa8, 0xe8, 0x8d, 0x23, 0xb8, 0x52, 0x55, 0xbf);
    
    public const uint SYNCMGRHANDLERFLAG_MASK = 15;
    
    public const uint SYNCMGRITEM_ITEMFLAGMASK = 127;
    
    public const uint SYNCMGRLOGERROR_ERRORFLAGS = 1;
    
    public const uint SYNCMGRLOGERROR_ERRORID = 2;
    
    public const uint SYNCMGRLOGERROR_ITEMID = 4;
    
    public const uint SYNCMGRPROGRESSITEM_MAXVALUE = 8;
    
    public const uint SYNCMGRPROGRESSITEM_PROGVALUE = 4;
    
    public const uint SYNCMGRPROGRESSITEM_STATUSTEXT = 1;
    
    public const uint SYNCMGRPROGRESSITEM_STATUSTYPE = 2;
    
    public const uint SYNCMGRREGISTERFLAGS_MASK = 7;
    
    public static Guid SyncMgrScheduleWizard => new(0x8d8b8e30, 0xc451, 0x421b, 0x85, 0x53, 0xd2, 0x97, 0x6a, 0xfa, 0x64, 0x8c);
    
    public static Guid SyncResultsFolder => new(0x71d99464, 0x3b6b, 0x475c, 0xb2, 0x41, 0xe1, 0x58, 0x83, 0x20, 0x75, 0x29);
    
    public static Guid SyncSetupFolder => new(0x2e9e59c0, 0xb437, 0x4981, 0xa6, 0x47, 0x9c, 0x34, 0xb9, 0xb9, 0x08, 0x91);
    
    public const uint SYNCSVC_FILTER_CALENDAR_WINDOW_WITH_RECURRENCE = 3;
    
    public const uint SYNCSVC_FILTER_CONTACTS_WITH_PHONE = 1;
    
    public const uint SYNCSVC_FILTER_NONE = 0;
    
    public const uint SYNCSVC_FILTER_TASK_ACTIVE = 2;
    
    public const string SZ_CONTENTTYPE_CDF = @"application/x-cdf";
    
    public const string SZ_CONTENTTYPE_CDFA = @"application/x-cdf";
    
    public const string SZ_CONTENTTYPE_CDFW = @"application/x-cdf";
    
    public const string SZ_CONTENTTYPE_HTML = @"text/html";
    
    public const string SZ_CONTENTTYPE_HTMLA = @"text/html";
    
    public const string SZ_CONTENTTYPE_HTMLW = @"text/html";
    
    public static Guid TaskbarList => new(0x56fdf344, 0xfd6d, 0x11d0, 0x95, 0x8a, 0x00, 0x60, 0x97, 0xc9, 0xa0, 0x90);
    
    public const uint TBIF_APPEND = 0;
    
    public const uint TBIF_DEFAULT = 0;
    
    public const uint TBIF_INTERNETBAR = 65536;
    
    public const uint TBIF_NOTOOLBAR = 196608;
    
    public const uint TBIF_PREPEND = 1;
    
    public const uint TBIF_REPLACE = 2;
    
    public const uint TBIF_STANDARDTOOLBAR = 131072;
    
    public const uint THBN_CLICKED = 6144;
    
    public static Guid ThumbnailStreamCache => new(0xcbe0fed3, 0x4b91, 0x4e90, 0x83, 0x54, 0x8a, 0x8c, 0x84, 0xec, 0x68, 0x72);
    
    public static Guid TimeCategorizer => new(0x3bb4118f, 0xddfd, 0x4d30, 0xa3, 0x48, 0x9f, 0xb5, 0xd6, 0xbf, 0x1a, 0xfe);
    
    public const uint TITLEBARNAMELEN = 40;
    
    public const uint TLMENUF_BACK = 16;
    
    public const uint TLMENUF_FORE = 32;
    
    public const uint TLMENUF_INCLUDECURRENT = 1;
    
    public const int TLOG_BACK = -1;
    
    public const uint TLOG_CURRENT = 0;
    
    public const uint TLOG_FORE = 1;
    
    public static Guid TrackShellMenu => new(0x8278f931, 0x2a3e, 0x11d2, 0x83, 0x8f, 0x00, 0xc0, 0x4f, 0xd9, 0x18, 0xd0);
    
    public static Guid TrayBandSiteService => new(0xf60ad0a0, 0xe5e1, 0x45cb, 0xb5, 0x1a, 0xe1, 0x5b, 0x9f, 0x8b, 0x29, 0x34);
    
    public static Guid TrayDeskBand => new(0xe6442437, 0x6c68, 0x4f52, 0x94, 0xdd, 0x2c, 0xfe, 0xd2, 0x67, 0xef, 0xb9);
    
    public const uint TYPE_AnchorSyncSvc = 1;
    
    public const uint TYPE_CalendarSvc = 0;
    
    public const uint TYPE_ContactsSvc = 0;
    
    public const uint TYPE_DeviceMetadataSvc = 0;
    
    public const uint TYPE_FullEnumSyncSvc = 1;
    
    public const uint TYPE_HintsSvc = 0;
    
    public const uint TYPE_MessageSvc = 0;
    
    public const uint TYPE_NotesSvc = 0;
    
    public const uint TYPE_RingtonesSvc = 0;
    
    public const uint TYPE_StatusSvc = 0;
    
    public const uint TYPE_TasksSvc = 0;
    
    public const uint URL_APPLY_DEFAULT = 1;
    
    public const uint URL_APPLY_FORCEAPPLY = 8;
    
    public const uint URL_APPLY_GUESSFILE = 4;
    
    public const uint URL_APPLY_GUESSSCHEME = 2;
    
    public const uint URL_BROWSER_MODE = 33554432;
    
    public const uint URL_CONVERT_IF_DOSPATH = 2097152;
    
    public const uint URL_DONT_ESCAPE_EXTRA_INFO = 33554432;
    
    public const uint URL_DONT_SIMPLIFY = 134217728;
    
    public const uint URL_DONT_UNESCAPE = 131072;
    
    public const uint URL_DONT_UNESCAPE_EXTRA_INFO = 33554432;
    
    public const uint URL_ESCAPE_AS_UTF8 = 262144;
    
    public const uint URL_ESCAPE_ASCII_URI_COMPONENT = 524288;
    
    public const uint URL_ESCAPE_PERCENT = 4096;
    
    public const uint URL_ESCAPE_SEGMENT_ONLY = 8192;
    
    public const uint URL_ESCAPE_SPACES_ONLY = 67108864;
    
    public const uint URL_ESCAPE_UNSAFE = 536870912;
    
    public const uint URL_FILE_USE_PATHURL = 65536;
    
    public const uint URL_INTERNAL_PATH = 8388608;
    
    public const uint URL_NO_META = 134217728;
    
    public const uint URL_PARTFLAG_KEEPSCHEME = 1;
    
    public const uint URL_PLUGGABLE_PROTOCOL = 1073741824;
    
    public const uint URL_UNESCAPE = 268435456;
    
    public const uint URL_UNESCAPE_AS_UTF8 = 262144;
    
    public const uint URL_UNESCAPE_HIGH_ANSI_ONLY = 4194304;
    
    public const uint URL_UNESCAPE_INPLACE = 1048576;
    
    public const uint URL_UNESCAPE_URI_COMPONENT = 262144;
    
    public const uint URL_WININET_COMPATIBILITY = 2147483648;
    
    public static Guid UserNotification => new(0x0010890e, 0x8789, 0x413c, 0xad, 0xbc, 0x48, 0xf5, 0xb5, 0x11, 0xb3, 0xaf);
    
    public static Guid V1PasswordCredentialProvider => new(0x6f45dc1e, 0x5384, 0x457a, 0xbc, 0x13, 0x2c, 0xd8, 0x1b, 0x0d, 0x28, 0xed);
    
    public static Guid V1SmartcardCredentialProvider => new(0x8bf9a910, 0xa8ff, 0x457f, 0x99, 0x9f, 0xa5, 0xca, 0x10, 0xb4, 0xa8, 0x85);
    
    public static Guid V1WinBioCredentialProvider => new(0xac3ac249, 0xe820, 0x4343, 0xa6, 0x5b, 0x37, 0x7a, 0xc6, 0x34, 0xdc, 0x09);
    
    public static Guid VaultProvider => new(0x503739d0, 0x4c5e, 0x4cfd, 0xb3, 0xba, 0xd8, 0x81, 0x33, 0x4f, 0x0d, 0xf2);
    
    public static Guid VID_Content => new(0x30c2c434, 0x0889, 0x4c8d, 0x98, 0x5d, 0xa9, 0xf7, 0x18, 0x30, 0xb0, 0xa9);
    
    public static Guid VID_Details => new(0x137e7700, 0x3573, 0x11cf, 0xae, 0x69, 0x08, 0x00, 0x2b, 0x2e, 0x12, 0x62);
    
    public static Guid VID_LargeIcons => new(0x0057d0e0, 0x3573, 0x11cf, 0xae, 0x69, 0x08, 0x00, 0x2b, 0x2e, 0x12, 0x62);
    
    public static Guid VID_List => new(0x0e1fa5e0, 0x3573, 0x11cf, 0xae, 0x69, 0x08, 0x00, 0x2b, 0x2e, 0x12, 0x62);
    
    public static Guid VID_SmallIcons => new(0x089000c0, 0x3573, 0x11cf, 0xae, 0x69, 0x08, 0x00, 0x2b, 0x2e, 0x12, 0x62);
    
    public static Guid VID_Thumbnails => new(0x8bebb290, 0x52d0, 0x11d0, 0xb7, 0xf4, 0x00, 0xc0, 0x4f, 0xd7, 0x06, 0xec);
    
    public static Guid VID_ThumbStrip => new(0x8eefa624, 0xd1e9, 0x445b, 0x94, 0xb7, 0x74, 0xfb, 0xce, 0x2e, 0xa1, 0x1a);
    
    public static Guid VID_Tile => new(0x65f125e5, 0x7be1, 0x4810, 0xba, 0x9d, 0xd2, 0x71, 0xc8, 0x43, 0x2c, 0xe3);
    
    public const uint VIEW_PRIORITY_CACHEHIT = 80;
    
    public const uint VIEW_PRIORITY_CACHEMISS = 48;
    
    public const uint VIEW_PRIORITY_DESPERATE = 16;
    
    public const uint VIEW_PRIORITY_INHERIT = 32;
    
    public const uint VIEW_PRIORITY_NONE = 0;
    
    public const uint VIEW_PRIORITY_RESTRICTED = 112;
    
    public const uint VIEW_PRIORITY_SHELLEXT = 64;
    
    public const uint VIEW_PRIORITY_SHELLEXT_ASBACKUP = 21;
    
    public const uint VIEW_PRIORITY_STALECACHEHIT = 69;
    
    public const uint VIEW_PRIORITY_USEASDEFAULT = 67;
    
    public static Guid VirtualDesktopManager => new(0xaa509086, 0x5ca9, 0x4c25, 0x8f, 0x95, 0x58, 0x9d, 0x3c, 0x07, 0xb4, 0x8a);
    
    public const string VOLUME_PREFIX = @"\\?\Volume";
    
    public const string WC_NETADDRESS = @"msctls_netaddress";
    
    public static Guid WebBrowser => new(0x8856f961, 0x340a, 0x11d0, 0xa9, 0x6b, 0x00, 0xc0, 0x4f, 0xd7, 0x05, 0xa2);
    
    public static Guid WebBrowser_V1 => new(0xeab22ac3, 0x30c1, 0x11cf, 0xa7, 0xeb, 0x00, 0x00, 0xc0, 0x5b, 0xae, 0x0b);
    
    public static Guid WebWizardHost => new(0xc827f149, 0x55c1, 0x4d28, 0x93, 0x5e, 0x57, 0xe4, 0x7c, 0xae, 0xd9, 0x73);
    
    public static Guid WinBioCredentialProvider => new(0xbec09223, 0xb018, 0x416d, 0xa0, 0xac, 0x52, 0x39, 0x71, 0xb6, 0x39, 0xf5);
    
    public const uint WM_CPL_LAUNCH = 2024;
    
    public const uint WM_CPL_LAUNCHED = 2025;
    
    public static Guid WPD_API_OPTIONS_V1 => new(0x10e54a3e, 0x052d, 0x4777, 0xa1, 0x3c, 0xde, 0x76, 0x14, 0xbe, 0x2b, 0xc4);
    
    public static Guid WPD_APPOINTMENT_OBJECT_PROPERTIES_V1 => new(0xf99efd03, 0x431d, 0x40d8, 0xa1, 0xc9, 0x4e, 0x22, 0x0d, 0x9c, 0x88, 0xd3);
    
    public static Guid WPD_CATEGORY_CAPABILITIES => new(0x0cabec78, 0x6b74, 0x41c6, 0x92, 0x16, 0x26, 0x39, 0xd1, 0xfc, 0xe3, 0x56);
    
    public static Guid WPD_CATEGORY_COMMON => new(0xf0422a9c, 0x5dc8, 0x4440, 0xb5, 0xbd, 0x5d, 0xf2, 0x88, 0x35, 0x65, 0x8a);
    
    public static Guid WPD_CATEGORY_DEVICE_HINTS => new(0x0d5fb92b, 0xcb46, 0x4c4f, 0x83, 0x43, 0x0b, 0xc3, 0xd3, 0xf1, 0x7c, 0x84);
    
    public static Guid WPD_CATEGORY_MEDIA_CAPTURE => new(0x59b433ba, 0xfe44, 0x4d8d, 0x80, 0x8c, 0x6b, 0xcb, 0x9b, 0x0f, 0x15, 0xe8);
    
    public static Guid WPD_CATEGORY_MTP_EXT_VENDOR_OPERATIONS => new(0x4d545058, 0x1a2e, 0x4106, 0xa3, 0x57, 0x77, 0x1e, 0x08, 0x19, 0xfc, 0x56);
    
    public static Guid WPD_CATEGORY_NETWORK_CONFIGURATION => new(0x78f9c6fc, 0x79b8, 0x473c, 0x90, 0x60, 0x6b, 0xd2, 0x3d, 0xd0, 0x72, 0xc4);
    
    public static Guid WPD_CATEGORY_NULL => new(0x00000000, 0x0000, 0x0000, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00);
    
    public static Guid WPD_CATEGORY_OBJECT_ENUMERATION => new(0xb7474e91, 0xe7f8, 0x4ad9, 0xb4, 0x00, 0xad, 0x1a, 0x4b, 0x58, 0xee, 0xec);
    
    public static Guid WPD_CATEGORY_OBJECT_MANAGEMENT => new(0xef1e43dd, 0xa9ed, 0x4341, 0x8b, 0xcc, 0x18, 0x61, 0x92, 0xae, 0xa0, 0x89);
    
    public static Guid WPD_CATEGORY_OBJECT_PROPERTIES => new(0x9e5582e4, 0x0814, 0x44e6, 0x98, 0x1a, 0xb2, 0x99, 0x8d, 0x58, 0x38, 0x04);
    
    public static Guid WPD_CATEGORY_OBJECT_PROPERTIES_BULK => new(0x11c824dd, 0x04cd, 0x4e4e, 0x8c, 0x7b, 0xf6, 0xef, 0xb7, 0x94, 0xd8, 0x4e);
    
    public static Guid WPD_CATEGORY_OBJECT_RESOURCES => new(0xb3a2b22d, 0xa595, 0x4108, 0xbe, 0x0a, 0xfc, 0x3c, 0x96, 0x5f, 0x3d, 0x4a);
    
    public static Guid WPD_CATEGORY_SERVICE_CAPABILITIES => new(0x24457e74, 0x2e9f, 0x44f9, 0x8c, 0x57, 0x1d, 0x1b, 0xcb, 0x17, 0x0b, 0x89);
    
    public static Guid WPD_CATEGORY_SERVICE_COMMON => new(0x322f071d, 0x36ef, 0x477f, 0xb4, 0xb5, 0x6f, 0x52, 0xd7, 0x34, 0xba, 0xee);
    
    public static Guid WPD_CATEGORY_SERVICE_METHODS => new(0x2d521ca8, 0xc1b0, 0x4268, 0xa3, 0x42, 0xcf, 0x19, 0x32, 0x15, 0x69, 0xbc);
    
    public static Guid WPD_CATEGORY_SMS => new(0xafc25d66, 0xfe0d, 0x4114, 0x90, 0x97, 0x97, 0x0c, 0x93, 0xe9, 0x20, 0xd1);
    
    public static Guid WPD_CATEGORY_STILL_IMAGE_CAPTURE => new(0x4fcd6982, 0x22a2, 0x4b05, 0xa4, 0x8b, 0x62, 0xd3, 0x8b, 0xf2, 0x7b, 0x32);
    
    public static Guid WPD_CATEGORY_STORAGE => new(0xd8f907a6, 0x34cc, 0x45fa, 0x97, 0xfb, 0xd0, 0x07, 0xfa, 0x47, 0xec, 0x94);
    
    public static Guid WPD_CLASS_EXTENSION_OPTIONS_V1 => new(0x6309ffef, 0xa87c, 0x4ca7, 0x84, 0x34, 0x79, 0x75, 0x76, 0xe4, 0x0a, 0x96);
    
    public static Guid WPD_CLASS_EXTENSION_OPTIONS_V2 => new(0x3e3595da, 0x4d71, 0x49fe, 0xa0, 0xb4, 0xd4, 0x40, 0x6c, 0x3a, 0xe9, 0x3f);
    
    public static Guid WPD_CLASS_EXTENSION_OPTIONS_V3 => new(0x65c160f8, 0x1367, 0x4ce2, 0x93, 0x9d, 0x83, 0x10, 0x83, 0x9f, 0x0d, 0x30);
    
    public static Guid WPD_CLASS_EXTENSION_V1 => new(0x33fb0d11, 0x64a3, 0x4fac, 0xb4, 0xc7, 0x3d, 0xfe, 0xaa, 0x99, 0xb0, 0x51);
    
    public static Guid WPD_CLASS_EXTENSION_V2 => new(0x7f0779b5, 0xfa2b, 0x4766, 0x9c, 0xb2, 0xf7, 0x3b, 0xa3, 0x0b, 0x67, 0x58);
    
    public static Guid WPD_CLIENT_INFORMATION_PROPERTIES_V1 => new(0x204d9f0c, 0x2292, 0x4080, 0x9f, 0x42, 0x40, 0x66, 0x4e, 0x70, 0xf8, 0x59);
    
    public static Guid WPD_COMMON_INFORMATION_OBJECT_PROPERTIES_V1 => new(0xb28ae94b, 0x05a4, 0x4e8e, 0xbe, 0x01, 0x72, 0xcc, 0x7e, 0x09, 0x9d, 0x8f);
    
    public static Guid WPD_CONTACT_OBJECT_PROPERTIES_V1 => new(0xfbd4fdab, 0x987d, 0x4777, 0xb3, 0xf9, 0x72, 0x61, 0x85, 0xa9, 0x31, 0x2b);
    
    public static Guid WPD_CONTENT_TYPE_ALL => new(0x80e170d2, 0x1055, 0x4a3e, 0xb9, 0x52, 0x82, 0xcc, 0x4f, 0x8a, 0x86, 0x89);
    
    public static Guid WPD_CONTENT_TYPE_APPOINTMENT => new(0x0fed060e, 0x8793, 0x4b1e, 0x90, 0xc9, 0x48, 0xac, 0x38, 0x9a, 0xc6, 0x31);
    
    public static Guid WPD_CONTENT_TYPE_AUDIO => new(0x4ad2c85e, 0x5e2d, 0x45e5, 0x88, 0x64, 0x4f, 0x22, 0x9e, 0x3c, 0x6c, 0xf0);
    
    public static Guid WPD_CONTENT_TYPE_AUDIO_ALBUM => new(0xaa18737e, 0x5009, 0x48fa, 0xae, 0x21, 0x85, 0xf2, 0x43, 0x83, 0xb4, 0xe6);
    
    public static Guid WPD_CONTENT_TYPE_CALENDAR => new(0xa1fd5967, 0x6023, 0x49a0, 0x9d, 0xf1, 0xf8, 0x06, 0x0b, 0xe7, 0x51, 0xb0);
    
    public static Guid WPD_CONTENT_TYPE_CERTIFICATE => new(0xdc3876e8, 0xa948, 0x4060, 0x90, 0x50, 0xcb, 0xd7, 0x7e, 0x8a, 0x3d, 0x87);
    
    public static Guid WPD_CONTENT_TYPE_CONTACT => new(0xeaba8313, 0x4525, 0x4707, 0x9f, 0x0e, 0x87, 0xc6, 0x80, 0x8e, 0x94, 0x35);
    
    public static Guid WPD_CONTENT_TYPE_CONTACT_GROUP => new(0x346b8932, 0x4c36, 0x40d8, 0x94, 0x15, 0x18, 0x28, 0x29, 0x1f, 0x9d, 0xe9);
    
    public static Guid WPD_CONTENT_TYPE_DOCUMENT => new(0x680adf52, 0x950a, 0x4041, 0x9b, 0x41, 0x65, 0xe3, 0x93, 0x64, 0x81, 0x55);
    
    public static Guid WPD_CONTENT_TYPE_EMAIL => new(0x8038044a, 0x7e51, 0x4f8f, 0x88, 0x3d, 0x1d, 0x06, 0x23, 0xd1, 0x45, 0x33);
    
    public static Guid WPD_CONTENT_TYPE_FOLDER => new(0x27e2e392, 0xa111, 0x48e0, 0xab, 0x0c, 0xe1, 0x77, 0x05, 0xa0, 0x5f, 0x85);
    
    public static Guid WPD_CONTENT_TYPE_FUNCTIONAL_OBJECT => new(0x99ed0160, 0x17ff, 0x4c44, 0x9d, 0x98, 0x1d, 0x7a, 0x6f, 0x94, 0x19, 0x21);
    
    public static Guid WPD_CONTENT_TYPE_GENERIC_FILE => new(0x0085e0a6, 0x8d34, 0x45d7, 0xbc, 0x5c, 0x44, 0x7e, 0x59, 0xc7, 0x3d, 0x48);
    
    public static Guid WPD_CONTENT_TYPE_GENERIC_MESSAGE => new(0xe80eaaf8, 0xb2db, 0x4133, 0xb6, 0x7e, 0x1b, 0xef, 0x4b, 0x4a, 0x6e, 0x5f);
    
    public static Guid WPD_CONTENT_TYPE_IMAGE => new(0xef2107d5, 0xa52a, 0x4243, 0xa2, 0x6b, 0x62, 0xd4, 0x17, 0x6d, 0x76, 0x03);
    
    public static Guid WPD_CONTENT_TYPE_IMAGE_ALBUM => new(0x75793148, 0x15f5, 0x4a30, 0xa8, 0x13, 0x54, 0xed, 0x8a, 0x37, 0xe2, 0x26);
    
    public static Guid WPD_CONTENT_TYPE_MEDIA_CAST => new(0x5e88b3cc, 0x3e65, 0x4e62, 0xbf, 0xff, 0x22, 0x94, 0x95, 0x25, 0x3a, 0xb0);
    
    public static Guid WPD_CONTENT_TYPE_MEMO => new(0x9cd20ecf, 0x3b50, 0x414f, 0xa6, 0x41, 0xe4, 0x73, 0xff, 0xe4, 0x57, 0x51);
    
    public static Guid WPD_CONTENT_TYPE_MIXED_CONTENT_ALBUM => new(0x00f0c3ac, 0xa593, 0x49ac, 0x92, 0x19, 0x24, 0xab, 0xca, 0x5a, 0x25, 0x63);
    
    public static Guid WPD_CONTENT_TYPE_NETWORK_ASSOCIATION => new(0x031da7ee, 0x18c8, 0x4205, 0x84, 0x7e, 0x89, 0xa1, 0x12, 0x61, 0xd0, 0xf3);
    
    public static Guid WPD_CONTENT_TYPE_PLAYLIST => new(0x1a33f7e4, 0xaf13, 0x48f5, 0x99, 0x4e, 0x77, 0x36, 0x9d, 0xfe, 0x04, 0xa3);
    
    public static Guid WPD_CONTENT_TYPE_PROGRAM => new(0xd269f96a, 0x247c, 0x4bff, 0x98, 0xfb, 0x97, 0xf3, 0xc4, 0x92, 0x20, 0xe6);
    
    public static Guid WPD_CONTENT_TYPE_SECTION => new(0x821089f5, 0x1d91, 0x4dc9, 0xbe, 0x3c, 0xbb, 0xb1, 0xb3, 0x5b, 0x18, 0xce);
    
    public static Guid WPD_CONTENT_TYPE_TASK => new(0x63252f2c, 0x887f, 0x4cb6, 0xb1, 0xac, 0xd2, 0x98, 0x55, 0xdc, 0xef, 0x6c);
    
    public static Guid WPD_CONTENT_TYPE_TELEVISION => new(0x60a169cf, 0xf2ae, 0x4e21, 0x93, 0x75, 0x96, 0x77, 0xf1, 0x1c, 0x1c, 0x6e);
    
    public static Guid WPD_CONTENT_TYPE_UNSPECIFIED => new(0x28d8d31e, 0x249c, 0x454e, 0xaa, 0xbc, 0x34, 0x88, 0x31, 0x68, 0xe6, 0x34);
    
    public static Guid WPD_CONTENT_TYPE_VIDEO => new(0x9261b03c, 0x3d78, 0x4519, 0x85, 0xe3, 0x02, 0xc5, 0xe1, 0xf5, 0x0b, 0xb9);
    
    public static Guid WPD_CONTENT_TYPE_VIDEO_ALBUM => new(0x012b0db7, 0xd4c1, 0x45d6, 0xb0, 0x81, 0x94, 0xb8, 0x77, 0x79, 0x61, 0x4f);
    
    public static Guid WPD_CONTENT_TYPE_WIRELESS_PROFILE => new(0x0bac070a, 0x9f5f, 0x4da4, 0xa8, 0xf6, 0x3d, 0xe4, 0x4d, 0x68, 0xfd, 0x6c);
    
    public const uint WPD_CONTROL_FUNCTION_GENERIC_MESSAGE = 66;
    
    public const string WPD_DEVICE_OBJECT_ID = @"DEVICE";
    
    public static Guid WPD_DEVICE_PROPERTIES_V1 => new(0x26d4979a, 0xe643, 0x4626, 0x9e, 0x2b, 0x73, 0x6d, 0xc0, 0xc9, 0x2f, 0xdc);
    
    public static Guid WPD_DEVICE_PROPERTIES_V2 => new(0x463dd662, 0x7fc4, 0x4291, 0x91, 0x1c, 0x7f, 0x4c, 0x9c, 0xca, 0x97, 0x99);
    
    public static Guid WPD_DEVICE_PROPERTIES_V3 => new(0x6c2b878c, 0xc2ec, 0x490d, 0xb4, 0x25, 0xd7, 0xa7, 0x5e, 0x23, 0xe5, 0xed);
    
    public static Guid WPD_DOCUMENT_OBJECT_PROPERTIES_V1 => new(0x0b110203, 0xeb95, 0x4f02, 0x93, 0xe0, 0x97, 0xc6, 0x31, 0x49, 0x3a, 0xd5);
    
    public static Guid WPD_EMAIL_OBJECT_PROPERTIES_V1 => new(0x41f8f65a, 0x5484, 0x4782, 0xb1, 0x3d, 0x47, 0x40, 0xdd, 0x7c, 0x37, 0xc5);
    
    public static Guid WPD_EVENT_ATTRIBUTES_V1 => new(0x10c96578, 0x2e81, 0x4111, 0xad, 0xde, 0xe0, 0x8c, 0xa6, 0x13, 0x8f, 0x6d);
    
    public static Guid WPD_EVENT_DEVICE_CAPABILITIES_UPDATED => new(0x36885aa1, 0xcd54, 0x4daa, 0xb3, 0xd0, 0xaf, 0xb3, 0xe0, 0x3f, 0x59, 0x99);
    
    public static Guid WPD_EVENT_DEVICE_REMOVED => new(0xe4cbca1b, 0x6918, 0x48b9, 0x85, 0xee, 0x02, 0xbe, 0x7c, 0x85, 0x0a, 0xf9);
    
    public static Guid WPD_EVENT_DEVICE_RESET => new(0x7755cf53, 0xc1ed, 0x44f3, 0xb5, 0xa2, 0x45, 0x1e, 0x2c, 0x37, 0x6b, 0x27);
    
    public static Guid WPD_EVENT_MTP_VENDOR_EXTENDED_EVENTS => new(0x00000000, 0x5738, 0x4ff2, 0x84, 0x45, 0xbe, 0x31, 0x26, 0x69, 0x10, 0x59);
    
    public static Guid WPD_EVENT_NOTIFICATION => new(0x2ba2e40a, 0x6b4c, 0x4295, 0xbb, 0x43, 0x26, 0x32, 0x2b, 0x99, 0xae, 0xb2);
    
    public static Guid WPD_EVENT_OBJECT_ADDED => new(0xa726da95, 0xe207, 0x4b02, 0x8d, 0x44, 0xbe, 0xf2, 0xe8, 0x6c, 0xbf, 0xfc);
    
    public static Guid WPD_EVENT_OBJECT_REMOVED => new(0xbe82ab88, 0xa52c, 0x4823, 0x96, 0xe5, 0xd0, 0x27, 0x26, 0x71, 0xfc, 0x38);
    
    public static Guid WPD_EVENT_OBJECT_TRANSFER_REQUESTED => new(0x8d16a0a1, 0xf2c6, 0x41da, 0x8f, 0x19, 0x5e, 0x53, 0x72, 0x1a, 0xdb, 0xf2);
    
    public static Guid WPD_EVENT_OBJECT_UPDATED => new(0x1445a759, 0x2e01, 0x485d, 0x9f, 0x27, 0xff, 0x07, 0xda, 0xe6, 0x97, 0xab);
    
    public static Guid WPD_EVENT_OPTIONS_V1 => new(0xb3d8dad7, 0xa361, 0x4b83, 0x8a, 0x48, 0x5b, 0x02, 0xce, 0x10, 0x71, 0x3b);
    
    public static Guid WPD_EVENT_PROPERTIES_V1 => new(0x15ab1953, 0xf817, 0x4fef, 0xa9, 0x21, 0x56, 0x76, 0xe8, 0x38, 0xf6, 0xe0);
    
    public static Guid WPD_EVENT_PROPERTIES_V2 => new(0x52807b8a, 0x4914, 0x4323, 0x9b, 0x9a, 0x74, 0xf6, 0x54, 0xb2, 0xb8, 0x46);
    
    public static Guid WPD_EVENT_SERVICE_METHOD_COMPLETE => new(0x8a33f5f8, 0x0acc, 0x4d9b, 0x9c, 0xc4, 0x11, 0x2d, 0x35, 0x3b, 0x86, 0xca);
    
    public static Guid WPD_EVENT_STORAGE_FORMAT => new(0x3782616b, 0x22bc, 0x4474, 0xa2, 0x51, 0x30, 0x70, 0xf8, 0xd3, 0x88, 0x57);
    
    public static Guid WPD_FOLDER_OBJECT_PROPERTIES_V1 => new(0x7e9a7abf, 0xe568, 0x4b34, 0xaa, 0x2f, 0x13, 0xbb, 0x12, 0xab, 0x17, 0x7d);
    
    public static Guid WPD_FORMAT_ATTRIBUTES_V1 => new(0xa0a02000, 0xbcaf, 0x4be8, 0xb3, 0xf5, 0x23, 0x3f, 0x23, 0x1c, 0xf5, 0x8f);
    
    public static Guid WPD_FUNCTIONAL_CATEGORY_ALL => new(0x2d8a6512, 0xa74c, 0x448e, 0xba, 0x8a, 0xf4, 0xac, 0x07, 0xc4, 0x93, 0x99);
    
    public static Guid WPD_FUNCTIONAL_CATEGORY_AUDIO_CAPTURE => new(0x3f2a1919, 0xc7c2, 0x4a00, 0x85, 0x5d, 0xf5, 0x7c, 0xf0, 0x6d, 0xeb, 0xbb);
    
    public static Guid WPD_FUNCTIONAL_CATEGORY_DEVICE => new(0x08ea466b, 0xe3a4, 0x4336, 0xa1, 0xf3, 0xa4, 0x4d, 0x2b, 0x5c, 0x43, 0x8c);
    
    public static Guid WPD_FUNCTIONAL_CATEGORY_NETWORK_CONFIGURATION => new(0x48f4db72, 0x7c6a, 0x4ab0, 0x9e, 0x1a, 0x47, 0x0e, 0x3c, 0xdb, 0xf2, 0x6a);
    
    public static Guid WPD_FUNCTIONAL_CATEGORY_RENDERING_INFORMATION => new(0x08600ba4, 0xa7ba, 0x4a01, 0xab, 0x0e, 0x00, 0x65, 0xd0, 0xa3, 0x56, 0xd3);
    
    public static Guid WPD_FUNCTIONAL_CATEGORY_SMS => new(0x0044a0b1, 0xc1e9, 0x4afd, 0xb3, 0x58, 0xa6, 0x2c, 0x61, 0x17, 0xc9, 0xcf);
    
    public static Guid WPD_FUNCTIONAL_CATEGORY_STILL_IMAGE_CAPTURE => new(0x613ca327, 0xab93, 0x4900, 0xb4, 0xfa, 0x89, 0x5b, 0xb5, 0x87, 0x4b, 0x79);
    
    public static Guid WPD_FUNCTIONAL_CATEGORY_STORAGE => new(0x23f05bbc, 0x15de, 0x4c2a, 0xa5, 0x5b, 0xa9, 0xaf, 0x5c, 0xe4, 0x12, 0xef);
    
    public static Guid WPD_FUNCTIONAL_CATEGORY_VIDEO_CAPTURE => new(0xe23e5f6b, 0x7243, 0x43aa, 0x8d, 0xf1, 0x0e, 0xb3, 0xd9, 0x68, 0xa9, 0x18);
    
    public static Guid WPD_FUNCTIONAL_OBJECT_PROPERTIES_V1 => new(0x8f052d93, 0xabca, 0x4fc5, 0xa5, 0xac, 0xb0, 0x1d, 0xf4, 0xdb, 0xe5, 0x98);
    
    public static Guid WPD_IMAGE_OBJECT_PROPERTIES_V1 => new(0x63d64908, 0x9fa1, 0x479f, 0x85, 0xba, 0x99, 0x52, 0x21, 0x64, 0x47, 0xdb);
    
    public static Guid WPD_MEDIA_PROPERTIES_V1 => new(0x2ed8ba05, 0x0ad3, 0x42dc, 0xb0, 0xd0, 0xbc, 0x95, 0xac, 0x39, 0x6a, 0xc8);
    
    public static Guid WPD_MEMO_OBJECT_PROPERTIES_V1 => new(0x5ffbfc7b, 0x7483, 0x41ad, 0xaf, 0xb9, 0xda, 0x3f, 0x4e, 0x59, 0x2b, 0x8d);
    
    public static Guid WPD_METHOD_ATTRIBUTES_V1 => new(0xf17a5071, 0xf039, 0x44af, 0x8e, 0xfe, 0x43, 0x2c, 0xf3, 0x2e, 0x43, 0x2a);
    
    public static Guid WPD_MUSIC_OBJECT_PROPERTIES_V1 => new(0xb324f56a, 0xdc5d, 0x46e5, 0xb6, 0xdf, 0xd2, 0xea, 0x41, 0x48, 0x88, 0xc6);
    
    public static Guid WPD_NETWORK_ASSOCIATION_PROPERTIES_V1 => new(0xe4c93c1f, 0xb203, 0x43f1, 0xa1, 0x00, 0x5a, 0x07, 0xd1, 0x1b, 0x02, 0x74);
    
    public static Guid WPD_OBJECT_FORMAT_3G2 => new(0xb9850000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_3G2A => new(0x1a11202d, 0x8759, 0x4e34, 0xba, 0x5e, 0xb1, 0x21, 0x10, 0x87, 0xee, 0xe4);
    
    public static Guid WPD_OBJECT_FORMAT_3GP => new(0xb9840000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_3GPA => new(0xe5172730, 0xf971, 0x41ef, 0xa1, 0x0b, 0x22, 0x71, 0xa0, 0x01, 0x9d, 0x7a);
    
    public static Guid WPD_OBJECT_FORMAT_AAC => new(0xb9030000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_ABSTRACT_CONTACT => new(0xbb810000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_ABSTRACT_CONTACT_GROUP => new(0xba060000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_ABSTRACT_MEDIA_CAST => new(0xba0b0000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_AIFF => new(0x30070000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_ALL => new(0xc1f62eb2, 0x4bb3, 0x479c, 0x9c, 0xfa, 0x05, 0xb5, 0xf3, 0xa5, 0x7b, 0x22);
    
    public static Guid WPD_OBJECT_FORMAT_AMR => new(0xb9080000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_ASF => new(0x300c0000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_ASXPLAYLIST => new(0xba130000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_ATSCTS => new(0xb9870000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_AUDIBLE => new(0xb9040000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_AVCHD => new(0xb9860000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_AVI => new(0x300a0000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_BMP => new(0x38040000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_CIFF => new(0x38050000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_DPOF => new(0x30060000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_DVBTS => new(0xb9880000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_EXECUTABLE => new(0x30030000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_EXIF => new(0x38010000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_FLAC => new(0xb9060000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_FLASHPIX => new(0x38030000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_GIF => new(0x38070000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_HTML => new(0x30050000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_ICALENDAR => new(0xbe030000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_ICON => new(0x077232ed, 0x102c, 0x4638, 0x9c, 0x22, 0x83, 0xf1, 0x42, 0xbf, 0xc8, 0x22);
    
    public static Guid WPD_OBJECT_FORMAT_JFIF => new(0x38080000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_JP2 => new(0x380f0000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_JPEGXR => new(0xb8040000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_JPX => new(0x38100000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_M3UPLAYLIST => new(0xba110000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_M4A => new(0x30aba7ac, 0x6ffd, 0x4c23, 0xa3, 0x59, 0x3e, 0x9b, 0x52, 0xf3, 0xf1, 0xc8);
    
    public static Guid WPD_OBJECT_FORMAT_MHT_COMPILED_HTML => new(0xba840000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_MICROSOFT_EXCEL => new(0xba850000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_MICROSOFT_POWERPOINT => new(0xba860000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_MICROSOFT_WFC => new(0xb1040000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_MICROSOFT_WORD => new(0xba830000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_MKV => new(0xb9900000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_MP2 => new(0xb9830000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_MP3 => new(0x30090000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_MP4 => new(0xb9820000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_MPEG => new(0x300b0000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_MPLPLAYLIST => new(0xba120000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_NETWORK_ASSOCIATION => new(0xb1020000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_OGG => new(0xb9020000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_PCD => new(0x38090000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_PICT => new(0x380a0000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_PLSPLAYLIST => new(0xba140000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_PNG => new(0x380b0000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_PROPERTIES_ONLY => new(0x30010000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_QCELP => new(0xb9070000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_SCRIPT => new(0x30020000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_TEXT => new(0x30040000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_TIFF => new(0x380d0000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_TIFFEP => new(0x38020000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_TIFFIT => new(0x380e0000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_UNSPECIFIED => new(0x30000000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_VCALENDAR1 => new(0xbe020000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_VCARD2 => new(0xbb820000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_VCARD3 => new(0xbb830000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_WAVE => new(0x30080000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_WBMP => new(0xb8030000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_WINDOWSIMAGEFORMAT => new(0xb8810000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_WMA => new(0xb9010000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_WMV => new(0xb9810000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_WPLPLAYLIST => new(0xba100000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_X509V3CERTIFICATE => new(0xb1030000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_FORMAT_XML => new(0xba820000, 0xae6c, 0x4804, 0x98, 0xba, 0xc5, 0x7b, 0x46, 0x96, 0x5f, 0xe7);
    
    public static Guid WPD_OBJECT_PROPERTIES_V1 => new(0xef6b490d, 0x5cd8, 0x437a, 0xaf, 0xfc, 0xda, 0x8b, 0x60, 0xee, 0x4a, 0x3c);
    
    public static Guid WPD_OBJECT_PROPERTIES_V2 => new(0x0373cd3d, 0x4a46, 0x40d7, 0xb4, 0xd8, 0x73, 0xe8, 0xda, 0x74, 0xe7, 0x75);
    
    public static Guid WPD_PARAMETER_ATTRIBUTES_V1 => new(0xe6864dd7, 0xf325, 0x45ea, 0xa1, 0xd5, 0x97, 0xcf, 0x73, 0xb6, 0xca, 0x58);
    
    public static Guid WPD_PROPERTIES_MTP_VENDOR_EXTENDED_DEVICE_PROPS => new(0x4d545058, 0x8900, 0x40b3, 0x8f, 0x1d, 0xdc, 0x24, 0x6e, 0x1e, 0x83, 0x70);
    
    public static Guid WPD_PROPERTIES_MTP_VENDOR_EXTENDED_OBJECT_PROPS => new(0x4d545058, 0x4fce, 0x4578, 0x95, 0xc8, 0x86, 0x98, 0xa9, 0xbc, 0x0f, 0x49);
    
    public static Guid WPD_PROPERTY_ATTRIBUTES_V1 => new(0xab7943d8, 0x6332, 0x445f, 0xa0, 0x0d, 0x8d, 0x5e, 0xf1, 0xe9, 0x6f, 0x37);
    
    public static Guid WPD_PROPERTY_ATTRIBUTES_V2 => new(0x5d9da160, 0x74ae, 0x43cc, 0x85, 0xa9, 0xfe, 0x55, 0x5a, 0x80, 0x79, 0x8e);
    
    public static Guid WPD_RENDERING_INFORMATION_OBJECT_PROPERTIES_V1 => new(0xc53d039f, 0xee23, 0x4a31, 0x85, 0x90, 0x76, 0x39, 0x87, 0x98, 0x70, 0xb4);
    
    public static Guid WPD_RESOURCE_ATTRIBUTES_V1 => new(0x1eb6f604, 0x9278, 0x429f, 0x93, 0xcc, 0x5b, 0xb8, 0xc0, 0x66, 0x56, 0xb6);
    
    public static Guid WPD_SECTION_OBJECT_PROPERTIES_V1 => new(0x516afd2b, 0xc64e, 0x44f0, 0x98, 0xdc, 0xbe, 0xe1, 0xc8, 0x8f, 0x7d, 0x66);
    
    public static Guid WPD_SERVICE_PROPERTIES_V1 => new(0x7510698a, 0xcb54, 0x481c, 0xb8, 0xdb, 0x0d, 0x75, 0xc9, 0x3f, 0x1c, 0x06);
    
    public static Guid WPD_SMS_OBJECT_PROPERTIES_V1 => new(0x7e1074cc, 0x50ff, 0x4dd1, 0xa7, 0x42, 0x53, 0xbe, 0x6f, 0x09, 0x3a, 0x0d);
    
    public static Guid WPD_STILL_IMAGE_CAPTURE_OBJECT_PROPERTIES_V1 => new(0x58c571ec, 0x1bcb, 0x42a7, 0x8a, 0xc5, 0xbb, 0x29, 0x15, 0x73, 0xa2, 0x60);
    
    public static Guid WPD_STORAGE_OBJECT_PROPERTIES_V1 => new(0x01a3057a, 0x74d6, 0x4e80, 0xbe, 0xa7, 0xdc, 0x4c, 0x21, 0x2c, 0xe5, 0x0a);
    
    public static Guid WPD_TASK_OBJECT_PROPERTIES_V1 => new(0xe354e95e, 0xd8a0, 0x4637, 0xa0, 0x3a, 0x0c, 0xb2, 0x68, 0x38, 0xdb, 0xc7);
    
    public static Guid WPD_VIDEO_OBJECT_PROPERTIES_V1 => new(0x346f2163, 0xf998, 0x4146, 0x8b, 0x01, 0xd1, 0x9b, 0x4c, 0x00, 0xde, 0x9a);
    
    public static Guid WPDNSE_OBJECT_PROPERTIES_V1 => new(0x34d71409, 0x4b47, 0x4d80, 0xaa, 0xac, 0x3a, 0x28, 0xa4, 0xa3, 0xb3, 0xe6);
    
    public const uint WPDNSE_PROPSHEET_CONTENT_DETAILS = 32;
    
    public const uint WPDNSE_PROPSHEET_CONTENT_GENERAL = 4;
    
    public const uint WPDNSE_PROPSHEET_CONTENT_REFERENCES = 8;
    
    public const uint WPDNSE_PROPSHEET_CONTENT_RESOURCES = 16;
    
    public const uint WPDNSE_PROPSHEET_DEVICE_GENERAL = 1;
    
    public const uint WPDNSE_PROPSHEET_STORAGE_GENERAL = 2;
    
    public static Guid WpdSerializer => new(0x0b91a74b, 0xad7c, 0x4a9d, 0xb5, 0x63, 0x29, 0xee, 0xf9, 0x16, 0x71, 0x72);
    
    public const uint WPSTYLE_CENTER = 0;
    
    public const uint WPSTYLE_CROPTOFIT = 4;
    
    public const uint WPSTYLE_KEEPASPECT = 3;
    
    public const uint WPSTYLE_MAX = 6;
    
    public const uint WPSTYLE_SPAN = 5;
    
    public const uint WPSTYLE_STRETCH = 2;
    
    public const uint WPSTYLE_TILE = 1;
    
}
