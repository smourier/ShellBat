#nullable enable
namespace ShellN.PropertyKeys;

public static class Microsoft
{
    public static class OneNote
    {
        /// <summary>
        /// <b>Microsoft.OneNote.LinkedNoteUri</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LinkedNoteUri => new(new("641064ba-9329-47e6-8f36-5fa81aa461a0"), 4);
        
        /// <summary>
        /// <b>Microsoft.OneNote.PageEditHistory</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PageEditHistory => new(new("641064ba-9329-47e6-8f36-5fa81aa461a0"), 2);
        
        /// <summary>
        /// <b>Microsoft.OneNote.TaggedNotes</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TaggedNotes => new(new("641064ba-9329-47e6-8f36-5fa81aa461a0"), 3);
    }
    
    public static class Visio
    {
        /// <summary>
        /// <b>Microsoft.Visio.MastersDetails</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MastersDetails => new(new("a4790b72-7113-4348-97ea-292bbc1f6770"), 6);
        
        /// <summary>
        /// <b>Microsoft.Visio.MastersKeywords</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MastersKeywords => new(new("a4790b72-7113-4348-97ea-292bbc1f6770"), 5);
    }
}

public static class System
{
    /// <summary>
    /// <b>System.AcquisitionID</b> of <b>VT_I4</b> type.
    /// </summary>
    public static PROPERTYKEY AcquisitionID => new(new("65a98875-3c80-40ab-abbc-efdaf77dbee2"), 100);
    
    /// <summary>
    /// <b>System.ActivityDate</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY ActivityDate => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 23);
    
    /// <summary>
    /// <b>System.ActivityIcon</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ActivityIcon => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 24);
    
    /// <summary>
    /// <b>System.ActivityInfo</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ActivityInfo => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 17);
    
    /// <summary>
    /// <b>System.AppZoneIdentifier</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY AppZoneIdentifier => new(new("502cfeab-47eb-459c-b960-e6d8728f7701"), 102);
    
    /// <summary>
    /// <b>System.ApplicationDefinedProperties</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY ApplicationDefinedProperties => new(new("cdbfc167-337e-41d8-af7c-8c09205429c7"), 100);
    
    /// <summary>
    /// <b>System.ApplicationName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ApplicationName => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 18);
    
    /// <summary>
    /// <b>System.Author</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY Author => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 4);
    
    /// <summary>
    /// <b>System.CachedFileUpdaterContentIdForConflictResolution</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY CachedFileUpdaterContentIdForConflictResolution => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 114);
    
    /// <summary>
    /// <b>System.CachedFileUpdaterContentIdForStream</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY CachedFileUpdaterContentIdForStream => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 113);
    
    /// <summary>
    /// <b>System.CameraRollDeduplicationId</b> of <b>VT_CLSID</b> type.
    /// </summary>
    public static PROPERTYKEY CameraRollDeduplicationId => new(new("75ee72ae-7d5f-482f-9487-f1c46ca819c1"), 100);
    
    /// <summary>
    /// <b>System.Capacity</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY Capacity => new(new("9b174b35-40ff-11d2-a27e-00c04fc30871"), 3);
    
    /// <summary>
    /// <b>System.Category</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY Category => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 2);
    
    /// <summary>
    /// <b>System.ChapterId</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY ChapterId => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 19);
    
    /// <summary>
    /// <b>System.CheckState</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY CheckState => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 8);
    
    /// <summary>
    /// <b>System.Comment</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY Comment => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 6);
    
    /// <summary>
    /// <b>System.Company</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY Company => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 15);
    
    /// <summary>
    /// <b>System.ComputerName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ComputerName => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 5);
    
    /// <summary>
    /// <b>System.Condition</b> of <b>VT_UNKNOWN</b> type.
    /// </summary>
    public static PROPERTYKEY Condition => new(new("f8245476-2ec6-44be-b2f7-82ec2537fa2e"), 100);
    
    /// <summary>
    /// <b>System.ConditionKey</b> of <b>VT_BLOB</b> type.
    /// </summary>
    public static PROPERTYKEY ConditionKey => new(new("f8245476-2ec6-44be-b2f7-82ec2537fa2e"), 101);
    
    /// <summary>
    /// <b>System.ContainedItems</b> of <b>VT_CLSID, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY ContainedItems => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 29);
    
    /// <summary>
    /// <b>System.ContentId</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ContentId => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 132);
    
    /// <summary>
    /// <b>System.ContentStatus</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ContentStatus => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 27);
    
    /// <summary>
    /// <b>System.ContentType</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ContentType => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 26);
    
    /// <summary>
    /// <b>System.ContentUri</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ContentUri => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 131);
    
    /// <summary>
    /// <b>System.ContentUrl</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ContentUrl => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 10);
    
    /// <summary>
    /// <b>System.Copyright</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY Copyright => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 11);
    
    /// <summary>
    /// <b>System.CreatorAppId</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY CreatorAppId => new(new("c2ea046e-033c-4e91-bd5b-d4942f6bbe49"), 2);
    
    /// <summary>
    /// <b>System.CreatorOpenWithUIOptions</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY CreatorOpenWithUIOptions => new(new("c2ea046e-033c-4e91-bd5b-d4942f6bbe49"), 3);
    
    /// <summary>
    /// <b>System.DataObjectFormat</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY DataObjectFormat => new(new("1e81a3f8-a30f-4247-b9ee-1d0368a9425c"), 2);
    
    /// <summary>
    /// <b>System.DateAccessed</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY DateAccessed => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 16);
    
    /// <summary>
    /// <b>System.DateAcquired</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY DateAcquired => new(new("2cbaa8f5-d81f-47ca-b17a-f8d822300131"), 100);
    
    /// <summary>
    /// <b>System.DateArchived</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY DateArchived => new(new("43f8d7b7-a444-4f87-9383-52271c9b915c"), 100);
    
    /// <summary>
    /// <b>System.DateCompleted</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY DateCompleted => new(new("72fab781-acda-43e5-b155-b2434f85e678"), 100);
    
    /// <summary>
    /// <b>System.DateCreated</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY DateCreated => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 15);
    
    /// <summary>
    /// <b>System.DateImported</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY DateImported => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 18258);
    
    /// <summary>
    /// <b>System.DateModified</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY DateModified => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 14);
    
    /// <summary>
    /// <b>System.DefaultGroupOrder</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY DefaultGroupOrder => new(new("83914d1a-c270-48bf-b00d-1c4e451b0150"), 100);
    
    /// <summary>
    /// <b>System.DefaultSaveLocationDisplay</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY DefaultSaveLocationDisplay => new(new("5d76b67f-9b3d-44bb-b6ae-25da4f638a67"), 10);
    
    /// <summary>
    /// <b>System.DefaultSaveLocationIconContainer</b> of <b>VT_BLOB</b> type.
    /// </summary>
    public static PROPERTYKEY DefaultSaveLocationIconContainer => new(new("5d76b67f-9b3d-44bb-b6ae-25da4f638a67"), 11);
    
    /// <summary>
    /// <b>System.DelegateIDList</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY DelegateIDList => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 32);
    
    /// <summary>
    /// <b>System.DelegateSourceId</b> of <b>VT_CLSID</b> type.
    /// </summary>
    public static PROPERTYKEY DelegateSourceId => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 11);
    
    /// <summary>
    /// <b>System.DelegationFlags</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY DelegationFlags => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 9);
    
    /// <summary>
    /// <b>System.DescriptionID</b> of <b>VT_UI1, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY DescriptionID => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 2);
    
    /// <summary>
    /// <b>System.DueDate</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY DueDate => new(new("3f8472b5-e0af-4db2-8071-c53fe76ae7ce"), 100);
    
    /// <summary>
    /// <b>System.DuiControlResource</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY DuiControlResource => new(new("ea810849-87ff-4b54-abd6-5b71adf466f8"), 1);
    
    /// <summary>
    /// <b>System.EndDate</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY EndDate => new(new("c75faa05-96fd-49e7-9cb4-9f601082d553"), 100);
    
    /// <summary>
    /// <b>System.ExpandoProperties</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY ExpandoProperties => new(new("6fa20de6-d11c-4d9d-a154-64317628c12d"), 100);
    
    /// <summary>
    /// <b>System.FileAllocationSize</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY FileAllocationSize => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 18);
    
    /// <summary>
    /// <b>System.FileAttributes</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY FileAttributes => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 13);
    
    /// <summary>
    /// <b>System.FileAttributesDisplay</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY FileAttributesDisplay => new(new("8d72aca1-0716-419a-9ac1-acb07b18dc32"), 2);
    
    /// <summary>
    /// <b>System.FileCount</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY FileCount => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 12);
    
    /// <summary>
    /// <b>System.FileDescription</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY FileDescription => new(new("0cef7d53-fa64-11d1-a203-0000f81fedee"), 3);
    
    /// <summary>
    /// <b>System.FileExtension</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY FileExtension => new(new("e4f10a3c-49e6-405d-8288-a23bd4eeaa6c"), 100);
    
    /// <summary>
    /// <b>System.FileFRN</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY FileFRN => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 21);
    
    /// <summary>
    /// <b>System.FileIndex</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY FileIndex => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 8);
    
    /// <summary>
    /// <b>System.FileName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY FileName => new(new("41cf5ae0-f75a-4806-bd87-59c7d9248eb9"), 100);
    
    /// <summary>
    /// <b>System.FileOfflineAvailabilityStatus</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY FileOfflineAvailabilityStatus => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 100);
    
    /// <summary>
    /// <b>System.FileOwner</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY FileOwner => new(new("9b174b34-40ff-11d2-a27e-00c04fc30871"), 4);
    
    /// <summary>
    /// <b>System.FilePlaceholderStatus</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY FilePlaceholderStatus => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 2);
    
    /// <summary>
    /// <b>System.FileReparsePointTag</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY FileReparsePointTag => new(new("3d75e4f5-a391-4952-81f7-c7072fe53025"), 100);
    
    /// <summary>
    /// <b>System.FileVersion</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY FileVersion => new(new("0cef7d53-fa64-11d1-a203-0000f81fedee"), 4);
    
    /// <summary>
    /// <b>System.FilterInfo</b> of <b>VT_STREAM</b> type.
    /// </summary>
    public static PROPERTYKEY FilterInfo => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 2);
    
    /// <summary>
    /// <b>System.FindData</b> of <b>VT_UI1, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY FindData => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 0);
    
    /// <summary>
    /// <b>System.FlagColor</b> of <b>VT_UI2</b> type.
    /// </summary>
    public static PROPERTYKEY FlagColor => new(new("67df94de-0ca7-4d6f-b792-053a3e4f03cf"), 100);
    
    /// <summary>
    /// <b>System.FlagColorText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY FlagColorText => new(new("45eae747-8e2a-40ae-8cbf-ca52aba6152a"), 100);
    
    /// <summary>
    /// <b>System.FlagStatus</b> of <b>VT_I4</b> type.
    /// </summary>
    public static PROPERTYKEY FlagStatus => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 12);
    
    /// <summary>
    /// <b>System.FlagStatusText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY FlagStatusText => new(new("dc54fd2e-189d-4871-aa01-08c2f57a4abc"), 100);
    
    /// <summary>
    /// <b>System.FolderKind</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY FolderKind => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 101);
    
    /// <summary>
    /// <b>System.FolderNameDisplay</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY FolderNameDisplay => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 25);
    
    /// <summary>
    /// <b>System.ForceFullText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ForceFullText => new(new("1173f62a-2a55-4f62-aed6-8c7112e0f7a3"), 5);
    
    /// <summary>
    /// <b>System.FreeSpace</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY FreeSpace => new(new("9b174b35-40ff-11d2-a27e-00c04fc30871"), 2);
    
    /// <summary>
    /// <b>System.FullText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY FullText => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 6);
    
    /// <summary>
    /// <b>System.HasLeafContainers</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY HasLeafContainers => new(new("a11c005a-ff95-4785-8617-beaf92399c3c"), 100);
    
    /// <summary>
    /// <b>System.HasNoAdditionalProperties</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY HasNoAdditionalProperties => new(new("27b381e0-61dd-449c-b739-815e0f031299"), 100);
    
    /// <summary>
    /// <b>System.HideInGrepSearch</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY HideInGrepSearch => new(new("c6f039e7-f6a4-4185-ae48-07938262c274"), 100);
    
    /// <summary>
    /// <b>System.HideOnDesktop</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY HideOnDesktop => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 34);
    
    /// <summary>
    /// <b>System.HighKeywords</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY HighKeywords => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 24);
    
    /// <summary>
    /// <b>System.HomeGroupSharingStatus</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY HomeGroupSharingStatus => new(new("4ac903f8-e780-4e4b-b7b8-4d00a99804fc"), 100);
    
    /// <summary>
    /// <b>System.IconIndex</b> of <b>VT_I4</b> type.
    /// </summary>
    public static PROPERTYKEY IconIndex => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 26);
    
    /// <summary>
    /// <b>System.IconPath</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY IconPath => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 25);
    
    /// <summary>
    /// <b>System.IconSecondaryStreamName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY IconSecondaryStreamName => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 43);
    
    /// <summary>
    /// <b>System.Identity</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY IdentityProperty => new(new("a26f4afc-7346-4299-be47-eb1ae613139f"), 100);
    
    /// <summary>
    /// <b>System.ImageParsingName</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY ImageParsingName => new(new("d7750ee0-c6a4-48ec-b53e-b87b52e6d073"), 100);
    
    /// <summary>
    /// <b>System.Importance</b> of <b>VT_I4</b> type.
    /// </summary>
    public static PROPERTYKEY Importance => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 11);
    
    /// <summary>
    /// <b>System.ImportanceText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ImportanceText => new(new("a3b29791-7713-4e1d-bb40-17db85f01831"), 100);
    
    /// <summary>
    /// <b>System.InfoTipText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY InfoTipText => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 17);
    
    /// <summary>
    /// <b>System.InternalName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY InternalName => new(new("0cef7d53-fa64-11d1-a203-0000f81fedee"), 5);
    
    /// <summary>
    /// <b>System.InvalidPathValue</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY InvalidPathValue => new(new("e5473742-4611-4aaf-9c49-a3417748cbc8"), 100);
    
    /// <summary>
    /// <b>System.IsAttachment</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsAttachment => new(new("f23f425c-71a1-4fa8-922f-678ea4a60408"), 100);
    
    /// <summary>
    /// <b>System.IsBarricadePage</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsBarricadePage => new(new("98f920d1-51e2-4722-9069-3c4b5cff5165"), 100);
    
    /// <summary>
    /// <b>System.IsDefaultNonOwnerSaveLocation</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsDefaultNonOwnerSaveLocation => new(new("5d76b67f-9b3d-44bb-b6ae-25da4f638a67"), 5);
    
    /// <summary>
    /// <b>System.IsDefaultSaveLocation</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsDefaultSaveLocation => new(new("5d76b67f-9b3d-44bb-b6ae-25da4f638a67"), 3);
    
    /// <summary>
    /// <b>System.IsDefaultSaveLocationForDisplay</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsDefaultSaveLocationForDisplay => new(new("5d76b67f-9b3d-44bb-b6ae-25da4f638a67"), 7);
    
    /// <summary>
    /// <b>System.IsDeleted</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsDeleted => new(new("5cda5fc8-33ee-4ff3-9094-ae7bd8868c4d"), 100);
    
    /// <summary>
    /// <b>System.IsEncrypted</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsEncrypted => new(new("90e5e14e-648b-4826-b2aa-acaf790e3513"), 10);
    
    /// <summary>
    /// <b>System.IsFlagged</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsFlagged => new(new("5da84765-e3ff-4278-86b0-a27967fbdd03"), 100);
    
    /// <summary>
    /// <b>System.IsFlaggedComplete</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsFlaggedComplete => new(new("a6f360d2-55f9-48de-b909-620e090a647c"), 100);
    
    /// <summary>
    /// <b>System.IsFolder</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsFolder => new(new("09329b74-40a3-4c68-bf07-af9a572f607c"), 100);
    
    /// <summary>
    /// <b>System.IsGroup</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsGroup => new(new("c64a866e-41ae-4c8c-b3d5-dd6dbf70c9c1"), 100);
    
    /// <summary>
    /// <b>System.IsIncomplete</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsIncomplete => new(new("346c8bd1-2e6a-4c45-89a4-61b78e8e700f"), 100);
    
    /// <summary>
    /// <b>System.IsLocationSupported</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsLocationSupported => new(new("5d76b67f-9b3d-44bb-b6ae-25da4f638a67"), 8);
    
    /// <summary>
    /// <b>System.IsPinnedToNameSpaceTree</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsPinnedToNameSpaceTree => new(new("5d76b67f-9b3d-44bb-b6ae-25da4f638a67"), 2);
    
    /// <summary>
    /// <b>System.IsRead</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsRead => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 10);
    
    /// <summary>
    /// <b>System.IsSearchOnlyItem</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsSearchOnlyItem => new(new("5d76b67f-9b3d-44bb-b6ae-25da4f638a67"), 4);
    
    /// <summary>
    /// <b>System.IsSendToTarget</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsSendToTarget => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 33);
    
    /// <summary>
    /// <b>System.IsShared</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsShared => new(new("ef884c5b-2bfe-41bb-aae5-76eedf4f9902"), 100);
    
    /// <summary>
    /// <b>System.IsSimpleItem</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY IsSimpleItem => new(new("32bcb03c-7f34-4e3f-bbb2-ebe63629f5e4"), 100);
    
    /// <summary>
    /// <b>System.ItemAfter</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY ItemAfter => new(new("b9b4b3fc-2b51-4a42-b5d8-324146afcf25"), 6);
    
    /// <summary>
    /// <b>System.ItemAuthors</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemAuthors => new(new("d0a04f0a-462a-48a4-bb2f-3706e88dbd7d"), 100);
    
    /// <summary>
    /// <b>System.ItemClassType</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemClassType => new(new("048658ad-2db8-41a4-bbb6-ac1ef1207eb1"), 100);
    
    /// <summary>
    /// <b>System.ItemContext</b> of <b>VT_UNKNOWN</b> type.
    /// </summary>
    public static PROPERTYKEY ItemContext => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 25);
    
    /// <summary>
    /// <b>System.ItemDate</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY ItemDate => new(new("f7db74b4-4287-4103-afba-f1b13dcd75cf"), 100);
    
    /// <summary>
    /// <b>System.ItemFolderNameDisplay</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemFolderNameDisplay => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 2);
    
    /// <summary>
    /// <b>System.ItemFolderPathDisplay</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemFolderPathDisplay => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 6);
    
    /// <summary>
    /// <b>System.ItemFolderPathDisplayNarrow</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemFolderPathDisplayNarrow => new(new("dabd30ed-0043-4789-a7f8-d013a4736622"), 100);
    
    /// <summary>
    /// <b>System.ItemId</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY ItemId => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 8);
    
    /// <summary>
    /// <b>System.ItemName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemName => new(new("6b8da074-3b5c-43bc-886f-0a2cdce00b6f"), 100);
    
    /// <summary>
    /// <b>System.ItemNameDisplay</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemNameDisplay => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 10);
    
    /// <summary>
    /// <b>System.ItemNameDisplayWithoutExtension</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemNameDisplayWithoutExtension => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 24);
    
    /// <summary>
    /// <b>System.ItemNamePrefix</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemNamePrefix => new(new("d7313ff1-a77a-401c-8c99-3dbdd68add36"), 100);
    
    /// <summary>
    /// <b>System.ItemNameSortOverride</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemNameSortOverride => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 23);
    
    /// <summary>
    /// <b>System.ItemParticipants</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemParticipants => new(new("d4d0aa16-9948-41a4-aa85-d97ff9646993"), 100);
    
    /// <summary>
    /// <b>System.ItemPathDisplay</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemPathDisplay => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 7);
    
    /// <summary>
    /// <b>System.ItemPathDisplayNarrow</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemPathDisplayNarrow => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 8);
    
    /// <summary>
    /// <b>System.ItemQueryCondition</b> of <b>VT_STREAM</b> type.
    /// </summary>
    public static PROPERTYKEY ItemQueryCondition => new(new("f50d2f5d-dda0-48d4-8d2b-e83729fb69a4"), 100);
    
    /// <summary>
    /// <b>System.ItemSearchLocation</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemSearchLocation => new(new("23620678-ccd4-47c0-9963-95a8405678a3"), 100);
    
    /// <summary>
    /// <b>System.ItemSearchLocationHeaderType</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY ItemSearchLocationHeaderType => new(new("23620678-ccd4-47c0-9963-95a8405678a3"), 101);
    
    /// <summary>
    /// <b>System.ItemSourceCLSID</b> of <b>VT_CLSID</b> type.
    /// </summary>
    public static PROPERTYKEY ItemSourceCLSID => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 10);
    
    /// <summary>
    /// <b>System.ItemSubType</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemSubType => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 37);
    
    /// <summary>
    /// <b>System.ItemType</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemType => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 11);
    
    /// <summary>
    /// <b>System.ItemTypeText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemTypeText => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 4);
    
    /// <summary>
    /// <b>System.ItemUrl</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ItemUrl => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 9);
    
    /// <summary>
    /// <b>System.ItemsInStack</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY ItemsInStack => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 7);
    
    /// <summary>
    /// <b>System.Keywords</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY Keywords => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 5);
    
    /// <summary>
    /// <b>System.Kind</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY Kind => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 3);
    
    /// <summary>
    /// <b>System.KindText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY KindText => new(new("f04bef95-c585-4197-a2b7-df46fdc9ee6d"), 100);
    
    /// <summary>
    /// <b>System.Language</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY Language => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 28);
    
    /// <summary>
    /// <b>System.LastSyncError</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY LastSyncError => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 107);
    
    /// <summary>
    /// <b>System.LastSyncWarning</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY LastSyncWarning => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 128);
    
    /// <summary>
    /// <b>System.LastWriterPackageFamilyName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY LastWriterPackageFamilyName => new(new("502cfeab-47eb-459c-b960-e6d8728f7701"), 101);
    
    /// <summary>
    /// <b>System.LibraryLocationSupportStatus</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY LibraryLocationSupportStatus => new(new("5d76b67f-9b3d-44bb-b6ae-25da4f638a67"), 9);
    
    /// <summary>
    /// <b>System.LibraryLocationsCount</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY LibraryLocationsCount => new(new("908696c7-8f87-44f2-80ed-a8c1c6894575"), 2);
    
    /// <summary>
    /// <b>System.LibraryLocationsList</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY LibraryLocationsList => new(new("908696c7-8f87-44f2-80ed-a8c1c6894575"), 4);
    
    /// <summary>
    /// <b>System.ListDescription</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY ListDescription => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 23);
    
    /// <summary>
    /// <b>System.LocationEmptyString</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY LocationEmptyString => new(new("62d2d9ab-8b64-498d-b865-402d4796f865"), 3);
    
    /// <summary>
    /// <b>System.LowKeywords</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY LowKeywords => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 25);
    
    /// <summary>
    /// <b>System.MIMEType</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY MIMEType => new(new("0b63e350-9ccc-11d0-bcdb-00805fccce04"), 5);
    
    /// <summary>
    /// <b>System.MaxStackCount</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY MaxStackCount => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 22);
    
    /// <summary>
    /// <b>System.MediumKeywords</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY MediumKeywords => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 26);
    
    /// <summary>
    /// <b>System.MileageInformation</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY MileageInformation => new(new("fdf84370-031a-4add-9e91-0d775f1c6605"), 100);
    
    /// <summary>
    /// <b>System.MusicStackThumbnailCacheIds</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY MusicStackThumbnailCacheIds => new(new("733cb147-8b1f-4c48-9966-192fde353c75"), 100);
    
    /// <summary>
    /// <b>System.NamespaceCLSID</b> of <b>VT_CLSID</b> type.
    /// </summary>
    public static PROPERTYKEY NamespaceCLSID => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 6);
    
    /// <summary>
    /// <b>System.NetworkLocation</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY NetworkLocation => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 4);
    
    /// <summary>
    /// <b>System.NetworkPlacesDefaultName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY NetworkPlacesDefaultName => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 35);
    
    /// <summary>
    /// <b>System.NetworkProvider</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY NetworkProvider => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 31);
    
    /// <summary>
    /// <b>System.NetworkResource</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY NetworkResource => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 1);
    
    /// <summary>
    /// <b>System.NewMenuAllowedTypes</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY NewMenuAllowedTypes => new(new("9b174b34-40ff-11d2-a27e-00c04fc30871"), 10);
    
    /// <summary>
    /// <b>System.NewMenuPreferredTypes</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY NewMenuPreferredTypes => new(new("9b174b34-40ff-11d2-a27e-00c04fc30871"), 8);
    
    /// <summary>
    /// <b>System.NotUserContent</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY NotUserContent => new(new("fcfb52aa-c1e5-4cd8-88bc-f80fd7390f20"), 100);
    
    /// <summary>
    /// <b>System.OfflineAvailability</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY OfflineAvailability => new(new("a94688b6-7d9f-4570-a648-e3dfc0ab2b3f"), 100);
    
    /// <summary>
    /// <b>System.OfflineStatus</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY OfflineStatus => new(new("6d24888f-4718-4bda-afed-ea0fb4386cd8"), 100);
    
    /// <summary>
    /// <b>System.OfflineSyncTime</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY OfflineSyncTime => new(new("4e9cfc01-5d36-406a-83cd-4e7423923604"), 2);
    
    /// <summary>
    /// <b>System.Order</b> of <b>VT_I4</b> type.
    /// </summary>
    public static PROPERTYKEY Order => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 26);
    
    /// <summary>
    /// <b>System.OriginalFileName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY OriginalFileName => new(new("0cef7d53-fa64-11d1-a203-0000f81fedee"), 6);
    
    /// <summary>
    /// <b>System.OwnerSID</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY OwnerSID => new(new("5d76b67f-9b3d-44bb-b6ae-25da4f638a67"), 6);
    
    /// <summary>
    /// <b>System.ParentalRating</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ParentalRating => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 21);
    
    /// <summary>
    /// <b>System.ParentalRatingReason</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ParentalRatingReason => new(new("10984e0a-f9f2-4321-b7ef-baf195af4319"), 100);
    
    /// <summary>
    /// <b>System.ParentalRatingsOrganization</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ParentalRatingsOrganization => new(new("a7fe0840-1344-46f0-8d37-52ed712a4bf9"), 100);
    
    /// <summary>
    /// <b>System.ParsingBindContext</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY ParsingBindContext => new(new("dfb9a04d-362f-4ca3-b30b-0254b17b5b84"), 100);
    
    /// <summary>
    /// <b>System.ParsingName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ParsingName => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 24);
    
    /// <summary>
    /// <b>System.ParsingPath</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ParsingPath => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 30);
    
    /// <summary>
    /// <b>System.PerceivedType</b> of <b>VT_I4</b> type.
    /// </summary>
    public static PROPERTYKEY PerceivedType => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 9);
    
    /// <summary>
    /// <b>System.PercentFull</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY PercentFull => new(new("9b174b35-40ff-11d2-a27e-00c04fc30871"), 5);
    
    /// <summary>
    /// <b>System.Platform</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY Platform => new(new("0cef7d53-fa64-11d1-a203-0000f81fedee"), 11);
    
    /// <summary>
    /// <b>System.Priority</b> of <b>VT_UI2</b> type.
    /// </summary>
    public static PROPERTYKEY Priority => new(new("9c1fcf74-2d97-41ba-b4ae-cb2e3661a6e4"), 5);
    
    /// <summary>
    /// <b>System.PriorityText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY PriorityText => new(new("d98be98b-b86b-4095-bf52-9d23b2e0a752"), 100);
    
    /// <summary>
    /// <b>System.Project</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY Project => new(new("39a7f922-477c-48de-8bc8-b28441e342e3"), 100);
    
    /// <summary>
    /// <b>System.ProviderItemID</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ProviderItemID => new(new("f21d9941-81f0-471a-adee-4e74b49217ed"), 100);
    
    /// <summary>
    /// <b>System.Rating</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY Rating => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 9);
    
    /// <summary>
    /// <b>System.RatingText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY RatingText => new(new("90197ca7-fd8f-4e8c-9da3-b57e1e609295"), 100);
    
    /// <summary>
    /// <b>System.RemoteConflictingFile</b> of <b>VT_UNKNOWN</b> type.
    /// </summary>
    public static PROPERTYKEY RemoteConflictingFile => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 115);
    
    /// <summary>
    /// <b>System.ResultSourceId</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ResultSourceId => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 17);
    
    /// <summary>
    /// <b>System.ResultType</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY ResultType => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 20);
    
    /// <summary>
    /// <b>System.SDID</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY SDID => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 2147483650);
    
    /// <summary>
    /// <b>System.SFGAOFlags</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY SFGAOFlags => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 25);
    
    /// <summary>
    /// <b>System.SearchScope</b> of <b>VT_STREAM</b> type.
    /// </summary>
    public static PROPERTYKEY SearchScope => new(new("cf5751fd-f4b3-443d-b31c-9a34740759ec"), 100);
    
    /// <summary>
    /// <b>System.Sensitivity</b> of <b>VT_UI2</b> type.
    /// </summary>
    public static PROPERTYKEY Sensitivity => new(new("f8d3f6ac-4874-42cb-be59-ab454b30716a"), 100);
    
    /// <summary>
    /// <b>System.SensitivityText</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY SensitivityText => new(new("d0c7f054-3f72-4725-8527-129a577cb269"), 100);
    
    /// <summary>
    /// <b>System.ShareScope</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY ShareScope => new(new("ef884c5b-2bfe-41bb-aae5-76eedf4f9902"), 400);
    
    /// <summary>
    /// <b>System.ShareUserRating</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY ShareUserRating => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 12);
    
    /// <summary>
    /// <b>System.SharedWith</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY SharedWith => new(new("ef884c5b-2bfe-41bb-aae5-76eedf4f9902"), 200);
    
    /// <summary>
    /// <b>System.SharingStatus</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY SharingStatus => new(new("ef884c5b-2bfe-41bb-aae5-76eedf4f9902"), 300);
    
    /// <summary>
    /// <b>System.SimpleRating</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY SimpleRating => new(new("a09f084e-ad41-489f-8076-aa5be3082bca"), 100);
    
    /// <summary>
    /// <b>System.Size</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY Size => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 12);
    
    /// <summary>
    /// <b>System.SoftwareUsed</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY SoftwareUsed => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 305);
    
    /// <summary>
    /// <b>System.SourceItem</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY SourceItem => new(new("668cdfa5-7a1b-4323-ae4b-e527393a1d81"), 100);
    
    /// <summary>
    /// <b>System.SourcePackageFamilyName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY SourcePackageFamilyName => new(new("ffae9db7-1c8d-43ff-818c-84403aa3732d"), 100);
    
    /// <summary>
    /// <b>System.StackThumbnailCacheIds</b> of <b>VT_NULL</b> type.
    /// </summary>
    public static PROPERTYKEY StackThumbnailCacheIds => new(new("4596208c-32fa-41d2-9695-af0cb9e8dcfe"), 100);
    
    /// <summary>
    /// <b>System.StartDate</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY StartDate => new(new("48fd6ec8-8a12-4cdf-a03e-4ec5a511edde"), 100);
    
    /// <summary>
    /// <b>System.Status</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY Status => new(new("000214a1-0000-0000-c000-000000000046"), 9);
    
    /// <summary>
    /// <b>System.StatusBarItemAvailability</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StatusBarItemAvailability => new(new("26dc287c-6e3d-4bd3-b2b0-6a26ba2e346d"), 5);
    
    /// <summary>
    /// <b>System.StatusBarSelectedItemCount</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StatusBarSelectedItemCount => new(new("26dc287c-6e3d-4bd3-b2b0-6a26ba2e346d"), 3);
    
    /// <summary>
    /// <b>System.StatusBarStorageProviderError</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StatusBarStorageProviderError => new(new("26dc287c-6e3d-4bd3-b2b0-6a26ba2e346d"), 6);
    
    /// <summary>
    /// <b>System.StatusBarViewItemCount</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StatusBarViewItemCount => new(new("26dc287c-6e3d-4bd3-b2b0-6a26ba2e346d"), 2);
    
    /// <summary>
    /// <b>System.StatusIcons</b> of <b>VT_BLOB</b> type.
    /// </summary>
    public static PROPERTYKEY StatusIcons => new(new("7a55582b-bd8c-4475-b94c-b87a388a7899"), 100);
    
    /// <summary>
    /// <b>System.StorageProviderAggregatedCustomStates</b> of <b>VT_UI1, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderAggregatedCustomStates => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 124);
    
    /// <summary>
    /// <b>System.StorageProviderCallerVersionInformation</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderCallerVersionInformation => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 7);
    
    /// <summary>
    /// <b>System.StorageProviderCustomPrimaryIcon</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderCustomPrimaryIcon => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 12);
    
    /// <summary>
    /// <b>System.StorageProviderCustomStates</b> of <b>VT_UI1, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderCustomStates => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 120);
    
    /// <summary>
    /// <b>System.StorageProviderDescendantSharingStatus</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderDescendantSharingStatus => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 118);
    
    /// <summary>
    /// <b>System.StorageProviderError</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderError => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 109);
    
    /// <summary>
    /// <b>System.StorageProviderFileChecksum</b> of <b>VT_UI1, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileChecksum => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 5);
    
    /// <summary>
    /// <b>System.StorageProviderFileCreatedBy</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileCreatedBy => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 10);
    
    /// <summary>
    /// <b>System.StorageProviderFileDateShared</b> of <b>VT_FILETIME</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileDateShared => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 14);
    
    /// <summary>
    /// <b>System.StorageProviderFileFlags</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileFlags => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 8);
    
    /// <summary>
    /// <b>System.StorageProviderFileHasConflict</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileHasConflict => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 9);
    
    /// <summary>
    /// <b>System.StorageProviderFileIdentifier</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileIdentifier => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 3);
    
    /// <summary>
    /// <b>System.StorageProviderFileModifiedBy</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileModifiedBy => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 11);
    
    /// <summary>
    /// <b>System.StorageProviderFileRemoteLocation</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileRemoteLocation => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 16);
    
    /// <summary>
    /// <b>System.StorageProviderFileRemoteUri</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileRemoteUri => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 112);
    
    /// <summary>
    /// <b>System.StorageProviderFileSharedBy</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileSharedBy => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 15);
    
    /// <summary>
    /// <b>System.StorageProviderFileVersion</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileVersion => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 4);
    
    /// <summary>
    /// <b>System.StorageProviderFileVersionWaterline</b> of <b>VT_UI1, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFileVersionWaterline => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 6);
    
    /// <summary>
    /// <b>System.StorageProviderFullyQualifiedId</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderFullyQualifiedId => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 119);
    
    /// <summary>
    /// <b>System.StorageProviderId</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderId => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 108);
    
    /// <summary>
    /// <b>System.StorageProviderNetworkConnected</b> of <b>VT_BOOL</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderNetworkConnected => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 125);
    
    /// <summary>
    /// <b>System.StorageProviderProtectionMode</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderProtectionMode => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 127);
    
    /// <summary>
    /// <b>System.StorageProviderShareStatuses</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderShareStatuses => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 111);
    
    /// <summary>
    /// <b>System.StorageProviderSharingStatus</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderSharingStatus => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 117);
    
    /// <summary>
    /// <b>System.StorageProviderState</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderState => new(new("e77e90df-6271-4f5b-834f-2dd1f245dda4"), 3);
    
    /// <summary>
    /// <b>System.StorageProviderStatus</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderStatus => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 110);
    
    /// <summary>
    /// <b>System.StorageProviderThumbnailDimensions</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderThumbnailDimensions => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 116);
    
    /// <summary>
    /// <b>System.StorageProviderTransferProgress</b> of <b>VT_UI4, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderTransferProgress => new(new("e77e90df-6271-4f5b-834f-2dd1f245dda4"), 4);
    
    /// <summary>
    /// <b>System.StorageProviderUIStatus</b> of <b>VT_BLOB</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderUIStatus => new(new("e77e90df-6271-4f5b-834f-2dd1f245dda4"), 2);
    
    /// <summary>
    /// <b>System.StorageProviderUserAccountKind</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderUserAccountKind => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 17);
    
    /// <summary>
    /// <b>System.StorageProviderUserId</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderUserId => new(new("b2f9b9d6-fec4-4dd5-94d7-8957488c807b"), 13);
    
    /// <summary>
    /// <b>System.StorageProviderWarningErrorState</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY StorageProviderWarningErrorState => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 126);
    
    /// <summary>
    /// <b>System.StorageSystemType</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY StorageSystemType => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 36);
    
    /// <summary>
    /// <b>System.Subject</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY Subject => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 3);
    
    /// <summary>
    /// <b>System.SyncAvailabilityFlags</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY SyncAvailabilityFlags => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 130);
    
    /// <summary>
    /// <b>System.SyncTransferStatus</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY SyncTransferStatus => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 103);
    
    /// <summary>
    /// <b>System.SyncTransferStatusFlags</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY SyncTransferStatusFlags => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 129);
    
    /// <summary>
    /// <b>System.Thumbnail</b> of <b>VT_CF</b> type.
    /// </summary>
    public static PROPERTYKEY Thumbnail => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 17);
    
    /// <summary>
    /// <b>System.ThumbnailCacheId</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY ThumbnailCacheId => new(new("446d16b1-8dad-4870-a748-402ea43d788c"), 100);
    
    /// <summary>
    /// <b>System.ThumbnailCacheIdParts</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
    /// </summary>
    public static PROPERTYKEY ThumbnailCacheIdParts => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 24);
    
    /// <summary>
    /// <b>System.ThumbnailStream</b> of <b>VT_STREAM</b> type.
    /// </summary>
    public static PROPERTYKEY ThumbnailStream => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 27);
    
    /// <summary>
    /// <b>System.Title</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY Title => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 2);
    
    /// <summary>
    /// <b>System.TitleSortOverride</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY TitleSortOverride => new(new("f0f7984d-222e-4ad2-82ab-1dd8ea40e57e"), 300);
    
    /// <summary>
    /// <b>System.TooltipThumbnailStream</b> of <b>VT_STREAM</b> type.
    /// </summary>
    public static PROPERTYKEY TooltipThumbnailStream => new(new("446d16b1-8dad-4870-a748-402ea43d788c"), 105);
    
    /// <summary>
    /// <b>System.TotalFileSize</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY TotalFileSize => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 14);
    
    /// <summary>
    /// <b>System.Trademarks</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY Trademarks => new(new("0cef7d53-fa64-11d1-a203-0000f81fedee"), 9);
    
    /// <summary>
    /// <b>System.TransferOrder</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY TransferOrder => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 106);
    
    /// <summary>
    /// <b>System.TransferPath</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY TransferPath => new(new("cebf9b37-26ae-466b-9fe9-c7550c4b0ce8"), 100);
    
    /// <summary>
    /// <b>System.TransferPosition</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY TransferPosition => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 104);
    
    /// <summary>
    /// <b>System.TransferSize</b> of <b>VT_UI8</b> type.
    /// </summary>
    public static PROPERTYKEY TransferSize => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 105);
    
    /// <summary>
    /// <b>System.UrlHostName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY UrlHostName => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 4);
    
    /// <summary>
    /// <b>System.UrlScheme</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY UrlScheme => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 3);
    
    /// <summary>
    /// <b>System.UserDisplayName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY UserDisplayName => new(new("54b3a473-59aa-445b-aecd-77541ba8b7c9"), 3);
    
    /// <summary>
    /// <b>System.UserName</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY UserName => new(new("54b3a473-59aa-445b-aecd-77541ba8b7c9"), 2);
    
    /// <summary>
    /// <b>System.UserProfilePath</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY UserProfilePath => new(new("54b3a473-59aa-445b-aecd-77541ba8b7c9"), 5);
    
    /// <summary>
    /// <b>System.VelocityFeatureId</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY VelocityFeatureId => new(new("5567bf77-2be2-4222-befa-d0c9c9cc4b6e"), 2);
    
    /// <summary>
    /// <b>System.VerbRestrictions</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY VerbRestrictions => new(new("705ccb0f-5a0d-41ea-b2ca-2c9b5cc7db41"), 100);
    
    /// <summary>
    /// <b>System.VolumeId</b> of <b>VT_CLSID</b> type.
    /// </summary>
    public static PROPERTYKEY VolumeId => new(new("446d16b1-8dad-4870-a748-402ea43d788c"), 104);
    
    /// <summary>
    /// <b>System.WebAccountID</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY WebAccountID => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 7);
    
    /// <summary>
    /// <b>System.WebDavPath</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY WebDavPath => new(new("1804d1fb-9fa4-441d-a536-76468ac43307"), 100);
    
    /// <summary>
    /// <b>System.WebPreviewUrl</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY WebPreviewUrl => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 23);
    
    /// <summary>
    /// <b>System.WhichFolder</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY WhichFolder => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 3);
    
    /// <summary>
    /// <b>System.WordWheel</b> of <b>VT_LPWSTR</b> type.
    /// </summary>
    public static PROPERTYKEY WordWheel => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 5);
    
    /// <summary>
    /// <b>System.ZoneIdentifier</b> of <b>VT_UI4</b> type.
    /// </summary>
    public static PROPERTYKEY ZoneIdentifier => new(new("502cfeab-47eb-459c-b960-e6d8728f7701"), 100);
    
    public static class Actions
    {
        /// <summary>
        /// <b>System.Actions.ActionName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ActionName => new(new("720eb626-dbe4-4113-835c-9315e1e2ff77"), 2);
        
        /// <summary>
        /// <b>System.Actions.ActivationContext</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ActivationContext => new(new("720eb626-dbe4-4113-835c-9315e1e2ff77"), 3);
    }
    
    public static class Activity
    {
        /// <summary>
        /// <b>System.Activity.AccountId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AccountId => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 626);
        
        /// <summary>
        /// <b>System.Activity.ActivationUri</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ActivationUri => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 640);
        
        /// <summary>
        /// <b>System.Activity.ActivityId</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY ActivityId => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 620);
        
        /// <summary>
        /// <b>System.Activity.AppDisplayName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AppDisplayName => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 600);
        
        /// <summary>
        /// <b>System.Activity.AppIdKind</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY AppIdKind => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 619);
        
        /// <summary>
        /// <b>System.Activity.AppIdList</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY AppIdList => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 635);
        
        /// <summary>
        /// <b>System.Activity.AppImageUri</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AppImageUri => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 601);
        
        /// <summary>
        /// <b>System.Activity.AttributionName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AttributionName => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 623);
        
        /// <summary>
        /// <b>System.Activity.BackgroundColor</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY BackgroundColor => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 602);
        
        /// <summary>
        /// <b>System.Activity.ContentImageUri</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContentImageUri => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 603);
        
        /// <summary>
        /// <b>System.Activity.ContentUri</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContentUri => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 604);
        
        /// <summary>
        /// <b>System.Activity.ContentVisualPropertiesHash</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY ContentVisualPropertiesHash => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 624);
        
        /// <summary>
        /// <b>System.Activity.Description</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Description => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 605);
        
        /// <summary>
        /// <b>System.Activity.DisplayText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DisplayText => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 606);
        
        /// <summary>
        /// <b>System.Activity.FallbackUri</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FallbackUri => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 641);
        
        /// <summary>
        /// <b>System.Activity.HasAdaptiveContent</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY HasAdaptiveContent => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 625);
        
        /// <summary>
        /// <b>System.Activity.SetCategory</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SetCategory => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 639);
        
        /// <summary>
        /// <b>System.Activity.SetId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SetId => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 634);
    }
    
    public static class ActivityHistory
    {
        /// <summary>
        /// <b>System.ActivityHistory.ActiveDays</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY ActiveDays => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 608);
        
        /// <summary>
        /// <b>System.ActivityHistory.ActiveDuration</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY ActiveDuration => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 609);
        
        /// <summary>
        /// <b>System.ActivityHistory.AppActivityId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AppActivityId => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 611);
        
        /// <summary>
        /// <b>System.ActivityHistory.AppId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AppId => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 612);
        
        /// <summary>
        /// <b>System.ActivityHistory.DaysActive</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DaysActive => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 627);
        
        /// <summary>
        /// <b>System.ActivityHistory.DeviceId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceId => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 614);
        
        /// <summary>
        /// <b>System.ActivityHistory.DeviceMake</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceMake => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 632);
        
        /// <summary>
        /// <b>System.ActivityHistory.DeviceModel</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceModel => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 633);
        
        /// <summary>
        /// <b>System.ActivityHistory.DeviceName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceName => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 613);
        
        /// <summary>
        /// <b>System.ActivityHistory.DeviceType</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceType => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 631);
        
        /// <summary>
        /// <b>System.ActivityHistory.EndTime</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY EndTime => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 616);
        
        /// <summary>
        /// <b>System.ActivityHistory.HoursActive</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY HoursActive => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 628);
        
        /// <summary>
        /// <b>System.ActivityHistory.Id</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY Id => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 617);
        
        /// <summary>
        /// <b>System.ActivityHistory.Importance</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY Importance => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 637);
        
        /// <summary>
        /// <b>System.ActivityHistory.IsHistoryAttributedToSetAnchor</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsHistoryAttributedToSetAnchor => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 638);
        
        /// <summary>
        /// <b>System.ActivityHistory.IsLocal</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsLocal => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 630);
        
        /// <summary>
        /// <b>System.ActivityHistory.LocalEndTime</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY LocalEndTime => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 622);
        
        /// <summary>
        /// <b>System.ActivityHistory.LocalStartTime</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY LocalStartTime => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 621);
        
        /// <summary>
        /// <b>System.ActivityHistory.LocationActivityId</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY LocationActivityId => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 629);
        
        /// <summary>
        /// <b>System.ActivityHistory.StartTime</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY StartTime => new(new("c5043536-932e-219e-5fb9-1c2807d7b03e"), 618);
    }
    
    public static class Address
    {
        /// <summary>
        /// <b>System.Address.Country</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Country => new(new("c07b4199-e1df-4493-b1e1-de5946fb58f8"), 100);
        
        /// <summary>
        /// <b>System.Address.CountryCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CountryCode => new(new("c07b4199-e1df-4493-b1e1-de5946fb58f8"), 101);
        
        /// <summary>
        /// <b>System.Address.Region</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Region => new(new("c07b4199-e1df-4493-b1e1-de5946fb58f8"), 102);
        
        /// <summary>
        /// <b>System.Address.RegionCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RegionCode => new(new("c07b4199-e1df-4493-b1e1-de5946fb58f8"), 103);
        
        /// <summary>
        /// <b>System.Address.Town</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Town => new(new("c07b4199-e1df-4493-b1e1-de5946fb58f8"), 104);
    }
    
    public static class AppContract
    {
        /// <summary>
        /// <b>System.AppContract.Category</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Category => new(new("f3c9b698-be85-47ce-888f-83874d9abcb4"), 6);
        
        /// <summary>
        /// <b>System.AppContract.Hidden</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Hidden => new(new("f3c9b698-be85-47ce-888f-83874d9abcb4"), 3);
        
        /// <summary>
        /// <b>System.AppContract.Pinned</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Pinned => new(new("f3c9b698-be85-47ce-888f-83874d9abcb4"), 2);
        
        /// <summary>
        /// <b>System.AppContract.PinnedOrder</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY PinnedOrder => new(new("f3c9b698-be85-47ce-888f-83874d9abcb4"), 4);
        
        /// <summary>
        /// <b>System.AppContract.Relevance</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY Relevance => new(new("f3c9b698-be85-47ce-888f-83874d9abcb4"), 5);
        
        /// <summary>
        /// <b>System.AppContract.SupportedFileTypes</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY SupportedFileTypes => new(new("f3c9b698-be85-47ce-888f-83874d9abcb4"), 7);
    }
    
    public static class AppUserModel
    {
        /// <summary>
        /// <b>System.AppUserModel.ActivationContext</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ActivationContext => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 20);
        
        /// <summary>
        /// <b>System.AppUserModel.BestShortcut</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY BestShortcut => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 10);
        
        /// <summary>
        /// <b>System.AppUserModel.DestListLogoUri</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DestListLogoUri => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 29);
        
        /// <summary>
        /// <b>System.AppUserModel.DestListProvidedDescription</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DestListProvidedDescription => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 28);
        
        /// <summary>
        /// <b>System.AppUserModel.DestListProvidedGroupName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DestListProvidedGroupName => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 30);
        
        /// <summary>
        /// <b>System.AppUserModel.DestListProvidedTitle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DestListProvidedTitle => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 27);
        
        /// <summary>
        /// <b>System.AppUserModel.ExcludeFromShowInNewInstall</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY ExcludeFromShowInNewInstall => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 8);
        
        /// <summary>
        /// <b>System.AppUserModel.ExcludedFromLauncher</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY ExcludedFromLauncher => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 23);
        
        /// <summary>
        /// <b>System.AppUserModel.FeatureOnDemand</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY FeatureOnDemand => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 35);
        
        /// <summary>
        /// <b>System.AppUserModel.HostEnvironment</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY HostEnvironment => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 14);
        
        /// <summary>
        /// <b>System.AppUserModel.ID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ID => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 5);
        
        /// <summary>
        /// <b>System.AppUserModel.InstalledBy</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY InstalledBy => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 18);
        
        /// <summary>
        /// <b>System.AppUserModel.IsDestListLink</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsDestListLink => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 7);
        
        /// <summary>
        /// <b>System.AppUserModel.IsDestListSeparator</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsDestListSeparator => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 6);
        
        /// <summary>
        /// <b>System.AppUserModel.IsDualMode</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsDualMode => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 11);
        
        /// <summary>
        /// <b>System.AppUserModel.IsSystemComponent</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsSystemComponent => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 40);
        
        /// <summary>
        /// <b>System.AppUserModel.IsUninstallable</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsUninstallable => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 39);
        
        /// <summary>
        /// <b>System.AppUserModel.PackageFamilyName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PackageFamilyName => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 17);
        
        /// <summary>
        /// <b>System.AppUserModel.PackageFullName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PackageFullName => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 21);
        
        /// <summary>
        /// <b>System.AppUserModel.PackageInstallPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PackageInstallPath => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 15);
        
        /// <summary>
        /// <b>System.AppUserModel.PackageRelativeApplicationID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PackageRelativeApplicationID => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 22);
        
        /// <summary>
        /// <b>System.AppUserModel.ParentID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ParentID => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 19);
        
        /// <summary>
        /// <b>System.AppUserModel.PreventPinning</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY PreventPinning => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 9);
        
        /// <summary>
        /// <b>System.AppUserModel.RecordState</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY RecordState => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 16);
        
        /// <summary>
        /// <b>System.AppUserModel.RelaunchCommand</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RelaunchCommand => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 2);
        
        /// <summary>
        /// <b>System.AppUserModel.RelaunchDisplayNameResource</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RelaunchDisplayNameResource => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 4);
        
        /// <summary>
        /// <b>System.AppUserModel.RelaunchIconResource</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RelaunchIconResource => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 3);
        
        /// <summary>
        /// <b>System.AppUserModel.Relevance</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY Relevance => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 13);
        
        /// <summary>
        /// <b>System.AppUserModel.RunFlags</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY RunFlags => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 25);
        
        /// <summary>
        /// <b>System.AppUserModel.SettingsCommand</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SettingsCommand => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 38);
        
        /// <summary>
        /// <b>System.AppUserModel.StartPinOption</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY StartPinOption => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 12);
        
        /// <summary>
        /// <b>System.AppUserModel.TileUniqueId</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY TileUniqueId => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 32);
        
        /// <summary>
        /// <b>System.AppUserModel.ToastActivatorCLSID</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY ToastActivatorCLSID => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 26);
        
        /// <summary>
        /// <b>System.AppUserModel.UninstallCommand</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UninstallCommand => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 37);
        
        /// <summary>
        /// <b>System.AppUserModel.VisualElementsManifestHintPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY VisualElementsManifestHintPath => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 31);
        
        public static class PinMigration
        {
            /// <summary>
            /// <b>System.AppUserModel.PinMigration.AppIds</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY AppIds => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 36);
            
            /// <summary>
            /// <b>System.AppUserModel.PinMigration.Executable</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY Executable => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 33);
            
            /// <summary>
            /// <b>System.AppUserModel.PinMigration.PackagedAppId</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY PackagedAppId => new(new("9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3"), 34);
        }
    }
    
    public static class Audio
    {
        /// <summary>
        /// <b>System.Audio.ChannelCount</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ChannelCount => new(new("64440490-4c8b-11d1-8b70-080036b11a03"), 7);
        
        /// <summary>
        /// <b>System.Audio.Compression</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Compression => new(new("64440490-4c8b-11d1-8b70-080036b11a03"), 10);
        
        /// <summary>
        /// <b>System.Audio.EncodingBitrate</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY EncodingBitrate => new(new("64440490-4c8b-11d1-8b70-080036b11a03"), 4);
        
        /// <summary>
        /// <b>System.Audio.Format</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Format => new(new("64440490-4c8b-11d1-8b70-080036b11a03"), 2);
        
        /// <summary>
        /// <b>System.Audio.IsVariableBitRate</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsVariableBitRate => new(new("e6822fee-8c17-4d62-823c-8e9cfcbd1d5c"), 100);
        
        /// <summary>
        /// <b>System.Audio.PeakValue</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY PeakValue => new(new("2579e5d0-1116-4084-bd9a-9b4f7cb4df5e"), 100);
        
        /// <summary>
        /// <b>System.Audio.SampleRate</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY SampleRate => new(new("64440490-4c8b-11d1-8b70-080036b11a03"), 5);
        
        /// <summary>
        /// <b>System.Audio.SampleSize</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY SampleSize => new(new("64440490-4c8b-11d1-8b70-080036b11a03"), 6);
        
        /// <summary>
        /// <b>System.Audio.StreamName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY StreamName => new(new("64440490-4c8b-11d1-8b70-080036b11a03"), 9);
        
        /// <summary>
        /// <b>System.Audio.StreamNumber</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY StreamNumber => new(new("64440490-4c8b-11d1-8b70-080036b11a03"), 8);
    }
    
    public static class Calendar
    {
        /// <summary>
        /// <b>System.Calendar.Duration</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Duration => new(new("293ca35a-09aa-4dd2-b180-1fe245728a52"), 100);
        
        /// <summary>
        /// <b>System.Calendar.IsOnline</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsOnline => new(new("bfee9149-e3e2-49a7-a862-c05988145cec"), 100);
        
        /// <summary>
        /// <b>System.Calendar.IsRecurring</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsRecurring => new(new("315b9c8d-80a9-4ef9-ae16-8e746da51d70"), 100);
        
        /// <summary>
        /// <b>System.Calendar.Location</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Location => new(new("f6272d18-cecc-40b1-b26a-3911717aa7bd"), 100);
        
        /// <summary>
        /// <b>System.Calendar.OptionalAttendeeAddresses</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY OptionalAttendeeAddresses => new(new("d55bae5a-3892-417a-a649-c6ac5aaaeab3"), 100);
        
        /// <summary>
        /// <b>System.Calendar.OptionalAttendeeNames</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY OptionalAttendeeNames => new(new("09429607-582d-437f-84c3-de93a2b24c3c"), 100);
        
        /// <summary>
        /// <b>System.Calendar.OrganizerAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OrganizerAddress => new(new("744c8242-4df5-456c-ab9e-014efb9021e3"), 100);
        
        /// <summary>
        /// <b>System.Calendar.OrganizerName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OrganizerName => new(new("aaa660f9-9865-458e-b484-01bc7fe3973e"), 100);
        
        /// <summary>
        /// <b>System.Calendar.ReminderTime</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY ReminderTime => new(new("72fc5ba4-24f9-4011-9f3f-add27afad818"), 100);
        
        /// <summary>
        /// <b>System.Calendar.RequiredAttendeeAddresses</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY RequiredAttendeeAddresses => new(new("0ba7d6c3-568d-4159-ab91-781a91fb71e5"), 100);
        
        /// <summary>
        /// <b>System.Calendar.RequiredAttendeeNames</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY RequiredAttendeeNames => new(new("b33af30b-f552-4584-936c-cb93e5cda29f"), 100);
        
        /// <summary>
        /// <b>System.Calendar.Resources</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Resources => new(new("00f58a38-c54b-4c40-8696-97235980eae1"), 100);
        
        /// <summary>
        /// <b>System.Calendar.ResponseStatus</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY ResponseStatus => new(new("188c1f91-3c40-4132-9ec5-d8b03b72a8a2"), 100);
        
        /// <summary>
        /// <b>System.Calendar.ShowTimeAs</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY ShowTimeAs => new(new("5bf396d4-5eb2-466f-bde9-2fb3f2361d6e"), 100);
        
        /// <summary>
        /// <b>System.Calendar.ShowTimeAsText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ShowTimeAsText => new(new("53da57cf-62c0-45c4-81de-7610bcefd7f5"), 100);
    }
    
    public static class Communication
    {
        /// <summary>
        /// <b>System.Communication.AccountName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AccountName => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 9);
        
        /// <summary>
        /// <b>System.Communication.DateItemExpires</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateItemExpires => new(new("428040ac-a177-4c8a-9760-f6f761227f9a"), 100);
        
        /// <summary>
        /// <b>System.Communication.Direction</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY Direction => new(new("8e531030-b960-4346-ae0d-66bc9a86fb94"), 100);
        
        /// <summary>
        /// <b>System.Communication.DirectoryServer</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DirectoryServer => new(new("12ea418f-d8cd-4cdf-9b23-457eaac7ff0d"), 100);
        
        /// <summary>
        /// <b>System.Communication.FollowupIconIndex</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY FollowupIconIndex => new(new("83a6347e-6fe4-4f40-ba9c-c4865240d1f4"), 100);
        
        /// <summary>
        /// <b>System.Communication.HeaderItem</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY HeaderItem => new(new("c9c34f84-2241-4401-b607-bd20ed75ae7f"), 100);
        
        /// <summary>
        /// <b>System.Communication.NewsgroupName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY NewsgroupName => new(new("9c1fcf74-2d97-41ba-b4ae-cb2e3661a6e4"), 7);
        
        /// <summary>
        /// <b>System.Communication.PolicyTag</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PolicyTag => new(new("ec0b4191-ab0b-4c66-90b6-c6637cdebbab"), 100);
        
        /// <summary>
        /// <b>System.Communication.SecurityFlags</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY SecurityFlags => new(new("8619a4b6-9f4d-4429-8c0f-b996ca59e335"), 100);
        
        /// <summary>
        /// <b>System.Communication.Suffix</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Suffix => new(new("807b653a-9e91-43ef-8f97-11ce04ee20c5"), 100);
        
        /// <summary>
        /// <b>System.Communication.TaskStatus</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY TaskStatus => new(new("be1a72c6-9a1d-46b7-afe7-afaf8cef4999"), 100);
        
        /// <summary>
        /// <b>System.Communication.TaskStatusText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TaskStatusText => new(new("a6744477-c237-475b-a075-54f34498292a"), 100);
    }
    
    public static class Computer
    {
        /// <summary>
        /// <b>System.Computer.DecoratedFreeSpace</b> of <b>VT_UI8, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DecoratedFreeSpace => new(new("9b174b35-40ff-11d2-a27e-00c04fc30871"), 7);
        
        /// <summary>
        /// <b>System.Computer.Description</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Description => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 27);
        
        /// <summary>
        /// <b>System.Computer.DomainName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DomainName => new(new("5cde9f0e-1de4-4453-96a9-56e8832efa3d"), 1);
        
        /// <summary>
        /// <b>System.Computer.Memory</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Memory => new(new("fd9d9fc7-38ec-436d-8fc6-ec39bad301e6"), 101);
        
        /// <summary>
        /// <b>System.Computer.Processor</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Processor => new(new("fd9d9fc7-38ec-436d-8fc6-ec39bad301e6"), 100);
        
        /// <summary>
        /// <b>System.Computer.SimpleName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SimpleName => new(new("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 10);
        
        /// <summary>
        /// <b>System.Computer.Workgroup</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Workgroup => new(new("5cde9f0e-1de4-4453-96a9-56e8832efa3d"), 2);
    }
    
    public static class ConnectedSearch
    {
        /// <summary>
        /// <b>System.ConnectedSearch.ActivationCommand</b> of <b>VT_UNKNOWN</b> type.
        /// </summary>
        public static PROPERTYKEY ActivationCommand => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 18);
        
        /// <summary>
        /// <b>System.ConnectedSearch.AddOpenInBrowserCommand</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY AddOpenInBrowserCommand => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 29);
        
        /// <summary>
        /// <b>System.ConnectedSearch.AppInstalledState</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY AppInstalledState => new(new("d76e7ba8-dfa6-48e7-9670-d62dfb07206b"), 4);
        
        /// <summary>
        /// <b>System.ConnectedSearch.AppMinVersion</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AppMinVersion => new(new("d76e7ba8-dfa6-48e7-9670-d62dfb07206b"), 3);
        
        /// <summary>
        /// <b>System.ConnectedSearch.ApplicationSearchScope</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ApplicationSearchScope => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 9);
        
        /// <summary>
        /// <b>System.ConnectedSearch.AutoComplete</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY AutoComplete => new(new("916d17ac-8a97-48af-85b7-867a88fad542"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.BypassViewAction</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY BypassViewAction => new(new("dce33a78-aa18-4b3d-b1df-a6621ac8bdd2"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.ChildCount</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ChildCount => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 11);
        
        /// <summary>
        /// <b>System.ConnectedSearch.ContractId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContractId => new(new("d76e7ba8-dfa6-48e7-9670-d62dfb07206b"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.CopyText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CopyText => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 28);
        
        /// <summary>
        /// <b>System.ConnectedSearch.DeferImagePrefetch</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY DeferImagePrefetch => new(new("a40294ef-d2b1-40ed-9512-dd3853b431f5"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.DisambiguationId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DisambiguationId => new(new("f27abe3a-7111-4dda-8cb2-29222ae23566"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.DisambiguationText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DisambiguationText => new(new("05e932b1-7ca2-491f-bd69-99b4cb266cbb"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.FallbackTemplate</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FallbackTemplate => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 4);
        
        /// <summary>
        /// <b>System.ConnectedSearch.HistoryDescription</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HistoryDescription => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 22);
        
        /// <summary>
        /// <b>System.ConnectedSearch.HistoryGlyph</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HistoryGlyph => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 23);
        
        /// <summary>
        /// <b>System.ConnectedSearch.HistoryTitle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HistoryTitle => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 21);
        
        /// <summary>
        /// <b>System.ConnectedSearch.ImagePrefetchStage</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ImagePrefetchStage => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 31);
        
        /// <summary>
        /// <b>System.ConnectedSearch.ImageUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ImageUrl => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 30);
        
        /// <summary>
        /// <b>System.ConnectedSearch.ImpressionId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ImpressionId => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 6);
        
        /// <summary>
        /// <b>System.ConnectedSearch.IsActivatable</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsActivatable => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 14);
        
        /// <summary>
        /// <b>System.ConnectedSearch.IsAppAvailable</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsAppAvailable => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 20);
        
        /// <summary>
        /// <b>System.ConnectedSearch.IsHistoryItem</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsHistoryItem => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 19);
        
        /// <summary>
        /// <b>System.ConnectedSearch.IsLocalItem</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsLocalItem => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 32);
        
        /// <summary>
        /// <b>System.ConnectedSearch.IsVisibilityTracked</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsVisibilityTracked => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 7);
        
        /// <summary>
        /// <b>System.ConnectedSearch.IsVisibleByDefault</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsVisibleByDefault => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 13);
        
        /// <summary>
        /// <b>System.ConnectedSearch.ItemSource</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ItemSource => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 17);
        
        /// <summary>
        /// <b>System.ConnectedSearch.JumpList</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JumpList => new(new("e2d40928-632c-4280-a202-e0c2ad1ea0f4"), 3);
        
        /// <summary>
        /// <b>System.ConnectedSearch.LinkText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LinkText => new(new("12fa14f5-c6fe-4545-bce2-1ed6cb6b8422"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.LocalWeights</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LocalWeights => new(new("79486778-4c6f-4dde-bc53-cd594311af99"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.ParentId</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY ParentId => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 10);
        
        /// <summary>
        /// <b>System.ConnectedSearch.QsCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY QsCode => new(new("e2d40928-632c-4280-a202-e0c2ad1ea0f4"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.ReferrerId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ReferrerId => new(new("a8a7a412-1927-4a34-b1d4-45f67cc672fb"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.RegionId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RegionId => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 16);
        
        /// <summary>
        /// <b>System.ConnectedSearch.RenderingTemplate</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RenderingTemplate => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 3);
        
        /// <summary>
        /// <b>System.ConnectedSearch.RequireInstall</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY RequireInstall => new(new("cc158e89-6581-4311-9637-a8da9002f118"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.RequireTemplate</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY RequireTemplate => new(new("73389854-0b42-4ea6-bc67-847d430899fd"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.RequiresConsent</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY RequiresConsent => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 27);
        
        /// <summary>
        /// <b>System.ConnectedSearch.SuggestionContext</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SuggestionContext => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 15);
        
        /// <summary>
        /// <b>System.ConnectedSearch.SuggestionDetailText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SuggestionDetailText => new(new("e9641eff-af25-4db7-947b-4128929f8ef5"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.SuppressLocalHero</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY SuppressLocalHero => new(new("b769d0fe-bc33-421a-8ce6-45add82ec756"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.TelemetryData</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TelemetryData => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 8);
        
        /// <summary>
        /// <b>System.ConnectedSearch.TelemetryId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TelemetryId => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 5);
        
        /// <summary>
        /// <b>System.ConnectedSearch.TopLevelId</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY TopLevelId => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 12);
        
        /// <summary>
        /// <b>System.ConnectedSearch.Type</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Type => new(new("e1ad4953-a752-443c-93bf-80c7525566c2"), 2);
        
        /// <summary>
        /// <b>System.ConnectedSearch.VoiceCommandExamples</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY VoiceCommandExamples => new(new("e2d40928-632c-4280-a202-e0c2ad1ea0f4"), 4);
    }
    
    public static class Contact
    {
        /// <summary>
        /// <b>System.Contact.AccountPictureDynamicVideo</b> of <b>VT_STREAM</b> type.
        /// </summary>
        public static PROPERTYKEY AccountPictureDynamicVideo => new(new("0b8bb018-2725-4b44-92ba-7933aeb2dde7"), 2);
        
        /// <summary>
        /// <b>System.Contact.AccountPictureLarge</b> of <b>VT_STREAM</b> type.
        /// </summary>
        public static PROPERTYKEY AccountPictureLarge => new(new("0b8bb018-2725-4b44-92ba-7933aeb2dde7"), 3);
        
        /// <summary>
        /// <b>System.Contact.AccountPictureSmall</b> of <b>VT_STREAM</b> type.
        /// </summary>
        public static PROPERTYKEY AccountPictureSmall => new(new("0b8bb018-2725-4b44-92ba-7933aeb2dde7"), 4);
        
        /// <summary>
        /// <b>System.Contact.Anniversary</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY Anniversary => new(new("9ad5badb-cea7-4470-a03d-b84e51b9949e"), 100);
        
        /// <summary>
        /// <b>System.Contact.AssistantName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AssistantName => new(new("cd102c9c-5540-4a88-a6f6-64e4981c8cd1"), 100);
        
        /// <summary>
        /// <b>System.Contact.AssistantTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AssistantTelephone => new(new("9a93244d-a7ad-4ff8-9b99-45ee4cc09af6"), 100);
        
        /// <summary>
        /// <b>System.Contact.Birthday</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY Birthday => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 47);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress => new(new("730fb6dd-cf7c-426b-a03f-bd166cc9ee24"), 100);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress1Country</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress1Country => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 119);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress1Locality</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress1Locality => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 117);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress1PostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress1PostalCode => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 120);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress1Region</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress1Region => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 118);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress1Street</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress1Street => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 116);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress2Country</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress2Country => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 124);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress2Locality</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress2Locality => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 122);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress2PostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress2PostalCode => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 125);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress2Region</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress2Region => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 123);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress2Street</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress2Street => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 121);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress3Country</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress3Country => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 129);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress3Locality</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress3Locality => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 127);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress3PostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress3PostalCode => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 130);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress3Region</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress3Region => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 128);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddress3Street</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddress3Street => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 126);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddressCity</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddressCity => new(new("402b5934-ec5a-48c3-93e6-85e86a2d934e"), 100);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddressCountry</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddressCountry => new(new("b0b87314-fcf6-4feb-8dff-a50da6af561c"), 100);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddressPostOfficeBox</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddressPostOfficeBox => new(new("bc4e71ce-17f9-48d5-bee9-021df0ea5409"), 100);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddressPostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddressPostalCode => new(new("e1d4a09e-d758-4cd1-b6ec-34a8b5a73f80"), 100);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddressState</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddressState => new(new("446f787f-10c4-41cb-a6c4-4d0343551597"), 100);
        
        /// <summary>
        /// <b>System.Contact.BusinessAddressStreet</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessAddressStreet => new(new("ddd1460f-c0bf-4553-8ce4-10433c908fb0"), 100);
        
        /// <summary>
        /// <b>System.Contact.BusinessEmailAddresses</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessEmailAddresses => new(new("f271c659-7e5e-471f-ba25-7f77b286f836"), 100);
        
        /// <summary>
        /// <b>System.Contact.BusinessFaxNumber</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessFaxNumber => new(new("91eff6f3-2e27-42ca-933e-7c999fbe310b"), 100);
        
        /// <summary>
        /// <b>System.Contact.BusinessHomePage</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessHomePage => new(new("56310920-2491-4919-99ce-eadb06fafdb2"), 100);
        
        /// <summary>
        /// <b>System.Contact.BusinessTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BusinessTelephone => new(new("6a15e5a0-0a1e-4cd7-bb8c-d2f1b0c929bc"), 100);
        
        /// <summary>
        /// <b>System.Contact.CallbackTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CallbackTelephone => new(new("bf53d1c3-49e0-4f7f-8567-5a821d8ac542"), 100);
        
        /// <summary>
        /// <b>System.Contact.CarTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CarTelephone => new(new("8fdc6dea-b929-412b-ba90-397a257465fe"), 100);
        
        /// <summary>
        /// <b>System.Contact.Children</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Children => new(new("d4729704-8ef1-43ef-9024-2bd381187fd5"), 100);
        
        /// <summary>
        /// <b>System.Contact.CompanyMainTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CompanyMainTelephone => new(new("8589e481-6040-473d-b171-7fa89c2708ed"), 100);
        
        /// <summary>
        /// <b>System.Contact.ConnectedServiceDisplayName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ConnectedServiceDisplayName => new(new("39b77f4f-a104-4863-b395-2db2ad8f7bc1"), 100);
        
        /// <summary>
        /// <b>System.Contact.ConnectedServiceIdentities</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY ConnectedServiceIdentities => new(new("80f41eb8-afc4-4208-aa5f-cce21a627281"), 100);
        
        /// <summary>
        /// <b>System.Contact.ConnectedServiceName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ConnectedServiceName => new(new("b5c84c9e-5927-46b5-a3cc-933c21b78469"), 100);
        
        /// <summary>
        /// <b>System.Contact.ConnectedServiceSupportedActions</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ConnectedServiceSupportedActions => new(new("a19fb7a9-024b-4371-a8bf-4d29c3e4e9c9"), 100);
        
        /// <summary>
        /// <b>System.Contact.DataSuppliers</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DataSuppliers => new(new("9660c283-fc3a-4a08-a096-eed3aac46da2"), 100);
        
        /// <summary>
        /// <b>System.Contact.Department</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Department => new(new("fc9f7306-ff8f-4d49-9fb6-3ffe5c0951ec"), 100);
        
        /// <summary>
        /// <b>System.Contact.DisplayBusinessPhoneNumbers</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DisplayBusinessPhoneNumbers => new(new("364028da-d895-41fe-a584-302b1bb70a76"), 100);
        
        /// <summary>
        /// <b>System.Contact.DisplayHomePhoneNumbers</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DisplayHomePhoneNumbers => new(new("5068bcdf-d697-4d85-8c53-1f1cdab01763"), 100);
        
        /// <summary>
        /// <b>System.Contact.DisplayMobilePhoneNumbers</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DisplayMobilePhoneNumbers => new(new("9cb0c358-9d7a-46b1-b466-dcc6f1a3d93d"), 100);
        
        /// <summary>
        /// <b>System.Contact.DisplayOtherPhoneNumbers</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DisplayOtherPhoneNumbers => new(new("03089873-8ee8-4191-bd60-d31f72b7900b"), 100);
        
        /// <summary>
        /// <b>System.Contact.EmailAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EmailAddress => new(new("f8fa7fa3-d12b-4785-8a4e-691a94f7a3e7"), 100);
        
        /// <summary>
        /// <b>System.Contact.EmailAddress2</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EmailAddress2 => new(new("38965063-edc8-4268-8491-b7723172cf29"), 100);
        
        /// <summary>
        /// <b>System.Contact.EmailAddress3</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EmailAddress3 => new(new("644d37b4-e1b3-4bad-b099-7e7c04966aca"), 100);
        
        /// <summary>
        /// <b>System.Contact.EmailAddresses</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY EmailAddresses => new(new("84d8f337-981d-44b3-9615-c7596dba17e3"), 100);
        
        /// <summary>
        /// <b>System.Contact.EmailName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EmailName => new(new("cc6f4f24-6083-4bd4-8754-674d0de87ab8"), 100);
        
        /// <summary>
        /// <b>System.Contact.FileAsName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FileAsName => new(new("f1a24aa7-9ca7-40f6-89ec-97def9ffe8db"), 100);
        
        /// <summary>
        /// <b>System.Contact.FirstName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FirstName => new(new("14977844-6b49-4aad-a714-a4513bf60460"), 100);
        
        /// <summary>
        /// <b>System.Contact.FullName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FullName => new(new("635e9051-50a5-4ba2-b9db-4ed056c77296"), 100);
        
        /// <summary>
        /// <b>System.Contact.Gender</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Gender => new(new("3c8cee58-d4f0-4cf9-b756-4e5d24447bcd"), 100);
        
        /// <summary>
        /// <b>System.Contact.GenderValue</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY GenderValue => new(new("3c8cee58-d4f0-4cf9-b756-4e5d24447bcd"), 101);
        
        /// <summary>
        /// <b>System.Contact.GivenName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY GivenName => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 70);
        
        /// <summary>
        /// <b>System.Contact.Hobbies</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Hobbies => new(new("5dc2253f-5e11-4adf-9cfe-910dd01e3e70"), 100);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress => new(new("98f98354-617a-46b8-8560-5b1b64bf1f89"), 100);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress1Country</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress1Country => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 104);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress1Locality</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress1Locality => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 102);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress1PostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress1PostalCode => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 105);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress1Region</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress1Region => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 103);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress1Street</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress1Street => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 101);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress2Country</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress2Country => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 109);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress2Locality</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress2Locality => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 107);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress2PostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress2PostalCode => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 110);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress2Region</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress2Region => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 108);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress2Street</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress2Street => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 106);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress3Country</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress3Country => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 114);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress3Locality</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress3Locality => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 112);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress3PostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress3PostalCode => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 115);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress3Region</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress3Region => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 113);
        
        /// <summary>
        /// <b>System.Contact.HomeAddress3Street</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddress3Street => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 111);
        
        /// <summary>
        /// <b>System.Contact.HomeAddressCity</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddressCity => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 65);
        
        /// <summary>
        /// <b>System.Contact.HomeAddressCountry</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddressCountry => new(new("08a65aa1-f4c9-43dd-9ddf-a33d8e7ead85"), 100);
        
        /// <summary>
        /// <b>System.Contact.HomeAddressPostOfficeBox</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddressPostOfficeBox => new(new("7b9f6399-0a3f-4b12-89bd-4adc51c918af"), 100);
        
        /// <summary>
        /// <b>System.Contact.HomeAddressPostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddressPostalCode => new(new("8afcc170-8a46-4b53-9eee-90bae7151e62"), 100);
        
        /// <summary>
        /// <b>System.Contact.HomeAddressState</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddressState => new(new("c89a23d0-7d6d-4eb8-87d4-776a82d493e5"), 100);
        
        /// <summary>
        /// <b>System.Contact.HomeAddressStreet</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeAddressStreet => new(new("0adef160-db3f-4308-9a21-06237b16fa2a"), 100);
        
        /// <summary>
        /// <b>System.Contact.HomeEmailAddresses</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeEmailAddresses => new(new("56c90e9d-9d46-4963-886f-2e1cd9a694ef"), 100);
        
        /// <summary>
        /// <b>System.Contact.HomeFaxNumber</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeFaxNumber => new(new("660e04d6-81ab-4977-a09f-82313113ab26"), 100);
        
        /// <summary>
        /// <b>System.Contact.HomeTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeTelephone => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 20);
        
        /// <summary>
        /// <b>System.Contact.IMAddress</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY IMAddress => new(new("d68dbd8a-3374-4b81-9972-3ec30682db3d"), 100);
        
        /// <summary>
        /// <b>System.Contact.Initials</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Initials => new(new("f3d8f40d-50cb-44a2-9718-40cb9119495d"), 100);
        
        /// <summary>
        /// <b>System.Contact.JobInfo1CompanyAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo1CompanyAddress => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 120);
        
        /// <summary>
        /// <b>System.Contact.JobInfo1CompanyName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo1CompanyName => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 102);
        
        /// <summary>
        /// <b>System.Contact.JobInfo1Department</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo1Department => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 106);
        
        /// <summary>
        /// <b>System.Contact.JobInfo1Manager</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo1Manager => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 105);
        
        /// <summary>
        /// <b>System.Contact.JobInfo1OfficeLocation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo1OfficeLocation => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 104);
        
        /// <summary>
        /// <b>System.Contact.JobInfo1Title</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo1Title => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 103);
        
        /// <summary>
        /// <b>System.Contact.JobInfo1YomiCompanyName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo1YomiCompanyName => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 101);
        
        /// <summary>
        /// <b>System.Contact.JobInfo2CompanyAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo2CompanyAddress => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 121);
        
        /// <summary>
        /// <b>System.Contact.JobInfo2CompanyName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo2CompanyName => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 108);
        
        /// <summary>
        /// <b>System.Contact.JobInfo2Department</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo2Department => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 113);
        
        /// <summary>
        /// <b>System.Contact.JobInfo2Manager</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo2Manager => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 112);
        
        /// <summary>
        /// <b>System.Contact.JobInfo2OfficeLocation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo2OfficeLocation => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 110);
        
        /// <summary>
        /// <b>System.Contact.JobInfo2Title</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo2Title => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 109);
        
        /// <summary>
        /// <b>System.Contact.JobInfo2YomiCompanyName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo2YomiCompanyName => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 107);
        
        /// <summary>
        /// <b>System.Contact.JobInfo3CompanyAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo3CompanyAddress => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 123);
        
        /// <summary>
        /// <b>System.Contact.JobInfo3CompanyName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo3CompanyName => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 115);
        
        /// <summary>
        /// <b>System.Contact.JobInfo3Department</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo3Department => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 119);
        
        /// <summary>
        /// <b>System.Contact.JobInfo3Manager</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo3Manager => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 118);
        
        /// <summary>
        /// <b>System.Contact.JobInfo3OfficeLocation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo3OfficeLocation => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 117);
        
        /// <summary>
        /// <b>System.Contact.JobInfo3Title</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo3Title => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 116);
        
        /// <summary>
        /// <b>System.Contact.JobInfo3YomiCompanyName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobInfo3YomiCompanyName => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 114);
        
        /// <summary>
        /// <b>System.Contact.JobTitle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY JobTitle => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 6);
        
        /// <summary>
        /// <b>System.Contact.Label</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Label => new(new("97b0ad89-df49-49cc-834e-660974fd755b"), 100);
        
        /// <summary>
        /// <b>System.Contact.LastName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LastName => new(new("8f367200-c270-457c-b1d4-e07c5bcd90c7"), 100);
        
        /// <summary>
        /// <b>System.Contact.MailingAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MailingAddress => new(new("c0ac206a-827e-4650-95ae-77e2bb74fcc9"), 100);
        
        /// <summary>
        /// <b>System.Contact.MiddleName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MiddleName => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 71);
        
        /// <summary>
        /// <b>System.Contact.MobileTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MobileTelephone => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 35);
        
        /// <summary>
        /// <b>System.Contact.NickName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY NickName => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 74);
        
        /// <summary>
        /// <b>System.Contact.OfficeLocation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OfficeLocation => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 7);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress => new(new("508161fa-313b-43d5-83a1-c1accf68622c"), 100);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress1Country</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress1Country => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 134);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress1Locality</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress1Locality => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 132);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress1PostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress1PostalCode => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 135);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress1Region</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress1Region => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 133);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress1Street</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress1Street => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 131);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress2Country</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress2Country => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 139);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress2Locality</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress2Locality => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 137);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress2PostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress2PostalCode => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 140);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress2Region</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress2Region => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 138);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress2Street</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress2Street => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 136);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress3Country</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress3Country => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 144);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress3Locality</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress3Locality => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 142);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress3PostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress3PostalCode => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 145);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress3Region</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress3Region => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 143);
        
        /// <summary>
        /// <b>System.Contact.OtherAddress3Street</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddress3Street => new(new("a7b6f596-d678-4bc1-b05f-0203d27e8aa1"), 141);
        
        /// <summary>
        /// <b>System.Contact.OtherAddressCity</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddressCity => new(new("6e682923-7f7b-4f0c-a337-cfca296687bf"), 100);
        
        /// <summary>
        /// <b>System.Contact.OtherAddressCountry</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddressCountry => new(new("8f167568-0aae-4322-8ed9-6055b7b0e398"), 100);
        
        /// <summary>
        /// <b>System.Contact.OtherAddressPostOfficeBox</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddressPostOfficeBox => new(new("8b26ea41-058f-43f6-aecc-4035681ce977"), 100);
        
        /// <summary>
        /// <b>System.Contact.OtherAddressPostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddressPostalCode => new(new("95c656c1-2abf-4148-9ed3-9ec602e3b7cd"), 100);
        
        /// <summary>
        /// <b>System.Contact.OtherAddressState</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddressState => new(new("71b377d6-e570-425f-a170-809fae73e54e"), 100);
        
        /// <summary>
        /// <b>System.Contact.OtherAddressStreet</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherAddressStreet => new(new("ff962609-b7d6-4999-862d-95180d529aea"), 100);
        
        /// <summary>
        /// <b>System.Contact.OtherEmailAddresses</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY OtherEmailAddresses => new(new("11d6336b-38c4-4ec9-84d6-eb38d0b150af"), 100);
        
        /// <summary>
        /// <b>System.Contact.PagerTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PagerTelephone => new(new("d6304e01-f8f5-4f45-8b15-d024a6296789"), 100);
        
        /// <summary>
        /// <b>System.Contact.PersonalTitle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PersonalTitle => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 69);
        
        /// <summary>
        /// <b>System.Contact.PhoneNumbersCanonical</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY PhoneNumbersCanonical => new(new("d042d2a1-927e-40b5-a503-6edbd42a517e"), 100);
        
        /// <summary>
        /// <b>System.Contact.Prefix</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Prefix => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 75);
        
        /// <summary>
        /// <b>System.Contact.PrimaryAddressCity</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryAddressCity => new(new("c8ea94f0-a9e3-4969-a94b-9c62a95324e0"), 100);
        
        /// <summary>
        /// <b>System.Contact.PrimaryAddressCountry</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryAddressCountry => new(new("e53d799d-0f3f-466e-b2ff-74634a3cb7a4"), 100);
        
        /// <summary>
        /// <b>System.Contact.PrimaryAddressPostOfficeBox</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryAddressPostOfficeBox => new(new("de5ef3c7-46e1-484e-9999-62c5308394c1"), 100);
        
        /// <summary>
        /// <b>System.Contact.PrimaryAddressPostalCode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryAddressPostalCode => new(new("18bbd425-ecfd-46ef-b612-7b4a6034eda0"), 100);
        
        /// <summary>
        /// <b>System.Contact.PrimaryAddressState</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryAddressState => new(new("f1176dfe-7138-4640-8b4c-ae375dc70a6d"), 100);
        
        /// <summary>
        /// <b>System.Contact.PrimaryAddressStreet</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryAddressStreet => new(new("63c25b20-96be-488f-8788-c09c407ad812"), 100);
        
        /// <summary>
        /// <b>System.Contact.PrimaryEmailAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryEmailAddress => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 48);
        
        /// <summary>
        /// <b>System.Contact.PrimaryTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryTelephone => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 25);
        
        /// <summary>
        /// <b>System.Contact.Profession</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Profession => new(new("7268af55-1ce4-4f6e-a41f-b6e4ef10e4a9"), 100);
        
        /// <summary>
        /// <b>System.Contact.SpouseName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SpouseName => new(new("9d2408b6-3167-422b-82b0-f583b7a7cfe3"), 100);
        
        /// <summary>
        /// <b>System.Contact.Suffix</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Suffix => new(new("176dc63c-2688-4e89-8143-a347800f25e9"), 73);
        
        /// <summary>
        /// <b>System.Contact.TTYTDDTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TTYTDDTelephone => new(new("aaf16bac-2b55-45e6-9f6d-415eb94910df"), 100);
        
        /// <summary>
        /// <b>System.Contact.TelexNumber</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TelexNumber => new(new("c554493c-c1f7-40c1-a76c-ef8c0614003e"), 100);
        
        /// <summary>
        /// <b>System.Contact.WebPage</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY WebPage => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 18);
        
        /// <summary>
        /// <b>System.Contact.Webpage2</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Webpage2 => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 124);
        
        /// <summary>
        /// <b>System.Contact.Webpage3</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Webpage3 => new(new("00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03"), 125);
        
        public static class JA
        {
            /// <summary>
            /// <b>System.Contact.JA.CompanyNamePhonetic</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY CompanyNamePhonetic => new(new("897b3694-fe9e-43e6-8066-260f590c0100"), 2);
            
            /// <summary>
            /// <b>System.Contact.JA.FirstNamePhonetic</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY FirstNamePhonetic => new(new("897b3694-fe9e-43e6-8066-260f590c0100"), 3);
            
            /// <summary>
            /// <b>System.Contact.JA.LastNamePhonetic</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY LastNamePhonetic => new(new("897b3694-fe9e-43e6-8066-260f590c0100"), 4);
        }
    }
    
    public static class ControlPanel
    {
        /// <summary>
        /// <b>System.ControlPanel.Category</b> of <b>VT_I4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Category => new(new("305ca226-d286-468e-b848-2b2e8e697b74"), 2);
        
        /// <summary>
        /// <b>System.ControlPanel.EnableInSafeMode</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY EnableInSafeMode => new(new("305ca226-d286-468e-b848-2b2e8e697b74"), 3);
        
        /// <summary>
        /// <b>System.ControlPanel.Module</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Module => new(new("305ca226-d286-468e-b848-2b2e8e697b74"), 4);
    }
    
    public static class DRM
    {
        /// <summary>
        /// <b>System.DRM.DatePlayExpires</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DatePlayExpires => new(new("aeac19e4-89ae-4508-b9b7-bb867abee2ed"), 6);
        
        /// <summary>
        /// <b>System.DRM.DatePlayStarts</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DatePlayStarts => new(new("aeac19e4-89ae-4508-b9b7-bb867abee2ed"), 5);
        
        /// <summary>
        /// <b>System.DRM.Description</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Description => new(new("aeac19e4-89ae-4508-b9b7-bb867abee2ed"), 3);
        
        /// <summary>
        /// <b>System.DRM.IsDisabled</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsDisabled => new(new("aeac19e4-89ae-4508-b9b7-bb867abee2ed"), 7);
        
        /// <summary>
        /// <b>System.DRM.IsProtected</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsProtected => new(new("aeac19e4-89ae-4508-b9b7-bb867abee2ed"), 2);
        
        /// <summary>
        /// <b>System.DRM.PlayCount</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY PlayCount => new(new("aeac19e4-89ae-4508-b9b7-bb867abee2ed"), 4);
    }
    
    public static class Device
    {
        /// <summary>
        /// <b>System.Device.PrinterURL</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrinterURL => new(new("0b48f35a-be6e-4f17-b108-3c4073d1669a"), 15);
    }
    
    public static class DeviceInterface
    {
        /// <summary>
        /// <b>System.DeviceInterface.PrinterDriverDirectory</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrinterDriverDirectory => new(new("847c66de-b8d6-4af9-abc3-6f4f926bc039"), 14);
        
        /// <summary>
        /// <b>System.DeviceInterface.PrinterDriverName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrinterDriverName => new(new("afc47170-14f5-498c-8f30-b0d19be449c6"), 11);
        
        /// <summary>
        /// <b>System.DeviceInterface.PrinterEnumerationFlag</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY PrinterEnumerationFlag => new(new("a00742a1-cd8c-4b37-95ab-70755587767a"), 3);
        
        /// <summary>
        /// <b>System.DeviceInterface.PrinterName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrinterName => new(new("0a7b84ef-0c27-463f-84ef-06c5070001be"), 10);
        
        /// <summary>
        /// <b>System.DeviceInterface.PrinterPortName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrinterPortName => new(new("eec7b761-6f94-41b1-949f-c729720dd13c"), 12);
        
        public static class Bluetooth
        {
            /// <summary>
            /// <b>System.DeviceInterface.Bluetooth.DeviceAddress</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY DeviceAddress => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 1);
            
            /// <summary>
            /// <b>System.DeviceInterface.Bluetooth.Flags</b> of <b>VT_UI4</b> type.
            /// </summary>
            public static PROPERTYKEY Flags => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 3);
            
            /// <summary>
            /// <b>System.DeviceInterface.Bluetooth.LastConnectedTime</b> of <b>VT_FILETIME</b> type.
            /// </summary>
            public static PROPERTYKEY LastConnectedTime => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 11);
            
            /// <summary>
            /// <b>System.DeviceInterface.Bluetooth.Manufacturer</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Manufacturer => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 4);
            
            /// <summary>
            /// <b>System.DeviceInterface.Bluetooth.ModelNumber</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY ModelNumber => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 5);
            
            /// <summary>
            /// <b>System.DeviceInterface.Bluetooth.ProductId</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY ProductId => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 8);
            
            /// <summary>
            /// <b>System.DeviceInterface.Bluetooth.ProductVersion</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY ProductVersion => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 9);
            
            /// <summary>
            /// <b>System.DeviceInterface.Bluetooth.ServiceGuid</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY ServiceGuid => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 2);
            
            /// <summary>
            /// <b>System.DeviceInterface.Bluetooth.VendorId</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY VendorId => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 7);
            
            /// <summary>
            /// <b>System.DeviceInterface.Bluetooth.VendorIdSource</b> of <b>VT_UI1</b> type.
            /// </summary>
            public static PROPERTYKEY VendorIdSource => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 6);
        }
        
        public static class Hid
        {
            /// <summary>
            /// <b>System.DeviceInterface.Hid.IsReadOnly</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsReadOnly => new(new("cbf38310-4a17-4310-a1eb-247f0b67593b"), 4);
            
            /// <summary>
            /// <b>System.DeviceInterface.Hid.ProductId</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY ProductId => new(new("cbf38310-4a17-4310-a1eb-247f0b67593b"), 6);
            
            /// <summary>
            /// <b>System.DeviceInterface.Hid.UsageId</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY UsageId => new(new("cbf38310-4a17-4310-a1eb-247f0b67593b"), 3);
            
            /// <summary>
            /// <b>System.DeviceInterface.Hid.UsagePage</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY UsagePage => new(new("cbf38310-4a17-4310-a1eb-247f0b67593b"), 2);
            
            /// <summary>
            /// <b>System.DeviceInterface.Hid.VendorId</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY VendorId => new(new("cbf38310-4a17-4310-a1eb-247f0b67593b"), 5);
            
            /// <summary>
            /// <b>System.DeviceInterface.Hid.VersionNumber</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY VersionNumber => new(new("cbf38310-4a17-4310-a1eb-247f0b67593b"), 7);
        }
        
        public static class Proximity
        {
            /// <summary>
            /// <b>System.DeviceInterface.Proximity.SupportsNfc</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsNfc => new(new("fb3842cd-9e2a-4f83-8fcc-4b0761139ae9"), 2);
        }
        
        public static class Sensors
        {
            /// <summary>
            /// <b>System.DeviceInterface.Sensors.PersistentUniqueId</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY PersistentUniqueId => new(new("d4247382-969d-4f24-bb14-fb9671870bbf"), 9);
            
            /// <summary>
            /// <b>System.DeviceInterface.Sensors.SensorCategory</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY SensorCategory => new(new("d4247382-969d-4f24-bb14-fb9671870bbf"), 3);
        }
        
        public static class Serial
        {
            /// <summary>
            /// <b>System.DeviceInterface.Serial.PortName</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY PortName => new(new("4c6bf15c-4c03-4aac-91f5-64c0f852bcf4"), 4);
            
            /// <summary>
            /// <b>System.DeviceInterface.Serial.UsbProductId</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY UsbProductId => new(new("4c6bf15c-4c03-4aac-91f5-64c0f852bcf4"), 3);
            
            /// <summary>
            /// <b>System.DeviceInterface.Serial.UsbVendorId</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY UsbVendorId => new(new("4c6bf15c-4c03-4aac-91f5-64c0f852bcf4"), 2);
        }
        
        public static class Spb
        {
            /// <summary>
            /// <b>System.DeviceInterface.Spb.ControllerFriendlyName</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY ControllerFriendlyName => new(new("37ebd11f-7e72-4ebc-9d4c-c790f8c277c2"), 2);
        }
        
        public static class WinUsb
        {
            /// <summary>
            /// <b>System.DeviceInterface.WinUsb.DeviceInterfaceClasses</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY DeviceInterfaceClasses => new(new("95e127b5-79cc-4e83-9c9e-8422187b3e0e"), 7);
            
            /// <summary>
            /// <b>System.DeviceInterface.WinUsb.UsbClass</b> of <b>VT_UI1</b> type.
            /// </summary>
            public static PROPERTYKEY UsbClass => new(new("95e127b5-79cc-4e83-9c9e-8422187b3e0e"), 4);
            
            /// <summary>
            /// <b>System.DeviceInterface.WinUsb.UsbProductId</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY UsbProductId => new(new("95e127b5-79cc-4e83-9c9e-8422187b3e0e"), 3);
            
            /// <summary>
            /// <b>System.DeviceInterface.WinUsb.UsbProtocol</b> of <b>VT_UI1</b> type.
            /// </summary>
            public static PROPERTYKEY UsbProtocol => new(new("95e127b5-79cc-4e83-9c9e-8422187b3e0e"), 6);
            
            /// <summary>
            /// <b>System.DeviceInterface.WinUsb.UsbSubClass</b> of <b>VT_UI1</b> type.
            /// </summary>
            public static PROPERTYKEY UsbSubClass => new(new("95e127b5-79cc-4e83-9c9e-8422187b3e0e"), 5);
            
            /// <summary>
            /// <b>System.DeviceInterface.WinUsb.UsbVendorId</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY UsbVendorId => new(new("95e127b5-79cc-4e83-9c9e-8422187b3e0e"), 2);
        }
    }
    
    public static class Devices
    {
        /// <summary>
        /// <b>System.Devices.AppPackageFamilyName</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY AppPackageFamilyName => new(new("51236583-0c4a-4fe8-b81f-166aec13f510"), 100);
        
        /// <summary>
        /// <b>System.Devices.BatteryLife</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY BatteryLife => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 10);
        
        /// <summary>
        /// <b>System.Devices.BatteryPlusCharging</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY BatteryPlusCharging => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 22);
        
        /// <summary>
        /// <b>System.Devices.BatteryPlusChargingText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BatteryPlusChargingText => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 23);
        
        /// <summary>
        /// <b>System.Devices.Category</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Category => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 91);
        
        /// <summary>
        /// <b>System.Devices.CategoryGroup</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY CategoryGroup => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 94);
        
        /// <summary>
        /// <b>System.Devices.CategoryIds</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY CategoryIds => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 90);
        
        /// <summary>
        /// <b>System.Devices.CategoryPlural</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY CategoryPlural => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 92);
        
        /// <summary>
        /// <b>System.Devices.ChallengeAep</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY ChallengeAep => new(new("0774315e-b714-48ec-8de8-8125c077ac11"), 2);
        
        /// <summary>
        /// <b>System.Devices.ChargingState</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY ChargingState => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 11);
        
        /// <summary>
        /// <b>System.Devices.Children</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Children => new(new("4340a6c5-93fa-4706-972c-7b648008a5a7"), 9);
        
        /// <summary>
        /// <b>System.Devices.ClassGuid</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY ClassGuid => new(new("a45c254e-df1c-4efd-8020-67d146a850e0"), 10);
        
        /// <summary>
        /// <b>System.Devices.CompatibleIds</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY CompatibleIds => new(new("a45c254e-df1c-4efd-8020-67d146a850e0"), 4);
        
        /// <summary>
        /// <b>System.Devices.Connected</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Connected => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 55);
        
        /// <summary>
        /// <b>System.Devices.ContainerId</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY ContainerId => new(new("8c7ed206-3f8a-4827-b3ab-ae9e1faefc6c"), 2);
        
        /// <summary>
        /// <b>System.Devices.DefaultTooltip</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DefaultTooltip => new(new("880f70a2-6082-47ac-8aab-a739d1a300c3"), 153);
        
        /// <summary>
        /// <b>System.Devices.DevObjectType</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DevObjectType => new(new("13673f42-a3d6-49f6-b4da-ae46e0c5237c"), 2);
        
        /// <summary>
        /// <b>System.Devices.DeviceCapabilities</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceCapabilities => new(new("a45c254e-df1c-4efd-8020-67d146a850e0"), 17);
        
        /// <summary>
        /// <b>System.Devices.DeviceCharacteristics</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceCharacteristics => new(new("a45c254e-df1c-4efd-8020-67d146a850e0"), 29);
        
        /// <summary>
        /// <b>System.Devices.DeviceDescription1</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceDescription1 => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 81);
        
        /// <summary>
        /// <b>System.Devices.DeviceDescription2</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceDescription2 => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 82);
        
        /// <summary>
        /// <b>System.Devices.DeviceHasProblem</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceHasProblem => new(new("540b947e-8b40-45bc-a8a2-6a0b894cbda2"), 6);
        
        /// <summary>
        /// <b>System.Devices.DeviceInstanceId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceInstanceId => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 256);
        
        /// <summary>
        /// <b>System.Devices.DeviceManufacturer</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeviceManufacturer => new(new("a45c254e-df1c-4efd-8020-67d146a850e0"), 13);
        
        /// <summary>
        /// <b>System.Devices.DiscoveryMethod</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DiscoveryMethod => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 52);
        
        /// <summary>
        /// <b>System.Devices.FriendlyName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FriendlyName => new(new("656a3bb3-ecc0-43fd-8477-4ae0404a96cd"), 12288);
        
        /// <summary>
        /// <b>System.Devices.FunctionPaths</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY FunctionPaths => new(new("d08dd4c0-3a9e-462e-8290-7b636b2576b9"), 3);
        
        /// <summary>
        /// <b>System.Devices.GlyphIcon</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY GlyphIcon => new(new("51236583-0c4a-4fe8-b81f-166aec13f510"), 123);
        
        /// <summary>
        /// <b>System.Devices.HardwareIds</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY HardwareIds => new(new("a45c254e-df1c-4efd-8020-67d146a850e0"), 3);
        
        /// <summary>
        /// <b>System.Devices.Icon</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Icon => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 57);
        
        /// <summary>
        /// <b>System.Devices.InLocalMachineContainer</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY InLocalMachineContainer => new(new("8c7ed206-3f8a-4827-b3ab-ae9e1faefc6c"), 4);
        
        /// <summary>
        /// <b>System.Devices.InterfaceClassGuid</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY InterfaceClassGuid => new(new("026e516e-b814-414b-83cd-856d6fef4822"), 4);
        
        /// <summary>
        /// <b>System.Devices.InterfaceEnabled</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY InterfaceEnabled => new(new("026e516e-b814-414b-83cd-856d6fef4822"), 3);
        
        /// <summary>
        /// <b>System.Devices.InterfacePaths</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY InterfacePaths => new(new("d08dd4c0-3a9e-462e-8290-7b636b2576b9"), 2);
        
        /// <summary>
        /// <b>System.Devices.IpAddress</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY IpAddress => new(new("656a3bb3-ecc0-43fd-8477-4ae0404a96cd"), 12297);
        
        /// <summary>
        /// <b>System.Devices.IsDefault</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsDefault => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 86);
        
        /// <summary>
        /// <b>System.Devices.IsNetworkConnected</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsNetworkConnected => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 85);
        
        /// <summary>
        /// <b>System.Devices.IsShared</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsShared => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 84);
        
        /// <summary>
        /// <b>System.Devices.IsSoftwareInstalling</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsSoftwareInstalling => new(new("83da6326-97a6-4088-9453-a1923f573b29"), 9);
        
        /// <summary>
        /// <b>System.Devices.LaunchDeviceStageFromExplorer</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY LaunchDeviceStageFromExplorer => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 77);
        
        /// <summary>
        /// <b>System.Devices.LocalMachine</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY LocalMachine => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 70);
        
        /// <summary>
        /// <b>System.Devices.LocationPaths</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY LocationPaths => new(new("a45c254e-df1c-4efd-8020-67d146a850e0"), 37);
        
        /// <summary>
        /// <b>System.Devices.Manufacturer</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Manufacturer => new(new("656a3bb3-ecc0-43fd-8477-4ae0404a96cd"), 8192);
        
        /// <summary>
        /// <b>System.Devices.MetadataPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MetadataPath => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 71);
        
        /// <summary>
        /// <b>System.Devices.MissedCalls</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY MissedCalls => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 5);
        
        /// <summary>
        /// <b>System.Devices.ModelId</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY ModelId => new(new("80d81ea6-7473-4b0c-8216-efc11a2c4c8b"), 2);
        
        /// <summary>
        /// <b>System.Devices.ModelName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ModelName => new(new("656a3bb3-ecc0-43fd-8477-4ae0404a96cd"), 8194);
        
        /// <summary>
        /// <b>System.Devices.ModelNumber</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ModelNumber => new(new("656a3bb3-ecc0-43fd-8477-4ae0404a96cd"), 8195);
        
        /// <summary>
        /// <b>System.Devices.NetworkName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY NetworkName => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 7);
        
        /// <summary>
        /// <b>System.Devices.NetworkType</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY NetworkType => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 8);
        
        /// <summary>
        /// <b>System.Devices.NetworkedTooltip</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY NetworkedTooltip => new(new("880f70a2-6082-47ac-8aab-a739d1a300c3"), 152);
        
        /// <summary>
        /// <b>System.Devices.NewPictures</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY NewPictures => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 4);
        
        /// <summary>
        /// <b>System.Devices.NotWorkingProperly</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY NotWorkingProperly => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 83);
        
        /// <summary>
        /// <b>System.Devices.Notification</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Notification => new(new("06704b0c-e830-4c81-9178-91e4e95a80a0"), 3);
        
        /// <summary>
        /// <b>System.Devices.NotificationStore</b> of <b>VT_UNKNOWN</b> type.
        /// </summary>
        public static PROPERTYKEY NotificationStore => new(new("06704b0c-e830-4c81-9178-91e4e95a80a0"), 2);
        
        /// <summary>
        /// <b>System.Devices.Paired</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Paired => new(new("78c34fc8-104a-4aca-9ea4-524d52996e57"), 56);
        
        /// <summary>
        /// <b>System.Devices.Parent</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Parent => new(new("4340a6c5-93fa-4706-972c-7b648008a5a7"), 8);
        
        /// <summary>
        /// <b>System.Devices.PhysicalDeviceLocation</b> of <b>VT_UI1, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY PhysicalDeviceLocation => new(new("540b947e-8b40-45bc-a8a2-6a0b894cbda2"), 9);
        
        /// <summary>
        /// <b>System.Devices.PlaybackPositionPercent</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY PlaybackPositionPercent => new(new("3633de59-6825-4381-a49b-9f6ba13a1471"), 5);
        
        /// <summary>
        /// <b>System.Devices.PlaybackState</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY PlaybackState => new(new("3633de59-6825-4381-a49b-9f6ba13a1471"), 2);
        
        /// <summary>
        /// <b>System.Devices.PlaybackTitle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PlaybackTitle => new(new("3633de59-6825-4381-a49b-9f6ba13a1471"), 3);
        
        /// <summary>
        /// <b>System.Devices.Present</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Present => new(new("540b947e-8b40-45bc-a8a2-6a0b894cbda2"), 5);
        
        /// <summary>
        /// <b>System.Devices.PresentationUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PresentationUrl => new(new("656a3bb3-ecc0-43fd-8477-4ae0404a96cd"), 8198);
        
        /// <summary>
        /// <b>System.Devices.PrimaryCategory</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryCategory => new(new("d08dd4c0-3a9e-462e-8290-7b636b2576b9"), 10);
        
        /// <summary>
        /// <b>System.Devices.RemainingDuration</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY RemainingDuration => new(new("3633de59-6825-4381-a49b-9f6ba13a1471"), 4);
        
        /// <summary>
        /// <b>System.Devices.RestrictedInterface</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY RestrictedInterface => new(new("026e516e-b814-414b-83cd-856d6fef4822"), 6);
        
        /// <summary>
        /// <b>System.Devices.Roaming</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY Roaming => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 9);
        
        /// <summary>
        /// <b>System.Devices.SafeRemovalRequired</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY SafeRemovalRequired => new(new("afd97640-86a3-4210-b67c-289c41aabe55"), 2);
        
        /// <summary>
        /// <b>System.Devices.SchematicName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SchematicName => new(new("026e516e-b814-414b-83cd-856d6fef4822"), 9);
        
        /// <summary>
        /// <b>System.Devices.ServiceAddress</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY ServiceAddress => new(new("656a3bb3-ecc0-43fd-8477-4ae0404a96cd"), 16384);
        
        /// <summary>
        /// <b>System.Devices.ServiceId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ServiceId => new(new("656a3bb3-ecc0-43fd-8477-4ae0404a96cd"), 16385);
        
        /// <summary>
        /// <b>System.Devices.SharedTooltip</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SharedTooltip => new(new("880f70a2-6082-47ac-8aab-a739d1a300c3"), 151);
        
        /// <summary>
        /// <b>System.Devices.SignalStrength</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY SignalStrength => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 2);
        
        /// <summary>
        /// <b>System.Devices.Status</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Status => new(new("d08dd4c0-3a9e-462e-8290-7b636b2576b9"), 259);
        
        /// <summary>
        /// <b>System.Devices.Status1</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Status1 => new(new("d08dd4c0-3a9e-462e-8290-7b636b2576b9"), 257);
        
        /// <summary>
        /// <b>System.Devices.Status2</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Status2 => new(new("d08dd4c0-3a9e-462e-8290-7b636b2576b9"), 258);
        
        /// <summary>
        /// <b>System.Devices.StorageCapacity</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY StorageCapacity => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 12);
        
        /// <summary>
        /// <b>System.Devices.StorageFreeSpace</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY StorageFreeSpace => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 13);
        
        /// <summary>
        /// <b>System.Devices.StorageFreeSpacePercent</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY StorageFreeSpacePercent => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 14);
        
        /// <summary>
        /// <b>System.Devices.TextMessages</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY TextMessages => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 3);
        
        /// <summary>
        /// <b>System.Devices.Voicemail</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY Voicemail => new(new("49cd1f76-5626-4b17-a4e8-18b4aa1a2213"), 6);
        
        /// <summary>
        /// <b>System.Devices.WiaDeviceType</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY WiaDeviceType => new(new("6bdd1fc6-810f-11d0-bec7-08002be2092f"), 2);
        
        /// <summary>
        /// <b>System.Devices.WinPhone8CameraFlags</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY WinPhone8CameraFlags => new(new("b7b4d61c-5a64-4187-a52e-b1539f359099"), 2);
        
        public static class Aep
        {
            /// <summary>
            /// <b>System.Devices.Aep.AepId</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY AepId => new(new("3b2ce006-5e61-4fde-bab8-9b8aac9b26df"), 8);
            
            /// <summary>
            /// <b>System.Devices.Aep.CanPair</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY CanPair => new(new("e7c3fb29-caa7-4f47-8c8b-be59b330d4c5"), 3);
            
            /// <summary>
            /// <b>System.Devices.Aep.Category</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY Category => new(new("a35996ab-11cf-4935-8b61-a6761081ecdf"), 17);
            
            /// <summary>
            /// <b>System.Devices.Aep.ContainerId</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY ContainerId => new(new("e7c3fb29-caa7-4f47-8c8b-be59b330d4c5"), 2);
            
            /// <summary>
            /// <b>System.Devices.Aep.DeviceAddress</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY DeviceAddress => new(new("a35996ab-11cf-4935-8b61-a6761081ecdf"), 12);
            
            /// <summary>
            /// <b>System.Devices.Aep.IsConnected</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsConnected => new(new("a35996ab-11cf-4935-8b61-a6761081ecdf"), 7);
            
            /// <summary>
            /// <b>System.Devices.Aep.IsPaired</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsPaired => new(new("a35996ab-11cf-4935-8b61-a6761081ecdf"), 16);
            
            /// <summary>
            /// <b>System.Devices.Aep.IsPresent</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsPresent => new(new("a35996ab-11cf-4935-8b61-a6761081ecdf"), 9);
            
            /// <summary>
            /// <b>System.Devices.Aep.Manufacturer</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Manufacturer => new(new("a35996ab-11cf-4935-8b61-a6761081ecdf"), 5);
            
            /// <summary>
            /// <b>System.Devices.Aep.ModelId</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY ModelId => new(new("a35996ab-11cf-4935-8b61-a6761081ecdf"), 4);
            
            /// <summary>
            /// <b>System.Devices.Aep.ModelName</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY ModelName => new(new("a35996ab-11cf-4935-8b61-a6761081ecdf"), 3);
            
            /// <summary>
            /// <b>System.Devices.Aep.ProtocolId</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY ProtocolId => new(new("3b2ce006-5e61-4fde-bab8-9b8aac9b26df"), 5);
            
            /// <summary>
            /// <b>System.Devices.Aep.SignalStrength</b> of <b>VT_I4</b> type.
            /// </summary>
            public static PROPERTYKEY SignalStrength => new(new("a35996ab-11cf-4935-8b61-a6761081ecdf"), 6);
            
            public static class Bluetooth
            {
                /// <summary>
                /// <b>System.Devices.Aep.Bluetooth.IssueInquiry</b> of <b>VT_BOOL</b> type.
                /// </summary>
                public static PROPERTYKEY IssueInquiry => new(new("9744311e-7951-4b2e-b6f0-ecb293cac119"), 1);
                
                /// <summary>
                /// <b>System.Devices.Aep.Bluetooth.LastSeenTime</b> of <b>VT_FILETIME</b> type.
                /// </summary>
                public static PROPERTYKEY LastSeenTime => new(new("2bd67d8b-8beb-48d5-87e0-6cda3428040a"), 12);
                
                public static class Cod
                {
                    /// <summary>
                    /// <b>System.Devices.Aep.Bluetooth.Cod.Major</b> of <b>VT_UI2</b> type.
                    /// </summary>
                    public static PROPERTYKEY Major => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 2);
                    
                    /// <summary>
                    /// <b>System.Devices.Aep.Bluetooth.Cod.Minor</b> of <b>VT_UI2</b> type.
                    /// </summary>
                    public static PROPERTYKEY Minor => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 3);
                    
                    public static class Services
                    {
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Cod.Services.Audio</b> of <b>VT_BOOL</b> type.
                        /// </summary>
                        public static PROPERTYKEY Audio => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 10);
                        
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Cod.Services.Capturing</b> of <b>VT_BOOL</b> type.
                        /// </summary>
                        public static PROPERTYKEY Capturing => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 8);
                        
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Cod.Services.Information</b> of <b>VT_BOOL</b> type.
                        /// </summary>
                        public static PROPERTYKEY Information => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 12);
                        
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Cod.Services.LimitedDiscovery</b> of <b>VT_BOOL</b> type.
                        /// </summary>
                        public static PROPERTYKEY LimitedDiscovery => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 4);
                        
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Cod.Services.Networking</b> of <b>VT_BOOL</b> type.
                        /// </summary>
                        public static PROPERTYKEY Networking => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 6);
                        
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Cod.Services.ObjectXfer</b> of <b>VT_BOOL</b> type.
                        /// </summary>
                        public static PROPERTYKEY ObjectXfer => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 9);
                        
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Cod.Services.Positioning</b> of <b>VT_BOOL</b> type.
                        /// </summary>
                        public static PROPERTYKEY Positioning => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 5);
                        
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Cod.Services.Rendering</b> of <b>VT_BOOL</b> type.
                        /// </summary>
                        public static PROPERTYKEY Rendering => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 7);
                        
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Cod.Services.Telephony</b> of <b>VT_BOOL</b> type.
                        /// </summary>
                        public static PROPERTYKEY Telephony => new(new("5fbd34cd-561a-412e-ba98-478a6b0fef1d"), 11);
                    }
                }
                
                public static class Le
                {
                    /// <summary>
                    /// <b>System.Devices.Aep.Bluetooth.Le.ActiveScanning</b> of <b>VT_BOOL</b> type.
                    /// </summary>
                    public static PROPERTYKEY ActiveScanning => new(new("9744311e-7951-4b2e-b6f0-ecb293cac119"), 2);
                    
                    /// <summary>
                    /// <b>System.Devices.Aep.Bluetooth.Le.AddressType</b> of <b>VT_UI1</b> type.
                    /// </summary>
                    public static PROPERTYKEY AddressType => new(new("995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69"), 4);
                    
                    /// <summary>
                    /// <b>System.Devices.Aep.Bluetooth.Le.Appearance</b> of <b>VT_UI2</b> type.
                    /// </summary>
                    public static PROPERTYKEY AppearanceProperty => new(new("995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69"), 1);
                    
                    /// <summary>
                    /// <b>System.Devices.Aep.Bluetooth.Le.IsCallControlClient</b> of <b>VT_BOOL</b> type.
                    /// </summary>
                    public static PROPERTYKEY IsCallControlClient => new(new("995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69"), 12);
                    
                    /// <summary>
                    /// <b>System.Devices.Aep.Bluetooth.Le.IsConnectable</b> of <b>VT_BOOL</b> type.
                    /// </summary>
                    public static PROPERTYKEY IsConnectable => new(new("995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69"), 8);
                    
                    /// <summary>
                    /// <b>System.Devices.Aep.Bluetooth.Le.ScanInterval</b> of <b>VT_UI2</b> type.
                    /// </summary>
                    public static PROPERTYKEY ScanInterval => new(new("9744311e-7951-4b2e-b6f0-ecb293cac119"), 3);
                    
                    /// <summary>
                    /// <b>System.Devices.Aep.Bluetooth.Le.ScanResponse</b> of <b>VT_BLOB</b> type.
                    /// </summary>
                    public static PROPERTYKEY ScanResponse => new(new("995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69"), 3);
                    
                    /// <summary>
                    /// <b>System.Devices.Aep.Bluetooth.Le.ScanWindow</b> of <b>VT_UI2</b> type.
                    /// </summary>
                    public static PROPERTYKEY ScanWindow => new(new("9744311e-7951-4b2e-b6f0-ecb293cac119"), 4);
                    
                    public static class Appearance
                    {
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Le.Appearance.Category</b> of <b>VT_UI2</b> type.
                        /// </summary>
                        public static PROPERTYKEY Category => new(new("995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69"), 5);
                        
                        /// <summary>
                        /// <b>System.Devices.Aep.Bluetooth.Le.Appearance.Subcategory</b> of <b>VT_UI2</b> type.
                        /// </summary>
                        public static PROPERTYKEY Subcategory => new(new("995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69"), 6);
                    }
                }
            }
            
            public static class PointOfService
            {
                /// <summary>
                /// <b>System.Devices.Aep.PointOfService.ConnectionTypes</b> of <b>VT_I4</b> type.
                /// </summary>
                public static PROPERTYKEY ConnectionTypes => new(new("d4bf61b3-442e-4ada-882d-fa7b70c832d9"), 6);
            }
        }
        
        public static class AepContainer
        {
            /// <summary>
            /// <b>System.Devices.AepContainer.CanPair</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY CanPair => new(new("0bba1ede-7566-4f47-90ec-25fc567ced2a"), 3);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.Categories</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY Categories => new(new("0bba1ede-7566-4f47-90ec-25fc567ced2a"), 9);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.Children</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY Children => new(new("0bba1ede-7566-4f47-90ec-25fc567ced2a"), 2);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.ContainerId</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY ContainerId => new(new("0bba1ede-7566-4f47-90ec-25fc567ced2a"), 12);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.IsPaired</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsPaired => new(new("0bba1ede-7566-4f47-90ec-25fc567ced2a"), 4);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.IsPresent</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsPresent => new(new("0bba1ede-7566-4f47-90ec-25fc567ced2a"), 11);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.Manufacturer</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Manufacturer => new(new("0bba1ede-7566-4f47-90ec-25fc567ced2a"), 6);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.ModelIds</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY ModelIds => new(new("0bba1ede-7566-4f47-90ec-25fc567ced2a"), 8);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.ModelName</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY ModelName => new(new("0bba1ede-7566-4f47-90ec-25fc567ced2a"), 7);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.ProtocolIds</b> of <b>VT_CLSID, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY ProtocolIds => new(new("0bba1ede-7566-4f47-90ec-25fc567ced2a"), 13);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportedUriSchemes</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY SupportedUriSchemes => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 5);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsAudio</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsAudio => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 2);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsCapturing</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsCapturing => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 11);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsImages</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsImages => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 4);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsInformation</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsInformation => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 14);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsLimitedDiscovery</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsLimitedDiscovery => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 7);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsNetworking</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsNetworking => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 9);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsObjectTransfer</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsObjectTransfer => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 12);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsPositioning</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsPositioning => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 8);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsRendering</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsRendering => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 10);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsTelephony</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsTelephony => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 13);
            
            /// <summary>
            /// <b>System.Devices.AepContainer.SupportsVideo</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SupportsVideo => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 3);
            
            public static class DialProtocol
            {
                /// <summary>
                /// <b>System.Devices.AepContainer.DialProtocol.InstalledApplications</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
                /// </summary>
                public static PROPERTYKEY InstalledApplications => new(new("6af55d45-38db-4495-acb0-d4728a3b8314"), 6);
            }
        }
        
        public static class AepService
        {
            /// <summary>
            /// <b>System.Devices.AepService.AepId</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY AepId => new(new("c9c141a9-1b4c-4f17-a9d1-f298538cadb8"), 6);
            
            /// <summary>
            /// <b>System.Devices.AepService.ContainerId</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY ContainerId => new(new("71724756-3e74-4432-9b59-e7b2f668a593"), 4);
            
            /// <summary>
            /// <b>System.Devices.AepService.FriendlyName</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY FriendlyName => new(new("71724756-3e74-4432-9b59-e7b2f668a593"), 2);
            
            /// <summary>
            /// <b>System.Devices.AepService.ParentAepIsPaired</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY ParentAepIsPaired => new(new("c9c141a9-1b4c-4f17-a9d1-f298538cadb8"), 7);
            
            /// <summary>
            /// <b>System.Devices.AepService.ProtocolId</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY ProtocolId => new(new("c9c141a9-1b4c-4f17-a9d1-f298538cadb8"), 5);
            
            /// <summary>
            /// <b>System.Devices.AepService.ServiceClassId</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY ServiceClassId => new(new("71724756-3e74-4432-9b59-e7b2f668a593"), 3);
            
            /// <summary>
            /// <b>System.Devices.AepService.ServiceId</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY ServiceId => new(new("c9c141a9-1b4c-4f17-a9d1-f298538cadb8"), 2);
            
            public static class Bluetooth
            {
                /// <summary>
                /// <b>System.Devices.AepService.Bluetooth.CacheMode</b> of <b>VT_UI1</b> type.
                /// </summary>
                public static PROPERTYKEY CacheMode => new(new("9744311e-7951-4b2e-b6f0-ecb293cac119"), 5);
                
                /// <summary>
                /// <b>System.Devices.AepService.Bluetooth.ServiceGuid</b> of <b>VT_CLSID</b> type.
                /// </summary>
                public static PROPERTYKEY ServiceGuid => new(new("a399aac7-c265-474e-b073-ffce57721716"), 2);
                
                /// <summary>
                /// <b>System.Devices.AepService.Bluetooth.TargetDevice</b> of <b>VT_UI8</b> type.
                /// </summary>
                public static PROPERTYKEY TargetDevice => new(new("9744311e-7951-4b2e-b6f0-ecb293cac119"), 6);
            }
            
            public static class IoT
            {
                /// <summary>
                /// <b>System.Devices.AepService.IoT.ServiceInterfaces</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
                /// </summary>
                public static PROPERTYKEY ServiceInterfaces => new(new("79d94e82-4d79-45aa-821a-74858b4e4ca6"), 2);
            }
        }
        
        public static class ApoDevice
        {
            /// <summary>
            /// <b>System.Devices.ApoDevice.PreviousInstallDate</b> of <b>VT_FILETIME</b> type.
            /// </summary>
            public static PROPERTYKEY PreviousInstallDate => new(new("97bfacc3-afc8-4174-a162-2ee77cb17704"), 2);
        }
        
        public static class AudioDevice
        {
            /// <summary>
            /// <b>System.Devices.AudioDevice.RawProcessingSupported</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY RawProcessingSupported => new(new("8943b373-388c-4395-b557-bc6dbaffafdb"), 2);
            
            /// <summary>
            /// <b>System.Devices.AudioDevice.SpeechProcessingSupported</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY SpeechProcessingSupported => new(new("fb1de864-e06d-47f4-82a6-8a0aef44493c"), 2);
            
            public static class Microphone
            {
                /// <summary>
                /// <b>System.Devices.AudioDevice.Microphone.IsFarField</b> of <b>VT_BOOL</b> type.
                /// </summary>
                public static PROPERTYKEY IsFarField => new(new("8943b373-388c-4395-b557-bc6dbaffafdb"), 6);
                
                /// <summary>
                /// <b>System.Devices.AudioDevice.Microphone.SensitivityInDbfs</b> of <b>VT_R8</b> type.
                /// </summary>
                public static PROPERTYKEY SensitivityInDbfs => new(new("8943b373-388c-4395-b557-bc6dbaffafdb"), 3);
                
                /// <summary>
                /// <b>System.Devices.AudioDevice.Microphone.SensitivityInDbfs2</b> of <b>VT_R8</b> type.
                /// </summary>
                public static PROPERTYKEY SensitivityInDbfs2 => new(new("8943b373-388c-4395-b557-bc6dbaffafdb"), 5);
                
                /// <summary>
                /// <b>System.Devices.AudioDevice.Microphone.SignalToNoiseRatioInDb</b> of <b>VT_R8</b> type.
                /// </summary>
                public static PROPERTYKEY SignalToNoiseRatioInDb => new(new("8943b373-388c-4395-b557-bc6dbaffafdb"), 4);
            }
        }
        
        public static class AudioPosture
        {
            /// <summary>
            /// <b>System.Devices.AudioPosture.Support</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY Support => new(new("766148e8-87ee-490c-8f9b-1c3686f2d121"), 2);
        }
        
        public static class DialProtocol
        {
            /// <summary>
            /// <b>System.Devices.DialProtocol.InstalledApplications</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY InstalledApplications => new(new("6845cc72-1b71-48c3-af86-b09171a19b14"), 3);
        }
        
        public static class Dnssd
        {
            /// <summary>
            /// <b>System.Devices.Dnssd.Domain</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Domain => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 3);
            
            /// <summary>
            /// <b>System.Devices.Dnssd.FullName</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY FullName => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 5);
            
            /// <summary>
            /// <b>System.Devices.Dnssd.HostName</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY HostName => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 7);
            
            /// <summary>
            /// <b>System.Devices.Dnssd.InstanceName</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY InstanceName => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 4);
            
            /// <summary>
            /// <b>System.Devices.Dnssd.NetworkAdapterId</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY NetworkAdapterId => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 11);
            
            /// <summary>
            /// <b>System.Devices.Dnssd.PortNumber</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY PortNumber => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 12);
            
            /// <summary>
            /// <b>System.Devices.Dnssd.Priority</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY Priority => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 9);
            
            /// <summary>
            /// <b>System.Devices.Dnssd.ServiceName</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY ServiceName => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 2);
            
            /// <summary>
            /// <b>System.Devices.Dnssd.TextAttributes</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY TextAttributes => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 6);
            
            /// <summary>
            /// <b>System.Devices.Dnssd.Ttl</b> of <b>VT_UI4</b> type.
            /// </summary>
            public static PROPERTYKEY Ttl => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 10);
            
            /// <summary>
            /// <b>System.Devices.Dnssd.Weight</b> of <b>VT_UI2</b> type.
            /// </summary>
            public static PROPERTYKEY Weight => new(new("bf79c0ab-bb74-4cee-b070-470b5ae202ea"), 8);
        }
        
        public static class MicrophoneArray
        {
            /// <summary>
            /// <b>System.Devices.MicrophoneArray.Geometry</b> of <b>VT_UI1, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY Geometry => new(new("a1829ea2-27eb-459e-935d-b2fad7b07762"), 2);
        }
        
        public static class Notifications
        {
            /// <summary>
            /// <b>System.Devices.Notifications.LowBattery</b> of <b>VT_UI1</b> type.
            /// </summary>
            public static PROPERTYKEY LowBattery => new(new("c4c07f2b-8524-4e66-ae3a-a6235f103beb"), 2);
            
            /// <summary>
            /// <b>System.Devices.Notifications.MissedCall</b> of <b>VT_UI1</b> type.
            /// </summary>
            public static PROPERTYKEY MissedCall => new(new("6614ef48-4efe-4424-9eda-c79f404edf3e"), 2);
            
            /// <summary>
            /// <b>System.Devices.Notifications.NewMessage</b> of <b>VT_UI1</b> type.
            /// </summary>
            public static PROPERTYKEY NewMessage => new(new("2be9260a-2012-4742-a555-f41b638b7dcb"), 2);
            
            /// <summary>
            /// <b>System.Devices.Notifications.NewVoicemail</b> of <b>VT_UI1</b> type.
            /// </summary>
            public static PROPERTYKEY NewVoicemail => new(new("59569556-0a08-4212-95b9-fae2ad6413db"), 2);
            
            /// <summary>
            /// <b>System.Devices.Notifications.StorageFull</b> of <b>VT_UI8</b> type.
            /// </summary>
            public static PROPERTYKEY StorageFull => new(new("a0e00ee1-f0c7-4d41-b8e7-26a7bd8d38b0"), 2);
            
            /// <summary>
            /// <b>System.Devices.Notifications.StorageFullLinkText</b> of <b>VT_UI8</b> type.
            /// </summary>
            public static PROPERTYKEY StorageFullLinkText => new(new("a0e00ee1-f0c7-4d41-b8e7-26a7bd8d38b0"), 3);
        }
        
        public static class Panel
        {
            /// <summary>
            /// <b>System.Devices.Panel.PanelGroup</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY PanelGroup => new(new("8dbc9c86-97a9-4bff-9bc6-bfe95d3e6dad"), 3);
            
            /// <summary>
            /// <b>System.Devices.Panel.PanelId</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY PanelId => new(new("8dbc9c86-97a9-4bff-9bc6-bfe95d3e6dad"), 2);
        }
        
        public static class PhoneLineTransportDevice
        {
            /// <summary>
            /// <b>System.Devices.PhoneLineTransportDevice.Connected</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY Connected => new(new("aecf2fe8-1d00-4fee-8a6d-a70d719b772b"), 2);
        }
        
        public static class SmartCards
        {
            /// <summary>
            /// <b>System.Devices.SmartCards.ReaderKind</b> of <b>VT_UI1</b> type.
            /// </summary>
            public static PROPERTYKEY ReaderKind => new(new("d6b5b883-18bd-4b4d-b2ec-9e38affeda82"), 2);
        }
        
        public static class WiFi
        {
            /// <summary>
            /// <b>System.Devices.WiFi.InterfaceGuid</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY InterfaceGuid => new(new("ef1167eb-cbfc-4341-a568-a7c91a68982c"), 2);
        }
        
        public static class WiFiDirect
        {
            /// <summary>
            /// <b>System.Devices.WiFiDirect.DeviceAddress</b> of <b>VT_UI1, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY DeviceAddress => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 13);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.GroupId</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY GroupId => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 4);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.InformationElements</b> of <b>VT_UI1, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY InformationElements => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 12);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.InterfaceAddress</b> of <b>VT_UI1, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY InterfaceAddress => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 2);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.InterfaceGuid</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY InterfaceGuid => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 3);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.IsConnected</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsConnected => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 5);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.IsLegacyDevice</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsLegacyDevice => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 7);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.IsMiracastLcpSupported</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsMiracastLcpSupported => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 9);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.IsVisible</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsVisible => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 6);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.MiracastVersion</b> of <b>VT_UI4</b> type.
            /// </summary>
            public static PROPERTYKEY MiracastVersion => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 8);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.Services</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY Services => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 10);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirect.SupportedChannelList</b> of <b>VT_UI1, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY SupportedChannelList => new(new("1506935d-e3e7-450f-8637-82233ebe5f6e"), 11);
        }
        
        public static class WiFiDirectServices
        {
            /// <summary>
            /// <b>System.Devices.WiFiDirectServices.AdvertisementId</b> of <b>VT_UI4</b> type.
            /// </summary>
            public static PROPERTYKEY AdvertisementId => new(new("31b37743-7c5e-4005-93e6-e953f92b82e9"), 5);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirectServices.RequestServiceInformation</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY RequestServiceInformation => new(new("31b37743-7c5e-4005-93e6-e953f92b82e9"), 7);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirectServices.ServiceAddress</b> of <b>VT_UI1, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY ServiceAddress => new(new("31b37743-7c5e-4005-93e6-e953f92b82e9"), 2);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirectServices.ServiceConfigMethods</b> of <b>VT_UI4</b> type.
            /// </summary>
            public static PROPERTYKEY ServiceConfigMethods => new(new("31b37743-7c5e-4005-93e6-e953f92b82e9"), 6);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirectServices.ServiceInformation</b> of <b>VT_UI1, VT_VECTOR</b> type.
            /// </summary>
            public static PROPERTYKEY ServiceInformation => new(new("31b37743-7c5e-4005-93e6-e953f92b82e9"), 4);
            
            /// <summary>
            /// <b>System.Devices.WiFiDirectServices.ServiceName</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY ServiceName => new(new("31b37743-7c5e-4005-93e6-e953f92b82e9"), 3);
        }
        
        public static class Wwan
        {
            /// <summary>
            /// <b>System.Devices.Wwan.InterfaceGuid</b> of <b>VT_CLSID</b> type.
            /// </summary>
            public static PROPERTYKEY InterfaceGuid => new(new("ff1167eb-cbfc-4341-a568-a7c91a68982c"), 2);
        }
    }
    
    public static class Document
    {
        /// <summary>
        /// <b>System.Document.ByteCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY ByteCount => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 4);
        
        /// <summary>
        /// <b>System.Document.CharacterCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY CharacterCount => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 16);
        
        /// <summary>
        /// <b>System.Document.ClientID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ClientID => new(new("276d7bb0-5b34-4fb0-aa4b-158ed12a1809"), 100);
        
        /// <summary>
        /// <b>System.Document.Contributor</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Contributor => new(new("f334115e-da1b-4509-9b3d-119504dc7abb"), 100);
        
        /// <summary>
        /// <b>System.Document.DateCreated</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateCreated => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 12);
        
        /// <summary>
        /// <b>System.Document.DatePrinted</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DatePrinted => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 11);
        
        /// <summary>
        /// <b>System.Document.DateSaved</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateSaved => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 13);
        
        /// <summary>
        /// <b>System.Document.Division</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Division => new(new("1e005ee6-bf27-428b-b01c-79676acd2870"), 100);
        
        /// <summary>
        /// <b>System.Document.DocumentID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DocumentID => new(new("e08805c8-e395-40df-80d2-54f0d6c43154"), 100);
        
        /// <summary>
        /// <b>System.Document.HiddenSlideCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY HiddenSlideCount => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 9);
        
        /// <summary>
        /// <b>System.Document.LastAuthor</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LastAuthor => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 8);
        
        /// <summary>
        /// <b>System.Document.LineCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY LineCount => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 5);
        
        /// <summary>
        /// <b>System.Document.LinksDirty</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY LinksDirty => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 16);
        
        /// <summary>
        /// <b>System.Document.Manager</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Manager => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 14);
        
        /// <summary>
        /// <b>System.Document.MultimediaClipCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY MultimediaClipCount => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 10);
        
        /// <summary>
        /// <b>System.Document.NoteCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY NoteCount => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 8);
        
        /// <summary>
        /// <b>System.Document.PageCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY PageCount => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 14);
        
        /// <summary>
        /// <b>System.Document.ParagraphCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY ParagraphCount => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 6);
        
        /// <summary>
        /// <b>System.Document.PresentationFormat</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PresentationFormat => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 3);
        
        /// <summary>
        /// <b>System.Document.RevisionNumber</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RevisionNumber => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 9);
        
        /// <summary>
        /// <b>System.Document.Scale</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Scale => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 11);
        
        /// <summary>
        /// <b>System.Document.Security</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY Security => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 19);
        
        /// <summary>
        /// <b>System.Document.SlideCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY SlideCount => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 7);
        
        /// <summary>
        /// <b>System.Document.Template</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Template => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 7);
        
        /// <summary>
        /// <b>System.Document.TotalEditingTime</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY TotalEditingTime => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 10);
        
        /// <summary>
        /// <b>System.Document.Version</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Version => new(new("d5cdd502-2e9c-101b-9397-08002b2cf9ae"), 29);
        
        /// <summary>
        /// <b>System.Document.WordCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY WordCount => new(new("f29f85e0-4ff9-1068-ab91-08002b27b3d9"), 15);
    }
    
    public static class Extensions
    {
        /// <summary>
        /// <b>System.Extensions.BlockedCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY BlockedCount => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 5);
        
        /// <summary>
        /// <b>System.Extensions.CLSID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CLSID => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 6);
        
        /// <summary>
        /// <b>System.Extensions.DateLastUsed</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateLastUsed => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 3);
        
        /// <summary>
        /// <b>System.Extensions.FileName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FileName => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 10);
        
        /// <summary>
        /// <b>System.Extensions.FilePath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FilePath => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 11);
        
        /// <summary>
        /// <b>System.Extensions.Flags</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Flags => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 12);
        
        /// <summary>
        /// <b>System.Extensions.Status</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Status => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 7);
        
        /// <summary>
        /// <b>System.Extensions.Suspect</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Suspect => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 9);
        
        /// <summary>
        /// <b>System.Extensions.Type</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Type => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 2);
        
        /// <summary>
        /// <b>System.Extensions.UsedCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY UsedCount => new(new("3f5d9b45-5e9f-4d5c-8a5e-403181bf177b"), 4);
    }
    
    public static class Fonts
    {
        /// <summary>
        /// <b>System.Fonts.ActiveStatus</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ActiveStatus => new(new("d6cf9145-d365-471b-bcb8-f0b4a96b891c"), 100);
        
        /// <summary>
        /// <b>System.Fonts.Category</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Category => new(new("4b486401-5468-4381-9b5a-42df4cb49f53"), 100);
        
        /// <summary>
        /// <b>System.Fonts.CollectionName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CollectionName => new(new("f3aecac4-5b8d-436a-ad0c-64ab194fdaf3"), 100);
        
        /// <summary>
        /// <b>System.Fonts.Description</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Description => new(new("3d658d4d-bc38-464a-b555-418d554a8df8"), 100);
        
        /// <summary>
        /// <b>System.Fonts.DesignedFor</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DesignedFor => new(new("5741cf9c-56fe-485b-8901-4786449e188d"), 100);
        
        /// <summary>
        /// <b>System.Fonts.FamilyName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FamilyName => new(new("de9e220b-41d4-4690-8b6b-3d89e231eef1"), 100);
        
        /// <summary>
        /// <b>System.Fonts.FileNames</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY FileNames => new(new("4530d076-b598-4a81-8813-9b11286ef6ea"), 7);
        
        /// <summary>
        /// <b>System.Fonts.FontEmbeddability</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY FontEmbeddability => new(new("4530d076-b598-4a81-8813-9b11286ef6ea"), 2);
        
        /// <summary>
        /// <b>System.Fonts.Styles</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Styles => new(new("596fd41b-af9b-4ba8-9b49-33b16f16678c"), 100);
        
        /// <summary>
        /// <b>System.Fonts.Type</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Type => new(new("4530d076-b598-4a81-8813-9b11286ef6ea"), 5);
        
        /// <summary>
        /// <b>System.Fonts.Vendors</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Vendors => new(new("49753869-849c-4323-a41f-26d73f28b53b"), 100);
        
        /// <summary>
        /// <b>System.Fonts.Version</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Version => new(new("fec7952b-4bf0-4c03-b6e1-2796818b7ca9"), 100);
    }
    
    public static class GPS
    {
        /// <summary>
        /// <b>System.GPS.Altitude</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY Altitude => new(new("827edb4f-5b73-44a7-891d-fdffabea35ca"), 100);
        
        /// <summary>
        /// <b>System.GPS.AltitudeDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY AltitudeDenominator => new(new("78342dcb-e358-4145-ae9a-6bfe4e0f9f51"), 100);
        
        /// <summary>
        /// <b>System.GPS.AltitudeNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY AltitudeNumerator => new(new("2dad1eb7-816d-40d3-9ec3-c9773be2aade"), 100);
        
        /// <summary>
        /// <b>System.GPS.AltitudeRef</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY AltitudeRef => new(new("46ac629d-75ea-4515-867f-6dc4321c5844"), 100);
        
        /// <summary>
        /// <b>System.GPS.AreaInformation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AreaInformation => new(new("972e333e-ac7e-49f1-8adf-a70d07a9bcab"), 100);
        
        /// <summary>
        /// <b>System.GPS.DOP</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY DOP => new(new("0cf8fb02-1837-42f1-a697-a7017aa289b9"), 100);
        
        /// <summary>
        /// <b>System.GPS.DOPDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DOPDenominator => new(new("a0be94c5-50ba-487b-bd35-0654be8881ed"), 100);
        
        /// <summary>
        /// <b>System.GPS.DOPNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DOPNumerator => new(new("47166b16-364f-4aa0-9f31-e2ab3df449c3"), 100);
        
        /// <summary>
        /// <b>System.GPS.Date</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY Date => new(new("3602c812-0f3b-45f0-85ad-603468d69423"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestBearing</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY DestBearing => new(new("c66d4b3c-e888-47cc-b99f-9dca3ee34dea"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestBearingDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DestBearingDenominator => new(new("7abcf4f8-7c3f-4988-ac91-8d2c2e97eca5"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestBearingNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DestBearingNumerator => new(new("ba3b1da9-86ee-4b5d-a2a4-a271a429f0cf"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestBearingRef</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DestBearingRef => new(new("9ab84393-2a0f-4b75-bb22-7279786977cb"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestDistance</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY DestDistance => new(new("a93eae04-6804-4f24-ac81-09b266452118"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestDistanceDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DestDistanceDenominator => new(new("9bc2c99b-ac71-4127-9d1c-2596d0d7dcb7"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestDistanceNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DestDistanceNumerator => new(new("2bda47da-08c6-4fe1-80bc-a72fc517c5d0"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestDistanceRef</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DestDistanceRef => new(new("ed4df2d3-8695-450b-856f-f5c1c53acb66"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestLatitude</b> of <b>VT_R8, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DestLatitude => new(new("9d1d7cc5-5c39-451c-86b3-928e2d18cc47"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestLatitudeDenominator</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DestLatitudeDenominator => new(new("3a372292-7fca-49a7-99d5-e47bb2d4e7ab"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestLatitudeNumerator</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DestLatitudeNumerator => new(new("ecf4b6f6-d5a6-433c-bb92-4076650fc890"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestLatitudeRef</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DestLatitudeRef => new(new("cea820b9-ce61-4885-a128-005d9087c192"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestLongitude</b> of <b>VT_R8, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DestLongitude => new(new("47a96261-cb4c-4807-8ad3-40b9d9dbc6bc"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestLongitudeDenominator</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DestLongitudeDenominator => new(new("425d69e5-48ad-4900-8d80-6eb6b8d0ac86"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestLongitudeNumerator</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DestLongitudeNumerator => new(new("a3250282-fb6d-48d5-9a89-dbcace75cccf"), 100);
        
        /// <summary>
        /// <b>System.GPS.DestLongitudeRef</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DestLongitudeRef => new(new("182c1ea6-7c1c-4083-ab4b-ac6c9f4ed128"), 100);
        
        /// <summary>
        /// <b>System.GPS.Differential</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY Differential => new(new("aaf4ee25-bd3b-4dd7-bfc4-47f77bb00f6d"), 100);
        
        /// <summary>
        /// <b>System.GPS.ImgDirection</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY ImgDirection => new(new("16473c91-d017-4ed9-ba4d-b6baa55dbcf8"), 100);
        
        /// <summary>
        /// <b>System.GPS.ImgDirectionDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ImgDirectionDenominator => new(new("10b24595-41a2-4e20-93c2-5761c1395f32"), 100);
        
        /// <summary>
        /// <b>System.GPS.ImgDirectionNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ImgDirectionNumerator => new(new("dc5877c7-225f-45f7-bac7-e81334b6130a"), 100);
        
        /// <summary>
        /// <b>System.GPS.ImgDirectionRef</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ImgDirectionRef => new(new("a4aaa5b7-1ad0-445f-811a-0f8f6e67f6b5"), 100);
        
        /// <summary>
        /// <b>System.GPS.Latitude</b> of <b>VT_R8, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Latitude => new(new("8727cfff-4868-4ec6-ad5b-81b98521d1ab"), 100);
        
        /// <summary>
        /// <b>System.GPS.LatitudeDecimal</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY LatitudeDecimal => new(new("0f55cde2-4f49-450d-92c1-dcd16301b1b7"), 100);
        
        /// <summary>
        /// <b>System.GPS.LatitudeDenominator</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY LatitudeDenominator => new(new("16e634ee-2bff-497b-bd8a-4341ad39eeb9"), 100);
        
        /// <summary>
        /// <b>System.GPS.LatitudeNumerator</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY LatitudeNumerator => new(new("7ddaaad1-ccc8-41ae-b750-b2cb8031aea2"), 100);
        
        /// <summary>
        /// <b>System.GPS.LatitudeRef</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LatitudeRef => new(new("029c0252-5b86-46c7-aca0-2769ffc8e3d4"), 100);
        
        /// <summary>
        /// <b>System.GPS.Longitude</b> of <b>VT_R8, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Longitude => new(new("c4c4dbb2-b593-466b-bbda-d03d27d5e43a"), 100);
        
        /// <summary>
        /// <b>System.GPS.LongitudeDecimal</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY LongitudeDecimal => new(new("4679c1b5-844d-4590-baf5-f322231f1b81"), 100);
        
        /// <summary>
        /// <b>System.GPS.LongitudeDenominator</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY LongitudeDenominator => new(new("be6e176c-4534-4d2c-ace5-31dedac1606b"), 100);
        
        /// <summary>
        /// <b>System.GPS.LongitudeNumerator</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY LongitudeNumerator => new(new("02b0f689-a914-4e45-821d-1dda452ed2c4"), 100);
        
        /// <summary>
        /// <b>System.GPS.LongitudeRef</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LongitudeRef => new(new("33dcf22b-28d5-464c-8035-1ee9efd25278"), 100);
        
        /// <summary>
        /// <b>System.GPS.MapDatum</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MapDatum => new(new("2ca2dae6-eddc-407d-bef1-773942abfa95"), 100);
        
        /// <summary>
        /// <b>System.GPS.MeasureMode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MeasureMode => new(new("a015ed5d-aaea-4d58-8a86-3c586920ea0b"), 100);
        
        /// <summary>
        /// <b>System.GPS.ProcessingMethod</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProcessingMethod => new(new("59d49e61-840f-4aa9-a939-e2099b7f6399"), 100);
        
        /// <summary>
        /// <b>System.GPS.Satellites</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Satellites => new(new("467ee575-1f25-4557-ad4e-b8b58b0d9c15"), 100);
        
        /// <summary>
        /// <b>System.GPS.Speed</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY Speed => new(new("da5d0862-6e76-4e1b-babd-70021bd25494"), 100);
        
        /// <summary>
        /// <b>System.GPS.SpeedDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY SpeedDenominator => new(new("7d122d5a-ae5e-4335-8841-d71e7ce72f53"), 100);
        
        /// <summary>
        /// <b>System.GPS.SpeedNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY SpeedNumerator => new(new("acc9ce3d-c213-4942-8b48-6d0820f21c6d"), 100);
        
        /// <summary>
        /// <b>System.GPS.SpeedRef</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SpeedRef => new(new("ecf7f4c9-544f-4d6d-9d98-8ad79adaf453"), 100);
        
        /// <summary>
        /// <b>System.GPS.Status</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Status => new(new("125491f4-818f-46b2-91b5-d537753617b2"), 100);
        
        /// <summary>
        /// <b>System.GPS.Track</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY Track => new(new("76c09943-7c33-49e3-9e7e-cdba872cfada"), 100);
        
        /// <summary>
        /// <b>System.GPS.TrackDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY TrackDenominator => new(new("c8d1920c-01f6-40c0-ac86-2f3a4ad00770"), 100);
        
        /// <summary>
        /// <b>System.GPS.TrackNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY TrackNumerator => new(new("702926f4-44a6-43e1-ae71-45627116893b"), 100);
        
        /// <summary>
        /// <b>System.GPS.TrackRef</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TrackRef => new(new("35dbe6fe-44c3-4400-aaae-d2c799c407e8"), 100);
        
        /// <summary>
        /// <b>System.GPS.VersionID</b> of <b>VT_UI1, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY VersionID => new(new("22704da4-c6b2-4a99-8e56-f16df8c92599"), 100);
    }
    
    public static class Game
    {
        /// <summary>
        /// <b>System.Game.Description</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Description => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 18);
        
        /// <summary>
        /// <b>System.Game.Developer</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Developer => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 22);
        
        /// <summary>
        /// <b>System.Game.DeveloperNoLink</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeveloperNoLink => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 57);
        
        /// <summary>
        /// <b>System.Game.Genre</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Genre => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 24);
        
        /// <summary>
        /// <b>System.Game.HomepageTask</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY HomepageTask => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 42);
        
        /// <summary>
        /// <b>System.Game.IntUpdateStatus</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY IntUpdateStatus => new(new("bb44403b-1399-4650-95eb-03c53a57c2cf"), 60);
        
        /// <summary>
        /// <b>System.Game.LastPlayed</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY LastPlayed => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 36);
        
        /// <summary>
        /// <b>System.Game.NonOptionalRatingDescriptors</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY NonOptionalRatingDescriptors => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 38);
        
        /// <summary>
        /// <b>System.Game.OptionalRatingDescriptors</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY OptionalRatingDescriptors => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 37);
        
        /// <summary>
        /// <b>System.Game.Publisher</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Publisher => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 21);
        
        /// <summary>
        /// <b>System.Game.PublisherNoLink</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PublisherNoLink => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 58);
        
        /// <summary>
        /// <b>System.Game.Rating</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Rating => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 28);
        
        /// <summary>
        /// <b>System.Game.RatingDescription</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RatingDescription => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 30);
        
        /// <summary>
        /// <b>System.Game.RatingDescriptors</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY RatingDescriptors => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 31);
        
        /// <summary>
        /// <b>System.Game.RatingLong</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RatingLong => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 29);
        
        /// <summary>
        /// <b>System.Game.RatingProviderName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RatingProviderName => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 39);
        
        /// <summary>
        /// <b>System.Game.RatingResourceId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RatingResourceId => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 34);
        
        /// <summary>
        /// <b>System.Game.RatingResourcePath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RatingResourcePath => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 32);
        
        /// <summary>
        /// <b>System.Game.RatingResourceType</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RatingResourceType => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 33);
        
        /// <summary>
        /// <b>System.Game.RatingUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RatingUrl => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 35);
        
        /// <summary>
        /// <b>System.Game.ReleaseDate</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY ReleaseDate => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 19);
        
        /// <summary>
        /// <b>System.Game.Restricted</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Restricted => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 52);
        
        /// <summary>
        /// <b>System.Game.RichApplicationName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RichApplicationName => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 53);
        
        /// <summary>
        /// <b>System.Game.RichComment</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RichComment => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 56);
        
        /// <summary>
        /// <b>System.Game.RichLevel</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RichLevel => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 55);
        
        /// <summary>
        /// <b>System.Game.RichSaveName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RichSaveName => new(new("26b8d54f-371f-4aeb-8a84-9224aea4d40a"), 54);
        
        /// <summary>
        /// <b>System.Game.Type</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Type => new(new("6ccd0131-c397-4744-b2d8-d2c13f457026"), 80);
        
        /// <summary>
        /// <b>System.Game.UpdateStatus</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UpdateStatus => new(new("c9b88dba-04db-4887-a200-cf0d3afe1146"), 99);
        
        /// <summary>
        /// <b>System.Game.WinSPRMinimum</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY WinSPRMinimum => new(new("e6c3d9ad-7b32-4efe-a167-0a868ffdf3af"), 100);
        
        /// <summary>
        /// <b>System.Game.WinSPRRecommended</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY WinSPRRecommended => new(new("b771b352-8692-42e6-ac33-cc7b062ad950"), 100);
    }
    
    public static class Generic
    {
        /// <summary>
        /// <b>System.Generic.Boolean</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Boolean => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 15);
        
        /// <summary>
        /// <b>System.Generic.DateTime</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateTime => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 14);
        
        /// <summary>
        /// <b>System.Generic.FloatingPoint</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY FloatingPoint => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 16);
        
        /// <summary>
        /// <b>System.Generic.Integer</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Integer => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 13);
        
        /// <summary>
        /// <b>System.Generic.String</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY String => new(new("1e3ee840-bc2b-476c-8237-2acd1a839b22"), 12);
    }
    
    public static class History
    {
        /// <summary>
        /// <b>System.History.CodePage</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY CodePage => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 6);
        
        /// <summary>
        /// <b>System.History.DateChanged</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY DateChanged => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 39);
        
        /// <summary>
        /// <b>System.History.DownloadLocation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DownloadLocation => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 10);
        
        /// <summary>
        /// <b>System.History.DownloadSize</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY DownloadSize => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 11);
        
        /// <summary>
        /// <b>System.History.FavoriteIconHash</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY FavoriteIconHash => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 42);
        
        /// <summary>
        /// <b>System.History.FavoriteIconKey</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY FavoriteIconKey => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 12);
        
        /// <summary>
        /// <b>System.History.Flags</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Flags => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 40);
        
        /// <summary>
        /// <b>System.History.IconBits</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY IconBits => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 24);
        
        /// <summary>
        /// <b>System.History.IconDate</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY IconDate => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 27);
        
        /// <summary>
        /// <b>System.History.IsDownload</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsDownload => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 9);
        
        /// <summary>
        /// <b>System.History.IsFavorite</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsFavorite => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 13);
        
        /// <summary>
        /// <b>System.History.IsFeed</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY IsFeed => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 18);
        
        /// <summary>
        /// <b>System.History.IsHistory</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsHistory => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 8);
        
        /// <summary>
        /// <b>System.History.IsOfflineFavorite</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsOfflineFavorite => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 14);
        
        /// <summary>
        /// <b>System.History.IsPinnedFavorite</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsPinnedFavorite => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 15);
        
        /// <summary>
        /// <b>System.History.IsTopLevel</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsTopLevel => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 17);
        
        /// <summary>
        /// <b>System.History.IsTypedUrl</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsTypedUrl => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 16);
        
        /// <summary>
        /// <b>System.History.Keywords</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Keywords => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 19);
        
        /// <summary>
        /// <b>System.History.Points</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Points => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 28);
        
        /// <summary>
        /// <b>System.History.SelectionCount</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY SelectionCount => new(new("1ce0d6bc-536c-4600-b0dd-7e0c66b350d5"), 8);
        
        /// <summary>
        /// <b>System.History.Sessions</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Sessions => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 29);
        
        /// <summary>
        /// <b>System.History.SubscriptionCookie</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY SubscriptionCookie => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 33);
        
        /// <summary>
        /// <b>System.History.TargetUrlHostName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TargetUrlHostName => new(new("1ce0d6bc-536c-4600-b0dd-7e0c66b350d5"), 9);
        
        /// <summary>
        /// <b>System.History.Tracking</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Tracking => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 34);
        
        /// <summary>
        /// <b>System.History.UrlExtraInfo</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UrlExtraInfo => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 5);
        
        /// <summary>
        /// <b>System.History.UrlHash</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY UrlHash => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 1);
        
        /// <summary>
        /// <b>System.History.UserDescription</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UserDescription => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 22);
        
        /// <summary>
        /// <b>System.History.UserKeywords</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UserKeywords => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 20);
        
        /// <summary>
        /// <b>System.History.VisitCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY VisitCount => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 7);
        
        /// <summary>
        /// <b>System.History.Watch</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Watch => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 41);
    }
    
    public static class Home
    {
        /// <summary>
        /// <b>System.Home.DriveType</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DriveType => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 10);
        
        /// <summary>
        /// <b>System.Home.FullName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FullName => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 8);
        
        /// <summary>
        /// <b>System.Home.GraphFileType</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY GraphFileType => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 19);
        
        /// <summary>
        /// <b>System.Home.Grouping</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Grouping => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 2);
        
        /// <summary>
        /// <b>System.Home.IsPinned</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsPinned => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 4);
        
        /// <summary>
        /// <b>System.Home.ItemFolderPathDisplay</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ItemFolderPathDisplay => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 6);
        
        /// <summary>
        /// <b>System.Home.ListId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ListId => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 11);
        
        /// <summary>
        /// <b>System.Home.ListItemId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ListItemId => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 12);
        
        /// <summary>
        /// <b>System.Home.ListItemUniqueId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ListItemUniqueId => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 13);
        
        /// <summary>
        /// <b>System.Home.RecommendationActivityDate</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY RecommendationActivityDate => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 22);
        
        /// <summary>
        /// <b>System.Home.RecommendationProviderSource</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY RecommendationProviderSource => new(new("5ca9b1cb-c69f-404b-abc6-fd336793a6a7"), 22);
        
        /// <summary>
        /// <b>System.Home.RecommendationReasonIcon</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RecommendationReasonIcon => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 21);
        
        /// <summary>
        /// <b>System.Home.Recommended</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Recommended => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 20);
        
        /// <summary>
        /// <b>System.Home.SiteId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SiteId => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 14);
        
        /// <summary>
        /// <b>System.Home.SiteUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SiteUrl => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 15);
        
        /// <summary>
        /// <b>System.Home.SortOrder</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY SortOrder => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 3);
        
        /// <summary>
        /// <b>System.Home.UserOneDriveCid</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UserOneDriveCid => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 18);
        
        /// <summary>
        /// <b>System.Home.WebId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY WebId => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 16);
        
        /// <summary>
        /// <b>System.Home.WebUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY WebUrl => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 9);
        
        public static class PropList
        {
            /// <summary>
            /// <b>System.Home.PropList.Sort</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Sort => new(new("30c8eef4-a832-41e2-ab32-e3c3ca28fd29"), 5);
        }
    }
    
    public static class Identity
    {
        /// <summary>
        /// <b>System.Identity.Blob</b> of <b>VT_BLOB</b> type.
        /// </summary>
        public static PROPERTYKEY Blob => new(new("8c3b93a4-baed-1a83-9a32-102ee313f6eb"), 100);
        
        /// <summary>
        /// <b>System.Identity.DisplayName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DisplayName => new(new("7d683fc9-d155-45a8-bb1f-89d19bcb792f"), 100);
        
        /// <summary>
        /// <b>System.Identity.InternetSid</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY InternetSid => new(new("6d6d5d49-265d-4688-9f4e-1fdd33e7cc83"), 100);
        
        /// <summary>
        /// <b>System.Identity.IsMeIdentity</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsMeIdentity => new(new("a4108708-09df-4377-9dfc-6d99986d5a67"), 100);
        
        /// <summary>
        /// <b>System.Identity.KeyProviderContext</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY KeyProviderContext => new(new("a26f4afc-7346-4299-be47-eb1ae613139f"), 17);
        
        /// <summary>
        /// <b>System.Identity.KeyProviderName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY KeyProviderName => new(new("a26f4afc-7346-4299-be47-eb1ae613139f"), 16);
        
        /// <summary>
        /// <b>System.Identity.LogonStatusString</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LogonStatusString => new(new("f18dedf3-337f-42c0-9e03-cee08708a8c3"), 100);
        
        /// <summary>
        /// <b>System.Identity.PrimaryEmailAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryEmailAddress => new(new("fcc16823-baed-4f24-9b32-a0982117f7fa"), 100);
        
        /// <summary>
        /// <b>System.Identity.PrimarySid</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimarySid => new(new("2b1b801e-c0c1-4987-9ec5-72fa89814787"), 100);
        
        /// <summary>
        /// <b>System.Identity.ProviderData</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProviderData => new(new("a8a74b92-361b-4e9a-b722-7c4a7330a312"), 100);
        
        /// <summary>
        /// <b>System.Identity.ProviderID</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY ProviderID => new(new("74a7de49-fa11-4d3d-a006-db7e08675916"), 100);
        
        /// <summary>
        /// <b>System.Identity.QualifiedUserName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY QualifiedUserName => new(new("da520e51-f4e9-4739-ac82-02e0a95c9030"), 100);
        
        /// <summary>
        /// <b>System.Identity.UniqueID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UniqueID => new(new("e55fc3b0-2b60-4220-918e-b21e8bf16016"), 100);
        
        /// <summary>
        /// <b>System.Identity.UserName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UserName => new(new("c4322503-78ca-49c6-9acc-a68e2afd7b6b"), 100);
    }
    
    public static class IdentityProvider
    {
        /// <summary>
        /// <b>System.IdentityProvider.Name</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Name => new(new("b96eff7b-35ca-4a35-8607-29e3a54c46ea"), 100);
        
        /// <summary>
        /// <b>System.IdentityProvider.Picture</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Picture => new(new("2425166f-5642-4864-992f-98fd98f294c3"), 100);
    }
    
    public static class Image
    {
        /// <summary>
        /// <b>System.Image.BitDepth</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY BitDepth => new(new("6444048f-4c8b-11d1-8b70-080036b11a03"), 7);
        
        /// <summary>
        /// <b>System.Image.ColorSpace</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY ColorSpace => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 40961);
        
        /// <summary>
        /// <b>System.Image.CompressedBitsPerPixel</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY CompressedBitsPerPixel => new(new("364b6fa9-37ab-482a-be2b-ae02f60d4318"), 100);
        
        /// <summary>
        /// <b>System.Image.CompressedBitsPerPixelDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY CompressedBitsPerPixelDenominator => new(new("1f8844e1-24ad-4508-9dfd-5326a415ce02"), 100);
        
        /// <summary>
        /// <b>System.Image.CompressedBitsPerPixelNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY CompressedBitsPerPixelNumerator => new(new("d21a7148-d32c-4624-8900-277210f79c0f"), 100);
        
        /// <summary>
        /// <b>System.Image.Compression</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY Compression => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 259);
        
        /// <summary>
        /// <b>System.Image.CompressionText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CompressionText => new(new("3f08e66f-2f44-4bb9-a682-ac35d2562322"), 100);
        
        /// <summary>
        /// <b>System.Image.Copyright</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Copyright => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 33432);
        
        /// <summary>
        /// <b>System.Image.Dimensions</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Dimensions => new(new("6444048f-4c8b-11d1-8b70-080036b11a03"), 13);
        
        /// <summary>
        /// <b>System.Image.HorizontalResolution</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY HorizontalResolution => new(new("6444048f-4c8b-11d1-8b70-080036b11a03"), 5);
        
        /// <summary>
        /// <b>System.Image.HorizontalSize</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY HorizontalSize => new(new("6444048f-4c8b-11d1-8b70-080036b11a03"), 3);
        
        /// <summary>
        /// <b>System.Image.ImageID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ImageID => new(new("10dabe05-32aa-4c29-bf1a-63e2d220587f"), 100);
        
        /// <summary>
        /// <b>System.Image.PropertyBag</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY PropertyBag => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 40096);
        
        /// <summary>
        /// <b>System.Image.ResolutionUnit</b> of <b>VT_I2</b> type.
        /// </summary>
        public static PROPERTYKEY ResolutionUnit => new(new("19b51fa6-1f92-4a5c-ab48-7df0abd67444"), 100);
        
        /// <summary>
        /// <b>System.Image.VerticalResolution</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY VerticalResolution => new(new("6444048f-4c8b-11d1-8b70-080036b11a03"), 6);
        
        /// <summary>
        /// <b>System.Image.VerticalSize</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY VerticalSize => new(new("6444048f-4c8b-11d1-8b70-080036b11a03"), 4);
    }
    
    public static class ItemCustomState
    {
        /// <summary>
        /// <b>System.ItemCustomState.IconReferences</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY IconReferences => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 123);
        
        /// <summary>
        /// <b>System.ItemCustomState.StateList</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY StateList => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 121);
        
        /// <summary>
        /// <b>System.ItemCustomState.Values</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Values => new(new("fceff153-e839-4cf3-a9e7-ea22832094b8"), 122);
    }
    
    public static class Journal
    {
        /// <summary>
        /// <b>System.Journal.Contacts</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Contacts => new(new("dea7c82c-1d89-4a66-9427-a4e3debabcb1"), 100);
        
        /// <summary>
        /// <b>System.Journal.EntryType</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EntryType => new(new("95beb1fc-326d-4644-b396-cd3ed90e6ddf"), 100);
    }
    
    public static class LOGON
    {
        /// <summary>
        /// <b>System.LOGON.AuthenticationPackage</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AuthenticationPackage => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 201);
        
        /// <summary>
        /// <b>System.LOGON.ClientName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ClientName => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 207);
        
        /// <summary>
        /// <b>System.LOGON.DnsDomainName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DnsDomainName => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 205);
        
        /// <summary>
        /// <b>System.LOGON.LUID</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY LUID => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 200);
        
        /// <summary>
        /// <b>System.LOGON.LogonServer</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LogonServer => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 204);
        
        /// <summary>
        /// <b>System.LOGON.LogonTime</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY LogonTime => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 203);
        
        /// <summary>
        /// <b>System.LOGON.Status</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Status => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 209);
        
        /// <summary>
        /// <b>System.LOGON.TSSession</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY TSSession => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 202);
        
        /// <summary>
        /// <b>System.LOGON.UPN</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UPN => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 206);
        
        /// <summary>
        /// <b>System.LOGON.WinStationName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY WinStationName => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 208);
    }
    
    public static class Launcher
    {
        /// <summary>
        /// <b>System.Launcher.AppState</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY AppState => new(new("0ded77b3-c614-456c-ae5b-285b38d7b01b"), 7);
        
        /// <summary>
        /// <b>System.Launcher.GroupID</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY GroupID => new(new("0ded77b3-c614-456c-ae5b-285b38d7b01b"), 3);
        
        /// <summary>
        /// <b>System.Launcher.Order</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY Order => new(new("0ded77b3-c614-456c-ae5b-285b38d7b01b"), 2);
        
        /// <summary>
        /// <b>System.Launcher.SplashScreenImagePath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SplashScreenImagePath => new(new("0ded77b3-c614-456c-ae5b-285b38d7b01b"), 10);
        
        /// <summary>
        /// <b>System.Launcher.TileSize</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY TileSize => new(new("0ded77b3-c614-456c-ae5b-285b38d7b01b"), 8);
        
        /// <summary>
        /// <b>System.Launcher.ViewID</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ViewID => new(new("0ded77b3-c614-456c-ae5b-285b38d7b01b"), 6);
        
        /// <summary>
        /// <b>System.Launcher.WinStoreCategoryId</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY WinStoreCategoryId => new(new("0ded77b3-c614-456c-ae5b-285b38d7b01b"), 21);
    }
    
    public static class LayoutPattern
    {
        /// <summary>
        /// <b>System.LayoutPattern.ContentViewModeForBrowse</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContentViewModeForBrowse => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 500);
        
        /// <summary>
        /// <b>System.LayoutPattern.ContentViewModeForSearch</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContentViewModeForSearch => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 501);
        
        /// <summary>
        /// <b>System.LayoutPattern.Group</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Group => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 504);
        
        /// <summary>
        /// <b>System.LayoutPattern.PlaceHolder</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY PlaceHolder => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 502);
        
        /// <summary>
        /// <b>System.LayoutPattern.TilesViewMode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TilesViewMode => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 503);
    }
    
    public static class Link
    {
        /// <summary>
        /// <b>System.Link.Arguments</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Arguments => new(new("436f2667-14e2-4feb-b30a-146c53b5b674"), 100);
        
        /// <summary>
        /// <b>System.Link.Comment</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Comment => new(new("b9b4b3fc-2b51-4a42-b5d8-324146afcf25"), 5);
        
        /// <summary>
        /// <b>System.Link.DateVisited</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateVisited => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 23);
        
        /// <summary>
        /// <b>System.Link.Description</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Description => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 21);
        
        /// <summary>
        /// <b>System.Link.FeedItemLocalId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FeedItemLocalId => new(new("8a2f99f9-3c37-465d-a8d7-69777a246d0c"), 2);
        
        /// <summary>
        /// <b>System.Link.HotKey</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY HotKey => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 36);
        
        /// <summary>
        /// <b>System.Link.ShowCmd</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY ShowCmd => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 37);
        
        /// <summary>
        /// <b>System.Link.Status</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY Status => new(new("b9b4b3fc-2b51-4a42-b5d8-324146afcf25"), 3);
        
        /// <summary>
        /// <b>System.Link.TargetExtension</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY TargetExtension => new(new("7a7d76f4-b630-4bd7-95ff-37cc51a975c9"), 2);
        
        /// <summary>
        /// <b>System.Link.TargetParsingPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TargetParsingPath => new(new("b9b4b3fc-2b51-4a42-b5d8-324146afcf25"), 2);
        
        /// <summary>
        /// <b>System.Link.TargetSFGAOFlags</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY TargetSFGAOFlags => new(new("b9b4b3fc-2b51-4a42-b5d8-324146afcf25"), 8);
        
        /// <summary>
        /// <b>System.Link.TargetSFGAOFlagsStrings</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY TargetSFGAOFlagsStrings => new(new("d6942081-d53b-443d-ad47-5e059d9cd27a"), 3);
        
        /// <summary>
        /// <b>System.Link.TargetUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TargetUrl => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 2);
        
        /// <summary>
        /// <b>System.Link.TargetUrlHostName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TargetUrlHostName => new(new("8a2f99f9-3c37-465d-a8d7-69777a246d0c"), 5);
        
        /// <summary>
        /// <b>System.Link.TargetUrlPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TargetUrlPath => new(new("8a2f99f9-3c37-465d-a8d7-69777a246d0c"), 6);
        
        /// <summary>
        /// <b>System.Link.WhatsNew</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY WhatsNew => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 38);
        
        /// <summary>
        /// <b>System.Link.WorkingFolderPath</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY WorkingFolderPath => new(new("5cbf2787-48cf-4208-b90e-ee5e5d420294"), 35);
    }
    
    public static class LzhFolder
    {
        /// <summary>
        /// <b>System.LzhFolder.CRC16</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY CRC16 => new(new("4f289a46-2bbb-4ae8-9eda-e5e034707a71"), 3);
        
        /// <summary>
        /// <b>System.LzhFolder.CompressedSize</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY CompressedSize => new(new("4f289a46-2bbb-4ae8-9eda-e5e034707a71"), 2);
        
        /// <summary>
        /// <b>System.LzhFolder.Method</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Method => new(new("4f289a46-2bbb-4ae8-9eda-e5e034707a71"), 4);
        
        /// <summary>
        /// <b>System.LzhFolder.Ratio</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Ratio => new(new("4f289a46-2bbb-4ae8-9eda-e5e034707a71"), 5);
    }
    
    public static class Media
    {
        /// <summary>
        /// <b>System.Media.AuthorUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AuthorUrl => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 32);
        
        /// <summary>
        /// <b>System.Media.AverageLevel</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY AverageLevel => new(new("09edd5b6-b301-43c5-9990-d00302effd46"), 100);
        
        /// <summary>
        /// <b>System.Media.ClassPrimaryID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ClassPrimaryID => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 13);
        
        /// <summary>
        /// <b>System.Media.ClassSecondaryID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ClassSecondaryID => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 14);
        
        /// <summary>
        /// <b>System.Media.CollectionGroupID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CollectionGroupID => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 24);
        
        /// <summary>
        /// <b>System.Media.CollectionID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CollectionID => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 25);
        
        /// <summary>
        /// <b>System.Media.ContentDistributor</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContentDistributor => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 18);
        
        /// <summary>
        /// <b>System.Media.ContentID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContentID => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 26);
        
        /// <summary>
        /// <b>System.Media.CreatorApplication</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CreatorApplication => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 27);
        
        /// <summary>
        /// <b>System.Media.CreatorApplicationVersion</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CreatorApplicationVersion => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 28);
        
        /// <summary>
        /// <b>System.Media.DVDID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DVDID => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 15);
        
        /// <summary>
        /// <b>System.Media.DateEncoded</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateEncoded => new(new("2e4b640d-5019-46d8-8881-55414cc5caa0"), 100);
        
        /// <summary>
        /// <b>System.Media.DateReleased</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DateReleased => new(new("de41cc29-6971-4290-b472-f59f2e2f31e2"), 100);
        
        /// <summary>
        /// <b>System.Media.DlnaProfileID</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY DlnaProfileID => new(new("cfa31b45-525d-4998-bb44-3f7d81542fa4"), 100);
        
        /// <summary>
        /// <b>System.Media.Duration</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY Duration => new(new("64440490-4c8b-11d1-8b70-080036b11a03"), 3);
        
        /// <summary>
        /// <b>System.Media.EncodedBy</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EncodedBy => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 36);
        
        /// <summary>
        /// <b>System.Media.EncodingSettings</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EncodingSettings => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 37);
        
        /// <summary>
        /// <b>System.Media.EpisodeNumber</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY EpisodeNumber => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 100);
        
        /// <summary>
        /// <b>System.Media.FrameCount</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FrameCount => new(new("6444048f-4c8b-11d1-8b70-080036b11a03"), 12);
        
        /// <summary>
        /// <b>System.Media.MCDI</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MCDI => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 16);
        
        /// <summary>
        /// <b>System.Media.MetadataContentProvider</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MetadataContentProvider => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 17);
        
        /// <summary>
        /// <b>System.Media.Producer</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Producer => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 22);
        
        /// <summary>
        /// <b>System.Media.PromotionUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PromotionUrl => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 33);
        
        /// <summary>
        /// <b>System.Media.ProtectionType</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProtectionType => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 38);
        
        /// <summary>
        /// <b>System.Media.ProviderRating</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProviderRating => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 39);
        
        /// <summary>
        /// <b>System.Media.ProviderStyle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProviderStyle => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 40);
        
        /// <summary>
        /// <b>System.Media.Publisher</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Publisher => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 30);
        
        /// <summary>
        /// <b>System.Media.SeasonNumber</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY SeasonNumber => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 101);
        
        /// <summary>
        /// <b>System.Media.SeriesName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SeriesName => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 42);
        
        /// <summary>
        /// <b>System.Media.Status</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Status => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 7);
        
        /// <summary>
        /// <b>System.Media.SubTitle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SubTitle => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 38);
        
        /// <summary>
        /// <b>System.Media.SubscriptionContentId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SubscriptionContentId => new(new("9aebae7a-9644-487d-a92c-657585ed751a"), 100);
        
        /// <summary>
        /// <b>System.Media.ThumbnailLargePath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ThumbnailLargePath => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 47);
        
        /// <summary>
        /// <b>System.Media.ThumbnailLargeUri</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ThumbnailLargeUri => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 48);
        
        /// <summary>
        /// <b>System.Media.ThumbnailSmallPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ThumbnailSmallPath => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 49);
        
        /// <summary>
        /// <b>System.Media.ThumbnailSmallUri</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ThumbnailSmallUri => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 50);
        
        /// <summary>
        /// <b>System.Media.UniqueFileIdentifier</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UniqueFileIdentifier => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 35);
        
        /// <summary>
        /// <b>System.Media.UserNoAutoInfo</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UserNoAutoInfo => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 41);
        
        /// <summary>
        /// <b>System.Media.UserWebUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UserWebUrl => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 34);
        
        /// <summary>
        /// <b>System.Media.Writer</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Writer => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 23);
        
        /// <summary>
        /// <b>System.Media.Year</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Year => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 5);
    }
    
    public static class Message
    {
        /// <summary>
        /// <b>System.Message.AttachmentContents</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AttachmentContents => new(new("3143bf7c-80a8-4854-8880-e2e40189bdd0"), 100);
        
        /// <summary>
        /// <b>System.Message.AttachmentNames</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY AttachmentNames => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 21);
        
        /// <summary>
        /// <b>System.Message.BccAddress</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY BccAddress => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 2);
        
        /// <summary>
        /// <b>System.Message.BccName</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY BccName => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 3);
        
        /// <summary>
        /// <b>System.Message.CcAddress</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY CcAddress => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 4);
        
        /// <summary>
        /// <b>System.Message.CcName</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY CcName => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 5);
        
        /// <summary>
        /// <b>System.Message.ConversationID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ConversationID => new(new("dc8f80bd-af1e-4289-85b6-3dfc1b493992"), 100);
        
        /// <summary>
        /// <b>System.Message.ConversationIndex</b> of <b>VT_UI1, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY ConversationIndex => new(new("dc8f80bd-af1e-4289-85b6-3dfc1b493992"), 101);
        
        /// <summary>
        /// <b>System.Message.DateReceived</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateReceived => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 20);
        
        /// <summary>
        /// <b>System.Message.DateSent</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateSent => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 19);
        
        /// <summary>
        /// <b>System.Message.Flags</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY Flags => new(new("a82d9ee7-ca67-4312-965e-226bcea85023"), 100);
        
        /// <summary>
        /// <b>System.Message.FromAddress</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY FromAddress => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 13);
        
        /// <summary>
        /// <b>System.Message.FromName</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY FromName => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 14);
        
        /// <summary>
        /// <b>System.Message.HasAttachments</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY HasAttachments => new(new("9c1fcf74-2d97-41ba-b4ae-cb2e3661a6e4"), 8);
        
        /// <summary>
        /// <b>System.Message.IsFwdOrReply</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY IsFwdOrReply => new(new("9a9bc088-4f6d-469e-9919-e705412040f9"), 100);
        
        /// <summary>
        /// <b>System.Message.MessageClass</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MessageClass => new(new("cd9ed458-08ce-418f-a70e-f912c7bb9c5c"), 103);
        
        /// <summary>
        /// <b>System.Message.Participants</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Participants => new(new("1a9ba605-8e7c-4d11-ad7d-a50ada18ba1b"), 2);
        
        /// <summary>
        /// <b>System.Message.ProofInProgress</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY ProofInProgress => new(new("9098f33c-9a7d-48a8-8de5-2e1227a64e91"), 100);
        
        /// <summary>
        /// <b>System.Message.Received</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY Received => new(new("9c1fcf74-2d97-41ba-b4ae-cb2e3661a6e4"), 17);
        
        /// <summary>
        /// <b>System.Message.SenderAddress</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SenderAddress => new(new("0be1c8e7-1981-4676-ae14-fdd78f05a6e7"), 100);
        
        /// <summary>
        /// <b>System.Message.SenderName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SenderName => new(new("0da41cfa-d224-4a18-ae2f-596158db4b3a"), 100);
        
        /// <summary>
        /// <b>System.Message.Store</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Store => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 15);
        
        /// <summary>
        /// <b>System.Message.ToAddress</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY ToAddress => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 16);
        
        /// <summary>
        /// <b>System.Message.ToDoFlags</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY ToDoFlags => new(new("1f856a9f-6900-4aba-9505-2d5f1b4d66cb"), 100);
        
        /// <summary>
        /// <b>System.Message.ToDoTitle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ToDoTitle => new(new("bccc8a3c-8cef-42e5-9b1c-c69079398bc7"), 100);
        
        /// <summary>
        /// <b>System.Message.ToName</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY ToName => new(new("e3e0584c-b788-4a5a-bb20-7f5a44c9acdd"), 17);
        
        /// <summary>
        /// <b>System.Message.Type</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Type => new(new("9c1fcf74-2d97-41ba-b4ae-cb2e3661a6e4"), 13);
    }
    
    public static class MsGraph
    {
        /// <summary>
        /// <b>System.MsGraph.ActivityType</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ActivityType => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 14);
        
        /// <summary>
        /// <b>System.MsGraph.CompositeId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CompositeId => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 2);
        
        /// <summary>
        /// <b>System.MsGraph.DateLastShared</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateLastShared => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 9);
        
        /// <summary>
        /// <b>System.MsGraph.DriveId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DriveId => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 3);
        
        /// <summary>
        /// <b>System.MsGraph.GraphFileType</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY GraphFileType => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 16);
        
        /// <summary>
        /// <b>System.MsGraph.IconUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY IconUrl => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 15);
        
        /// <summary>
        /// <b>System.MsGraph.ItemId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ItemId => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 4);
        
        /// <summary>
        /// <b>System.MsGraph.PrimaryActivityActorDisplayName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryActivityActorDisplayName => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 13);
        
        /// <summary>
        /// <b>System.MsGraph.PrimaryActivityActorUpn</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PrimaryActivityActorUpn => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 12);
        
        /// <summary>
        /// <b>System.MsGraph.RecommendationReason</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RecommendationReason => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 8);
        
        /// <summary>
        /// <b>System.MsGraph.RecommendationReferenceId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RecommendationReferenceId => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 5);
        
        /// <summary>
        /// <b>System.MsGraph.RecommendationResultSourceId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RecommendationResultSourceId => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 7);
        
        /// <summary>
        /// <b>System.MsGraph.SharedByEmail</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SharedByEmail => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 11);
        
        /// <summary>
        /// <b>System.MsGraph.SharedByName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SharedByName => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 10);
        
        /// <summary>
        /// <b>System.MsGraph.WebAccountId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY WebAccountId => new(new("4f85567e-fff0-4df5-b1d9-98b314ff0729"), 6);
    }
    
    public static class Music
    {
        /// <summary>
        /// <b>System.Music.AlbumArtist</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AlbumArtist => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 13);
        
        /// <summary>
        /// <b>System.Music.AlbumArtistSortOverride</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AlbumArtistSortOverride => new(new("f1fdb4af-f78c-466c-bb05-56e92db0b8ec"), 103);
        
        /// <summary>
        /// <b>System.Music.AlbumID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AlbumID => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 100);
        
        /// <summary>
        /// <b>System.Music.AlbumTitle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AlbumTitle => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 4);
        
        /// <summary>
        /// <b>System.Music.AlbumTitleSortOverride</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AlbumTitleSortOverride => new(new("13eb7ffc-ec89-4346-b19d-ccc6f1784223"), 101);
        
        /// <summary>
        /// <b>System.Music.Artist</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Artist => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 2);
        
        /// <summary>
        /// <b>System.Music.ArtistSortOverride</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY ArtistSortOverride => new(new("deeb2db5-0696-4ce0-94fe-a01f77a45fb5"), 102);
        
        /// <summary>
        /// <b>System.Music.BeatsPerMinute</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BeatsPerMinute => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 35);
        
        /// <summary>
        /// <b>System.Music.Composer</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Composer => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 19);
        
        /// <summary>
        /// <b>System.Music.ComposerSortOverride</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY ComposerSortOverride => new(new("00bc20a3-bd48-4085-872c-a88d77f5097e"), 105);
        
        /// <summary>
        /// <b>System.Music.Conductor</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Conductor => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 36);
        
        /// <summary>
        /// <b>System.Music.ContentGroupDescription</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContentGroupDescription => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 33);
        
        /// <summary>
        /// <b>System.Music.DiscNumber</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DiscNumber => new(new("6afe7437-9bcd-49c7-80fe-4a5c65fa5874"), 104);
        
        /// <summary>
        /// <b>System.Music.DisplayArtist</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DisplayArtist => new(new("fd122953-fa93-4ef7-92c3-04c946b2f7c8"), 100);
        
        /// <summary>
        /// <b>System.Music.Genre</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Genre => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 11);
        
        /// <summary>
        /// <b>System.Music.InitialKey</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY InitialKey => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 34);
        
        /// <summary>
        /// <b>System.Music.IsCompilation</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsCompilation => new(new("c449d5cb-9ea4-4809-82e8-af9d59ded6d1"), 100);
        
        /// <summary>
        /// <b>System.Music.Lyrics</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Lyrics => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 12);
        
        /// <summary>
        /// <b>System.Music.Mood</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Mood => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 39);
        
        /// <summary>
        /// <b>System.Music.PartOfSet</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PartOfSet => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 37);
        
        /// <summary>
        /// <b>System.Music.Period</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Period => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 31);
        
        /// <summary>
        /// <b>System.Music.SynchronizedLyrics</b> of <b>VT_BLOB</b> type.
        /// </summary>
        public static PROPERTYKEY SynchronizedLyrics => new(new("6b223b6a-162e-4aa9-b39f-05d678fc6d77"), 100);
        
        /// <summary>
        /// <b>System.Music.TrackNumber</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY TrackNumber => new(new("56a3372e-ce9c-11d2-9f0e-006097c686f6"), 7);
    }
    
    public static class Namespace
    {
        /// <summary>
        /// <b>System.Namespace.GroupId</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY GroupId => new(new("ed2631a8-795c-4496-b5d6-c480b170710b"), 2);
    }
    
    public static class Note
    {
        /// <summary>
        /// <b>System.Note.Color</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY Color => new(new("4776cafa-bce4-4cb1-a23e-265e76d8eb11"), 100);
        
        /// <summary>
        /// <b>System.Note.ColorText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ColorText => new(new("46b4e8de-cdb2-440d-885c-1658eb65b914"), 100);
    }
    
    public static class OfflineFiles
    {
        /// <summary>
        /// <b>System.OfflineFiles.CreatedOffline</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY CreatedOffline => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 1);
        
        /// <summary>
        /// <b>System.OfflineFiles.DeletedOffline</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY DeletedOffline => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 2);
        
        /// <summary>
        /// <b>System.OfflineFiles.Encrypted</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Encrypted => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 15);
        
        /// <summary>
        /// <b>System.OfflineFiles.Modified</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Modified => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 3);
        
        /// <summary>
        /// <b>System.OfflineFiles.ModifiedAttributes</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY ModifiedAttributes => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 4);
        
        /// <summary>
        /// <b>System.OfflineFiles.ModifiedData</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY ModifiedData => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 5);
        
        /// <summary>
        /// <b>System.OfflineFiles.ModifiedTime</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY ModifiedTime => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 6);
        
        /// <summary>
        /// <b>System.OfflineFiles.Pinned</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Pinned => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 7);
        
        /// <summary>
        /// <b>System.OfflineFiles.PinnedForComputer</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY PinnedForComputer => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 8);
        
        /// <summary>
        /// <b>System.OfflineFiles.PinnedForFolderRedirection</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY PinnedForFolderRedirection => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 9);
        
        /// <summary>
        /// <b>System.OfflineFiles.PinnedForUser</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY PinnedForUser => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 10);
        
        /// <summary>
        /// <b>System.OfflineFiles.PinnedForUserByPolicy</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY PinnedForUserByPolicy => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 11);
        
        /// <summary>
        /// <b>System.OfflineFiles.Sparse</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Sparse => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 12);
        
        /// <summary>
        /// <b>System.OfflineFiles.Suspended</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Suspended => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 13);
        
        /// <summary>
        /// <b>System.OfflineFiles.SuspendedRoot</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY SuspendedRoot => new(new("25a73d40-cbba-46f7-980d-b346cc767a4c"), 14);
    }
    
    public static class PROFILE
    {
        /// <summary>
        /// <b>System.PROFILE.GUID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY GUID => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 501);
        
        /// <summary>
        /// <b>System.PROFILE.Path</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Path => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 500);
    }
    
    public static class Photo
    {
        /// <summary>
        /// <b>System.Photo.Aperture</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY Aperture => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 37378);
        
        /// <summary>
        /// <b>System.Photo.ApertureDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ApertureDenominator => new(new("e1a9a38b-6685-46bd-875e-570dc7ad7320"), 100);
        
        /// <summary>
        /// <b>System.Photo.ApertureNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ApertureNumerator => new(new("0337ecec-39fb-4581-a0bd-4c4cc51e9914"), 100);
        
        /// <summary>
        /// <b>System.Photo.Brightness</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY Brightness => new(new("1a701bf6-478c-4361-83ab-3701bb053c58"), 100);
        
        /// <summary>
        /// <b>System.Photo.BrightnessDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY BrightnessDenominator => new(new("6ebe6946-2321-440a-90f0-c043efd32476"), 100);
        
        /// <summary>
        /// <b>System.Photo.BrightnessNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY BrightnessNumerator => new(new("9e7d118f-b314-45a0-8cfb-d654b917c9e9"), 100);
        
        /// <summary>
        /// <b>System.Photo.CameraManufacturer</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CameraManufacturer => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 271);
        
        /// <summary>
        /// <b>System.Photo.CameraModel</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CameraModel => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 272);
        
        /// <summary>
        /// <b>System.Photo.CameraSerialNumber</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CameraSerialNumber => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 273);
        
        /// <summary>
        /// <b>System.Photo.Contrast</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Contrast => new(new("2a785ba9-8d23-4ded-82e6-60a350c86a10"), 100);
        
        /// <summary>
        /// <b>System.Photo.ContrastText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContrastText => new(new("59dde9f2-5253-40ea-9a8b-479e96c6249a"), 100);
        
        /// <summary>
        /// <b>System.Photo.DateTaken</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateTaken => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 36867);
        
        /// <summary>
        /// <b>System.Photo.DigitalZoom</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY DigitalZoom => new(new("f85bf840-a925-4bc2-b0c4-8e36b598679e"), 100);
        
        /// <summary>
        /// <b>System.Photo.DigitalZoomDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DigitalZoomDenominator => new(new("745baf0e-e5c1-4cfb-8a1b-d031a0a52393"), 100);
        
        /// <summary>
        /// <b>System.Photo.DigitalZoomNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY DigitalZoomNumerator => new(new("16cbb924-6500-473b-a5be-f1599bcbe413"), 100);
        
        /// <summary>
        /// <b>System.Photo.EXIFVersion</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EXIFVersion => new(new("d35f743a-eb2e-47f2-a286-844132cb1427"), 100);
        
        /// <summary>
        /// <b>System.Photo.Event</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Event => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 18248);
        
        /// <summary>
        /// <b>System.Photo.ExposureBias</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureBias => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 37380);
        
        /// <summary>
        /// <b>System.Photo.ExposureBiasDenominator</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureBiasDenominator => new(new("ab205e50-04b7-461c-a18c-2f233836e627"), 100);
        
        /// <summary>
        /// <b>System.Photo.ExposureBiasNumerator</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureBiasNumerator => new(new("738bf284-1d87-420b-92cf-5834bf6ef9ed"), 100);
        
        /// <summary>
        /// <b>System.Photo.ExposureIndex</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureIndex => new(new("967b5af8-995a-46ed-9e11-35b3c5b9782d"), 100);
        
        /// <summary>
        /// <b>System.Photo.ExposureIndexDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureIndexDenominator => new(new("93112f89-c28b-492f-8a9d-4be2062cee8a"), 100);
        
        /// <summary>
        /// <b>System.Photo.ExposureIndexNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureIndexNumerator => new(new("cdedcf30-8919-44df-8f4c-4eb2ffdb8d89"), 100);
        
        /// <summary>
        /// <b>System.Photo.ExposureProgram</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureProgram => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 34850);
        
        /// <summary>
        /// <b>System.Photo.ExposureProgramText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureProgramText => new(new("fec690b7-5f30-4646-ae47-4caafba884a3"), 100);
        
        /// <summary>
        /// <b>System.Photo.ExposureTime</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureTime => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 33434);
        
        /// <summary>
        /// <b>System.Photo.ExposureTimeDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureTimeDenominator => new(new("55e98597-ad16-42e0-b624-21599a199838"), 100);
        
        /// <summary>
        /// <b>System.Photo.ExposureTimeNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ExposureTimeNumerator => new(new("257e44e2-9031-4323-ac38-85c552871b2e"), 100);
        
        /// <summary>
        /// <b>System.Photo.FNumber</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY FNumber => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 33437);
        
        /// <summary>
        /// <b>System.Photo.FNumberDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FNumberDenominator => new(new("e92a2496-223b-4463-a4e3-30eabba79d80"), 100);
        
        /// <summary>
        /// <b>System.Photo.FNumberNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FNumberNumerator => new(new("1b97738a-fdfc-462f-9d93-1957e08be90c"), 100);
        
        /// <summary>
        /// <b>System.Photo.Flash</b> of <b>VT_UI1</b> type.
        /// </summary>
        public static PROPERTYKEY Flash => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 37385);
        
        /// <summary>
        /// <b>System.Photo.FlashEnergy</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY FlashEnergy => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 41483);
        
        /// <summary>
        /// <b>System.Photo.FlashEnergyDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FlashEnergyDenominator => new(new("d7b61c70-6323-49cd-a5fc-c84277162c97"), 100);
        
        /// <summary>
        /// <b>System.Photo.FlashEnergyNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FlashEnergyNumerator => new(new("fcad3d3d-0858-400f-aaa3-2f66cce2a6bc"), 100);
        
        /// <summary>
        /// <b>System.Photo.FlashFired</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY FlashFired => new(new("2d152b40-ca39-40db-b2cc-573725b2fec5"), 100);
        
        /// <summary>
        /// <b>System.Photo.FlashManufacturer</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FlashManufacturer => new(new("aabaf6c9-e0c5-4719-8585-57b103e584fe"), 100);
        
        /// <summary>
        /// <b>System.Photo.FlashModel</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FlashModel => new(new("fe83bb35-4d1a-42e2-916b-06f3e1af719e"), 100);
        
        /// <summary>
        /// <b>System.Photo.FlashText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FlashText => new(new("6b8b68f6-200b-47ea-8d25-d8050f57339f"), 100);
        
        /// <summary>
        /// <b>System.Photo.FocalLength</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY FocalLength => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 37386);
        
        /// <summary>
        /// <b>System.Photo.FocalLengthDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FocalLengthDenominator => new(new("305bc615-dca1-44a5-9fd4-10c0ba79412e"), 100);
        
        /// <summary>
        /// <b>System.Photo.FocalLengthInFilm</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY FocalLengthInFilm => new(new("a0e74609-b84d-4f49-b860-462bd9971f98"), 100);
        
        /// <summary>
        /// <b>System.Photo.FocalLengthNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FocalLengthNumerator => new(new("776b6b3b-1e3d-4b0c-9a0e-8fbaf2a8492a"), 100);
        
        /// <summary>
        /// <b>System.Photo.FocalPlaneXResolution</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY FocalPlaneXResolution => new(new("cfc08d97-c6f7-4484-89dd-ebef4356fe76"), 100);
        
        /// <summary>
        /// <b>System.Photo.FocalPlaneXResolutionDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FocalPlaneXResolutionDenominator => new(new("0933f3f5-4786-4f46-a8e8-d64dd37fa521"), 100);
        
        /// <summary>
        /// <b>System.Photo.FocalPlaneXResolutionNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FocalPlaneXResolutionNumerator => new(new("dccb10af-b4e2-4b88-95f9-031b4d5ab490"), 100);
        
        /// <summary>
        /// <b>System.Photo.FocalPlaneYResolution</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY FocalPlaneYResolution => new(new("4fffe4d0-914f-4ac4-8d6f-c9c61de169b1"), 100);
        
        /// <summary>
        /// <b>System.Photo.FocalPlaneYResolutionDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FocalPlaneYResolutionDenominator => new(new("1d6179a6-a876-4031-b013-3347b2b64dc8"), 100);
        
        /// <summary>
        /// <b>System.Photo.FocalPlaneYResolutionNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FocalPlaneYResolutionNumerator => new(new("a2e541c5-4440-4ba8-867e-75cfc06828cd"), 100);
        
        /// <summary>
        /// <b>System.Photo.GainControl</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY GainControl => new(new("fa304789-00c7-4d80-904a-1e4dcc7265aa"), 100);
        
        /// <summary>
        /// <b>System.Photo.GainControlDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY GainControlDenominator => new(new("42864dfd-9da4-4f77-bded-4aad7b256735"), 100);
        
        /// <summary>
        /// <b>System.Photo.GainControlNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY GainControlNumerator => new(new("8e8ecf7c-b7b8-4eb8-a63f-0ee715c96f9e"), 100);
        
        /// <summary>
        /// <b>System.Photo.GainControlText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY GainControlText => new(new("c06238b2-0bf9-4279-a723-25856715cb9d"), 100);
        
        /// <summary>
        /// <b>System.Photo.ISOSpeed</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY ISOSpeed => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 34855);
        
        /// <summary>
        /// <b>System.Photo.LensManufacturer</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LensManufacturer => new(new("e6ddcaf7-29c5-4f0a-9a68-d19412ec7090"), 100);
        
        /// <summary>
        /// <b>System.Photo.LensModel</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LensModel => new(new("e1277516-2b5f-4869-89b1-2e585bd38b7a"), 100);
        
        /// <summary>
        /// <b>System.Photo.LightSource</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY LightSource => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 37384);
        
        /// <summary>
        /// <b>System.Photo.MakerNote</b> of <b>VT_UI1, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY MakerNote => new(new("fa303353-b659-4052-85e9-bcac79549b84"), 100);
        
        /// <summary>
        /// <b>System.Photo.MakerNoteOffset</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY MakerNoteOffset => new(new("813f4124-34e6-4d17-ab3e-6b1f3c2247a1"), 100);
        
        /// <summary>
        /// <b>System.Photo.MaxAperture</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY MaxAperture => new(new("08f6d7c2-e3f2-44fc-af1e-5aa5c81a2d3e"), 100);
        
        /// <summary>
        /// <b>System.Photo.MaxApertureDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY MaxApertureDenominator => new(new("c77724d4-601f-46c5-9b89-c53f93bceb77"), 100);
        
        /// <summary>
        /// <b>System.Photo.MaxApertureNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY MaxApertureNumerator => new(new("c107e191-a459-44c5-9ae6-b952ad4b906d"), 100);
        
        /// <summary>
        /// <b>System.Photo.MeteringMode</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY MeteringMode => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 37383);
        
        /// <summary>
        /// <b>System.Photo.MeteringModeText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY MeteringModeText => new(new("f628fd8c-7ba8-465a-a65b-c5aa79263a9e"), 100);
        
        /// <summary>
        /// <b>System.Photo.Orientation</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY Orientation => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 274);
        
        /// <summary>
        /// <b>System.Photo.OrientationText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OrientationText => new(new("a9ea193c-c511-498a-a06b-58e2776dcc28"), 100);
        
        /// <summary>
        /// <b>System.Photo.PeopleNames</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY PeopleNames => new(new("e8309b6e-084c-49b4-b1fc-90a80331b638"), 100);
        
        /// <summary>
        /// <b>System.Photo.PhotometricInterpretation</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY PhotometricInterpretation => new(new("341796f1-1df9-4b1c-a564-91bdefa43877"), 100);
        
        /// <summary>
        /// <b>System.Photo.PhotometricInterpretationText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PhotometricInterpretationText => new(new("821437d6-9eab-4765-a589-3b1cbbd22a61"), 100);
        
        /// <summary>
        /// <b>System.Photo.ProgramMode</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ProgramMode => new(new("6d217f6d-3f6a-4825-b470-5f03ca2fbe9b"), 100);
        
        /// <summary>
        /// <b>System.Photo.ProgramModeText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProgramModeText => new(new("7fe3aa27-2648-42f3-89b0-454e5cb150c3"), 100);
        
        /// <summary>
        /// <b>System.Photo.RelatedSoundFile</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RelatedSoundFile => new(new("318a6b45-087f-4dc2-b8cc-05359551fc9e"), 100);
        
        /// <summary>
        /// <b>System.Photo.Saturation</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Saturation => new(new("49237325-a95a-4f67-b211-816b2d45d2e0"), 100);
        
        /// <summary>
        /// <b>System.Photo.SaturationText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SaturationText => new(new("61478c08-b600-4a84-bbe4-e99c45f0a072"), 100);
        
        /// <summary>
        /// <b>System.Photo.Sharpness</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Sharpness => new(new("fc6976db-8349-4970-ae97-b3c5316a08f0"), 100);
        
        /// <summary>
        /// <b>System.Photo.SharpnessText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SharpnessText => new(new("51ec3f47-dd50-421d-8769-334f50424b1e"), 100);
        
        /// <summary>
        /// <b>System.Photo.ShutterSpeed</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY ShutterSpeed => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 37377);
        
        /// <summary>
        /// <b>System.Photo.ShutterSpeedDenominator</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY ShutterSpeedDenominator => new(new("e13d8975-81c7-4948-ae3f-37cae11e8ff7"), 100);
        
        /// <summary>
        /// <b>System.Photo.ShutterSpeedNumerator</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY ShutterSpeedNumerator => new(new("16ea4042-d6f4-4bca-8349-7c78d30fb333"), 100);
        
        /// <summary>
        /// <b>System.Photo.SubjectDistance</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY SubjectDistance => new(new("14b81da1-0135-4d31-96d9-6cbfc9671a99"), 37382);
        
        /// <summary>
        /// <b>System.Photo.SubjectDistanceDenominator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY SubjectDistanceDenominator => new(new("0c840a88-b043-466d-9766-d4b26da3fa77"), 100);
        
        /// <summary>
        /// <b>System.Photo.SubjectDistanceNumerator</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY SubjectDistanceNumerator => new(new("8af4961c-f526-43e5-aa81-db768219178d"), 100);
        
        /// <summary>
        /// <b>System.Photo.TagViewAggregate</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY TagViewAggregate => new(new("b812f15d-c2d8-4bbf-bacd-79744346113f"), 100);
        
        /// <summary>
        /// <b>System.Photo.TranscodedForSync</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY TranscodedForSync => new(new("9a8ebb75-6458-4e82-bacb-35c0095b03bb"), 100);
        
        /// <summary>
        /// <b>System.Photo.WhiteBalance</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY WhiteBalance => new(new("ee3d3d8a-5381-4cfa-b13b-aaf66b5f4ec9"), 100);
        
        /// <summary>
        /// <b>System.Photo.WhiteBalanceText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY WhiteBalanceText => new(new("6336b95e-c7a7-426d-86fd-7ae3d39c84b4"), 100);
    }
    
    public static class PrintStatus
    {
        /// <summary>
        /// <b>System.PrintStatus.Comment</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Comment => new(new("dd141766-313a-4a30-90f0-056a7c968437"), 5);
        
        /// <summary>
        /// <b>System.PrintStatus.DocumentCount</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DocumentCount => new(new("dd141766-313a-4a30-90f0-056a7c968437"), 2);
        
        /// <summary>
        /// <b>System.PrintStatus.ErrorStatus</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ErrorStatus => new(new("dd141766-313a-4a30-90f0-056a7c968437"), 3);
        
        /// <summary>
        /// <b>System.PrintStatus.InfoStatus</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY InfoStatus => new(new("dd141766-313a-4a30-90f0-056a7c968437"), 8);
        
        /// <summary>
        /// <b>System.PrintStatus.Location</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Location => new(new("dd141766-313a-4a30-90f0-056a7c968437"), 4);
        
        /// <summary>
        /// <b>System.PrintStatus.Preferences</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Preferences => new(new("dd141766-313a-4a30-90f0-056a7c968437"), 6);
        
        /// <summary>
        /// <b>System.PrintStatus.WarningStatus</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY WarningStatus => new(new("dd141766-313a-4a30-90f0-056a7c968437"), 7);
    }
    
    public static class Printer
    {
        /// <summary>
        /// <b>System.Printer.Default</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Default => new(new("fe9e4c12-aacb-4aa3-966d-91a29e6128b5"), 3);
        
        /// <summary>
        /// <b>System.Printer.Location</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Location => new(new("fe9e4c12-aacb-4aa3-966d-91a29e6128b5"), 4);
        
        /// <summary>
        /// <b>System.Printer.Model</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Model => new(new("fe9e4c12-aacb-4aa3-966d-91a29e6128b5"), 5);
        
        /// <summary>
        /// <b>System.Printer.QueueSize</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY QueueSize => new(new("fe9e4c12-aacb-4aa3-966d-91a29e6128b5"), 6);
        
        /// <summary>
        /// <b>System.Printer.Status</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Status => new(new("fe9e4c12-aacb-4aa3-966d-91a29e6128b5"), 7);
    }
    
    public static class PropGroup
    {
        /// <summary>
        /// <b>System.PropGroup.Advanced</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Advanced => new(new("900a403b-097b-4b95-8ae2-071fdaeeb118"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Audio</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Audio => new(new("2804d469-788f-48aa-8570-71b9c187e138"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Calendar</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Calendar => new(new("9973d2b5-bfd8-438a-ba94-5349b293181a"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Camera</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Camera => new(new("de00de32-547e-4981-ad4b-542f2e9007d8"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Contact</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Contact => new(new("df975fd3-250a-4004-858f-34e29a3e37aa"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Content</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Content => new(new("d0dab0ba-368a-4050-a882-6c010fd19a4f"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Description</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Description => new(new("8969b275-9475-4e00-a887-ff93b8b41e44"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.FileSystem</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY FileSystem => new(new("e3a7d2c1-80fc-4b40-8f34-30ea111bdc2e"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.GPS</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY GPS => new(new("f3713ada-90e3-4e11-aae5-fdc17685b9be"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.General</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY General => new(new("cc301630-b192-4c22-b372-9f4c6d338e07"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Image</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Image => new(new("e3690a87-0fa8-4a2a-9a9f-fce8827055ac"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Media</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Media => new(new("61872cf7-6b5e-4b4b-ac2d-59da84459248"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.MediaAdvanced</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY MediaAdvanced => new(new("8859a284-de7e-4642-99ba-d431d044b1ec"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Message</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Message => new(new("7fd7259d-16b4-4135-9f97-7c96ecd2fa9e"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Music</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Music => new(new("68dd6094-7216-40f1-a029-43fe7127043f"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Origin</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Origin => new(new("2598d2fb-5569-4367-95df-5cd3a177e1a5"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.PhotoAdvanced</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY PhotoAdvanced => new(new("0cb2bf5a-9ee7-4a86-8222-f01e07fdadaf"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.RecordedTV</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY RecordedTV => new(new("e7b33238-6584-4170-a5c0-ac25efd9da56"), 100);
        
        /// <summary>
        /// <b>System.PropGroup.Video</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Video => new(new("bebe0920-7671-4c54-a3eb-49fddfc191ee"), 100);
    }
    
    public static class PropList
    {
        /// <summary>
        /// <b>System.PropList.ConflictPrompt</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ConflictPrompt => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 11);
        
        /// <summary>
        /// <b>System.PropList.ContentViewModeForBrowse</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContentViewModeForBrowse => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 13);
        
        /// <summary>
        /// <b>System.PropList.ContentViewModeForSearch</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContentViewModeForSearch => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 14);
        
        /// <summary>
        /// <b>System.PropList.DetailsPaneNullSelect</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DetailsPaneNullSelect => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 510);
        
        /// <summary>
        /// <b>System.PropList.DetailsPaneNullSelectTitle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DetailsPaneNullSelectTitle => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 511);
        
        /// <summary>
        /// <b>System.PropList.ExtendedTileInfo</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ExtendedTileInfo => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 9);
        
        /// <summary>
        /// <b>System.PropList.FileOperationPrompt</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FileOperationPrompt => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 10);
        
        /// <summary>
        /// <b>System.PropList.FullDetails</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FullDetails => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 2);
        
        /// <summary>
        /// <b>System.PropList.InfoTip</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY InfoTip => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 4);
        
        /// <summary>
        /// <b>System.PropList.NonPersonal</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY NonPersonal => new(new("49d1091f-082e-493f-b23f-d2308aa9668c"), 100);
        
        /// <summary>
        /// <b>System.PropList.PreviewDetails</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PreviewDetails => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 8);
        
        /// <summary>
        /// <b>System.PropList.PreviewTitle</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PreviewTitle => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 6);
        
        /// <summary>
        /// <b>System.PropList.QuickTip</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY QuickTip => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 5);
        
        /// <summary>
        /// <b>System.PropList.SetDefaultsFor</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SetDefaultsFor => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 12);
        
        /// <summary>
        /// <b>System.PropList.StatusBar</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY StatusBar => new(new("26dc287c-6e3d-4bd3-b2b0-6a26ba2e346d"), 4);
        
        /// <summary>
        /// <b>System.PropList.StatusIcons</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY StatusIcons => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 16);
        
        /// <summary>
        /// <b>System.PropList.StatusIconsDisplayFlag</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY StatusIconsDisplayFlag => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 18);
        
        /// <summary>
        /// <b>System.PropList.TileInfo</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TileInfo => new(new("c9944a21-a406-48fe-8225-aec7e24c211b"), 3);
        
        /// <summary>
        /// <b>System.PropList.XPDetailsPanel</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY XPDetailsPanel => new(new("f2275480-f782-4291-bd94-f13693513aec"), 0);
    }
    
    public static class Protocol
    {
        /// <summary>
        /// <b>System.Protocol.Name</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Name => new(new("d9c22960-532c-4bc6-9876-7b12b52593d7"), 2);
    }
    
    public static class RecordedTV
    {
        /// <summary>
        /// <b>System.RecordedTV.ChannelNumber</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ChannelNumber => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 7);
        
        /// <summary>
        /// <b>System.RecordedTV.Credits</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Credits => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 4);
        
        /// <summary>
        /// <b>System.RecordedTV.DateContentExpires</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateContentExpires => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 15);
        
        /// <summary>
        /// <b>System.RecordedTV.EpisodeName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EpisodeName => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 2);
        
        /// <summary>
        /// <b>System.RecordedTV.IsATSCContent</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsATSCContent => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 16);
        
        /// <summary>
        /// <b>System.RecordedTV.IsClosedCaptioningAvailable</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsClosedCaptioningAvailable => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 12);
        
        /// <summary>
        /// <b>System.RecordedTV.IsDTVContent</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsDTVContent => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 17);
        
        /// <summary>
        /// <b>System.RecordedTV.IsHDContent</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsHDContent => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 18);
        
        /// <summary>
        /// <b>System.RecordedTV.IsRepeatBroadcast</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsRepeatBroadcast => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 13);
        
        /// <summary>
        /// <b>System.RecordedTV.IsSAP</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsSAP => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 14);
        
        /// <summary>
        /// <b>System.RecordedTV.NetworkAffiliation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY NetworkAffiliation => new(new("2c53c813-fb63-4e22-a1ab-0b331ca1e273"), 100);
        
        /// <summary>
        /// <b>System.RecordedTV.OriginalBroadcastDate</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY OriginalBroadcastDate => new(new("4684fe97-8765-4842-9c13-f006447b178c"), 100);
        
        /// <summary>
        /// <b>System.RecordedTV.ProgramDescription</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProgramDescription => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 3);
        
        /// <summary>
        /// <b>System.RecordedTV.RecordingTime</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY RecordingTime => new(new("a5477f61-7a82-4eca-9dde-98b69b2479b3"), 100);
        
        /// <summary>
        /// <b>System.RecordedTV.StationCallSign</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY StationCallSign => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 5);
        
        /// <summary>
        /// <b>System.RecordedTV.StationName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY StationName => new(new("1b5439e7-eba1-4af8-bdd7-7af1d4549493"), 100);
        
        /// <summary>
        /// <b>System.RecordedTV.VideoQuality</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY VideoQuality => new(new("6d748de2-8d38-4cc3-ac60-f009b057c557"), 10);
    }
    
    public static class Recycle
    {
        /// <summary>
        /// <b>System.Recycle.DateDeleted</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateDeleted => new(new("9b174b33-40ff-11d2-a27e-00c04fc30871"), 3);
        
        /// <summary>
        /// <b>System.Recycle.DeletedFrom</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DeletedFrom => new(new("9b174b33-40ff-11d2-a27e-00c04fc30871"), 2);
    }
    
    public static class SAM
    {
        /// <summary>
        /// <b>System.SAM.AccountIsDisabledForLogonUI</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY AccountIsDisabledForLogonUI => new(new("8bf6b9f6-b4f5-482f-a2c2-44bdad2fcfa9"), 51);
        
        /// <summary>
        /// <b>System.SAM.AccountName</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY AccountName => new(new("9c1fcf74-2d97-41ba-b4ae-cb2e3661a6e4"), 10);
        
        /// <summary>
        /// <b>System.SAM.AdminComment</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AdminComment => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 14);
        
        /// <summary>
        /// <b>System.SAM.AllowedLogon</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY AllowedLogon => new(new("7ba3535d-69aa-4525-a938-f3ec79485377"), 2);
        
        /// <summary>
        /// <b>System.SAM.BatchLogin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY BatchLogin => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 38);
        
        /// <summary>
        /// <b>System.SAM.CodePage</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY CodePage => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 22);
        
        /// <summary>
        /// <b>System.SAM.CountryCode</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY CountryCode => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 21);
        
        /// <summary>
        /// <b>System.SAM.DateAccountExpires</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateAccountExpires => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 6);
        
        /// <summary>
        /// <b>System.SAM.DateChanged</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateChanged => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 4);
        
        /// <summary>
        /// <b>System.SAM.DenyBatchLogin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY DenyBatchLogin => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 43);
        
        /// <summary>
        /// <b>System.SAM.DenyInteractiveLogin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY DenyInteractiveLogin => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 41);
        
        /// <summary>
        /// <b>System.SAM.DenyNetworkLogin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY DenyNetworkLogin => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 42);
        
        /// <summary>
        /// <b>System.SAM.DenyRemoteInteractiveLogin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY DenyRemoteInteractiveLogin => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 45);
        
        /// <summary>
        /// <b>System.SAM.DenyServiceLogin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY DenyServiceLogin => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 44);
        
        /// <summary>
        /// <b>System.SAM.Domain</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Domain => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 26);
        
        /// <summary>
        /// <b>System.SAM.DontEnumerateForLogon</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY DontEnumerateForLogon => new(new("7ba3535d-69aa-4525-a938-f3ec79485377"), 3);
        
        /// <summary>
        /// <b>System.SAM.DontShowInLogonUI</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY DontShowInLogonUI => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 46);
        
        /// <summary>
        /// <b>System.SAM.FullName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FullName => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 9);
        
        /// <summary>
        /// <b>System.SAM.GroupMembers</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY GroupMembers => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 102);
        
        /// <summary>
        /// <b>System.SAM.Groups</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Groups => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 31);
        
        /// <summary>
        /// <b>System.SAM.HomeDirectory</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeDirectory => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 10);
        
        /// <summary>
        /// <b>System.SAM.HomeDirectoryDrive</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HomeDirectoryDrive => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 11);
        
        /// <summary>
        /// <b>System.SAM.InteractiveLogin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY InteractiveLogin => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 36);
        
        /// <summary>
        /// <b>System.SAM.LogonHours</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY LogonHours => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 20);
        
        /// <summary>
        /// <b>System.SAM.Name</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Name => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 2);
        
        /// <summary>
        /// <b>System.SAM.NetworkLogin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY NetworkLogin => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 37);
        
        /// <summary>
        /// <b>System.SAM.Password</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Password => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 17);
        
        /// <summary>
        /// <b>System.SAM.PasswordCanChange</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY PasswordCanChange => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 7);
        
        /// <summary>
        /// <b>System.SAM.PasswordExpired</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY PasswordExpired => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 23);
        
        /// <summary>
        /// <b>System.SAM.PasswordHint</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PasswordHint => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 25);
        
        /// <summary>
        /// <b>System.SAM.PasswordIsEmpty</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY PasswordIsEmpty => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 50);
        
        /// <summary>
        /// <b>System.SAM.PasswordLastSet</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY PasswordLastSet => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 5);
        
        /// <summary>
        /// <b>System.SAM.PasswordMustChange</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY PasswordMustChange => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 8);
        
        /// <summary>
        /// <b>System.SAM.ProfilePath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProfilePath => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 13);
        
        /// <summary>
        /// <b>System.SAM.RemoteInteractiveLogin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY RemoteInteractiveLogin => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 40);
        
        /// <summary>
        /// <b>System.SAM.ResidualID</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ResidualID => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 103);
        
        /// <summary>
        /// <b>System.SAM.ScriptPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ScriptPath => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 12);
        
        /// <summary>
        /// <b>System.SAM.SecurityID</b> of <b>VT_BLOB</b> type.
        /// </summary>
        public static PROPERTYKEY SecurityID => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 18);
        
        /// <summary>
        /// <b>System.SAM.ServiceLogin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY ServiceLogin => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 39);
        
        /// <summary>
        /// <b>System.SAM.ShellAdminObjectProps</b> of <b>VT_BLOB</b> type.
        /// </summary>
        public static PROPERTYKEY ShellAdminObjectProps => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 47);
        
        /// <summary>
        /// <b>System.SAM.Type</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Type => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 32);
        
        /// <summary>
        /// <b>System.SAM.UserAccountControl</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY UserAccountControl => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 19);
        
        /// <summary>
        /// <b>System.SAM.UserComment</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UserComment => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 16);
        
        /// <summary>
        /// <b>System.SAM.UserPicture</b> of <b>VT_BLOB</b> type.
        /// </summary>
        public static PROPERTYKEY UserPicture => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 24);
        
        /// <summary>
        /// <b>System.SAM.Version</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY Version => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 3);
        
        /// <summary>
        /// <b>System.SAM.Workstations</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Workstations => new(new("705d8364-7547-468c-8c88-84860bcbed4c"), 15);
    }
    
    public static class ScanStatus
    {
        /// <summary>
        /// <b>System.ScanStatus.Profile</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Profile => new(new("dd141766-313a-4a30-90f0-056a7c968437"), 9);
    }
    
    public static class Search
    {
        /// <summary>
        /// <b>System.Search.AccessCount</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY AccessCount => new(new("0b63e350-9ccc-11d0-bcdb-00805fccce04"), 9);
        
        /// <summary>
        /// <b>System.Search.AutoCategory</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY AutoCategory => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 31);
        
        /// <summary>
        /// <b>System.Search.AutoSummary</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY AutoSummary => new(new("560c36c0-503a-11cf-baa1-00004c752a9a"), 2);
        
        /// <summary>
        /// <b>System.Search.ClassID</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY ClassID => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 3);
        
        /// <summary>
        /// <b>System.Search.ContainerHash</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ContainerHash => new(new("bceee283-35df-4d53-826a-f36a3eefc6be"), 100);
        
        /// <summary>
        /// <b>System.Search.ContentSnippet</b> of <b>VT_BLOB</b> type.
        /// </summary>
        public static PROPERTYKEY ContentSnippet => new(new("0abe4d16-9384-426b-b41a-eac3c8e0f147"), 2);
        
        /// <summary>
        /// <b>System.Search.Contents</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Contents => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 19);
        
        /// <summary>
        /// <b>System.Search.EntryID</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY EntryID => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 5);
        
        /// <summary>
        /// <b>System.Search.GatherTime</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY GatherTime => new(new("0b63e350-9ccc-11d0-bcdb-00805fccce04"), 8);
        
        /// <summary>
        /// <b>System.Search.HitCount</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY HitCount => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 4);
        
        /// <summary>
        /// <b>System.Search.IsClosedDirectory</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsClosedDirectory => new(new("0b63e343-9ccc-11d0-bcdb-00805fccce04"), 23);
        
        /// <summary>
        /// <b>System.Search.IsFullyContained</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsFullyContained => new(new("0b63e343-9ccc-11d0-bcdb-00805fccce04"), 24);
        
        /// <summary>
        /// <b>System.Search.LastChangeUSN</b> of <b>VT_I8</b> type.
        /// </summary>
        public static PROPERTYKEY LastChangeUSN => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 9);
        
        /// <summary>
        /// <b>System.Search.LastIndexedTotalTime</b> of <b>VT_R8</b> type.
        /// </summary>
        public static PROPERTYKEY LastIndexedTotalTime => new(new("0b63e350-9ccc-11d0-bcdb-00805fccce04"), 11);
        
        /// <summary>
        /// <b>System.Search.MatchKind</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY MatchKind => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 29);
        
        /// <summary>
        /// <b>System.Search.MatchTags</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY MatchTags => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 30);
        
        /// <summary>
        /// <b>System.Search.OcrContent</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OcrContent => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 28);
        
        /// <summary>
        /// <b>System.Search.ProviderAttributes</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ProviderAttributes => new(new("ad763ac7-f1ed-4039-9fb4-b7b84ef33cef"), 2);
        
        /// <summary>
        /// <b>System.Search.ProviderClass</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProviderClass => new(new("0b63e343-9ccc-11d0-bcdb-00805fccce04"), 25);
        
        /// <summary>
        /// <b>System.Search.ProviderResultLimit</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ProviderResultLimit => new(new("0b63e343-9ccc-11d0-bcdb-00805fccce04"), 27);
        
        /// <summary>
        /// <b>System.Search.ProviderWebDomain</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProviderWebDomain => new(new("0b63e343-9ccc-11d0-bcdb-00805fccce04"), 26);
        
        /// <summary>
        /// <b>System.Search.QueryFocusedSummary</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY QueryFocusedSummary => new(new("560c36c0-503a-11cf-baa1-00004c752a9a"), 3);
        
        /// <summary>
        /// <b>System.Search.QueryFocusedSummaryWithFallback</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY QueryFocusedSummaryWithFallback => new(new("560c36c0-503a-11cf-baa1-00004c752a9a"), 4);
        
        /// <summary>
        /// <b>System.Search.QueryPropertyHits</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY QueryPropertyHits => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 21);
        
        /// <summary>
        /// <b>System.Search.Rank</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY Rank => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 3);
        
        /// <summary>
        /// <b>System.Search.ResultSetAggregateAttributes</b> of <b>VT_UNKNOWN</b> type.
        /// </summary>
        public static PROPERTYKEY ResultSetAggregateAttributes => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 28);
        
        /// <summary>
        /// <b>System.Search.ResultsRank</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY ResultsRank => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 2);
        
        /// <summary>
        /// <b>System.Search.ReverseFileName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ReverseFileName => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 8);
        
        /// <summary>
        /// <b>System.Search.RowID</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY RowID => new(new("49691c90-7e17-101a-a91c-08002b2ecda9"), 15);
        
        /// <summary>
        /// <b>System.Search.Scope</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Scope => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 22);
        
        /// <summary>
        /// <b>System.Search.ShortName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ShortName => new(new("b725f130-47ef-101a-a5f1-02608c9eebac"), 20);
        
        /// <summary>
        /// <b>System.Search.Store</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Store => new(new("a06992b3-8caf-4ed7-a547-b259e32ac9fc"), 100);
        
        /// <summary>
        /// <b>System.Search.UrlToIndex</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UrlToIndex => new(new("0b63e343-9ccc-11d0-bcdb-00805fccce04"), 2);
        
        /// <summary>
        /// <b>System.Search.UrlToIndexWithModificationTime</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY UrlToIndexWithModificationTime => new(new("0b63e343-9ccc-11d0-bcdb-00805fccce04"), 12);
    }
    
    public static class SecondaryTile
    {
        /// <summary>
        /// <b>System.SecondaryTile.IsUninstalled</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsUninstalled => new(new("09736039-456b-4219-ba3e-ec573b58cf97"), 2);
    }
    
    public static class Security
    {
        /// <summary>
        /// <b>System.Security.AllowedEnterpriseDataProtectionIdentities</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY AllowedEnterpriseDataProtectionIdentities => new(new("38d43380-d418-4830-84d5-46935a81c5c6"), 32);
        
        /// <summary>
        /// <b>System.Security.EncryptionOwners</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY EncryptionOwners => new(new("5f5aff6a-37e5-4780-97ea-80c7565cf535"), 34);
        
        /// <summary>
        /// <b>System.Security.EncryptionOwnersDisplay</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY EncryptionOwnersDisplay => new(new("de621b8f-e125-43a3-a32d-5665446d632a"), 25);
    }
    
    public static class Setting
    {
        /// <summary>
        /// <b>System.Setting.Condition</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Condition => new(new("5ab5c75f-15e1-4d65-924a-04754567243c"), 6);
        
        /// <summary>
        /// <b>System.Setting.FontFamily</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FontFamily => new(new("5ab5c75f-15e1-4d65-924a-04754567243c"), 9);
        
        /// <summary>
        /// <b>System.Setting.Glyph</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Glyph => new(new("5ab5c75f-15e1-4d65-924a-04754567243c"), 7);
        
        /// <summary>
        /// <b>System.Setting.GlyphRtl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY GlyphRtl => new(new("5ab5c75f-15e1-4d65-924a-04754567243c"), 8);
        
        /// <summary>
        /// <b>System.Setting.GroupID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY GroupID => new(new("5ab5c75f-15e1-4d65-924a-04754567243c"), 5);
        
        /// <summary>
        /// <b>System.Setting.HostID</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY HostID => new(new("5ab5c75f-15e1-4d65-924a-04754567243c"), 2);
        
        /// <summary>
        /// <b>System.Setting.PageID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PageID => new(new("5ab5c75f-15e1-4d65-924a-04754567243c"), 4);
        
        /// <summary>
        /// <b>System.Setting.SettingID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SettingID => new(new("5ab5c75f-15e1-4d65-924a-04754567243c"), 3);
        
        /// <summary>
        /// <b>System.Setting.SettingsEnvironmentID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SettingsEnvironmentID => new(new("5ab5c75f-15e1-4d65-924a-04754567243c"), 10);
    }
    
    public static class ShareTarget
    {
        /// <summary>
        /// <b>System.ShareTarget.Description</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Description => new(new("86407db8-9df7-48cd-b986-f999adc19731"), 2);
    }
    
    public static class Shell
    {
        /// <summary>
        /// <b>System.Shell.CopilotKeyProviderFastPathMessage</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CopilotKeyProviderFastPathMessage => new(new("38652bca-4329-4e74-86f9-39cf29345eea"), 2);
        
        /// <summary>
        /// <b>System.Shell.Exclusion</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Exclusion => new(new("49eb6558-c09c-46dc-8668-1f848c290d0b"), 1);
        
        /// <summary>
        /// <b>System.Shell.IsDavResource</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsDavResource => new(new("0be3fd71-3f87-40e0-aead-0294cf674635"), 2);
        
        /// <summary>
        /// <b>System.Shell.ItemOfflineStatus</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ItemOfflineStatus => new(new("49eb6558-c09c-46dc-8668-1f848c290d0b"), 3);
        
        /// <summary>
        /// <b>System.Shell.NavPaneRootData</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY NavPaneRootData => new(new("23cdfd25-f531-4c84-a129-488d1506820f"), 1);
        
        /// <summary>
        /// <b>System.Shell.NavPaneRootFilter</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY NavPaneRootFilter => new(new("23cdfd25-f531-4c84-a129-488d1506820f"), 2);
        
        /// <summary>
        /// <b>System.Shell.OmitFromView</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY OmitFromView => new(new("de35258c-c695-4cbc-b982-38b0ad24ced0"), 2);
        
        /// <summary>
        /// <b>System.Shell.SFGAOFlagsStrings</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY SFGAOFlagsStrings => new(new("d6942081-d53b-443d-ad47-5e059d9cd27a"), 2);
    }
    
    public static class Software
    {
        /// <summary>
        /// <b>System.Software.AppId</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY AppId => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 5);
        
        /// <summary>
        /// <b>System.Software.Comments</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Comments => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 20);
        
        /// <summary>
        /// <b>System.Software.DateInstalled</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateInstalled => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 11);
        
        /// <summary>
        /// <b>System.Software.DateLastUsed</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateLastUsed => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 16);
        
        /// <summary>
        /// <b>System.Software.HelpLink</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HelpLink => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 8);
        
        /// <summary>
        /// <b>System.Software.InstallLocation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY InstallLocation => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 9);
        
        /// <summary>
        /// <b>System.Software.InstallSource</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY InstallSource => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 10);
        
        /// <summary>
        /// <b>System.Software.ParentName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ParentName => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 18);
        
        /// <summary>
        /// <b>System.Software.ProductID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProductID => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 19);
        
        /// <summary>
        /// <b>System.Software.ProductName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProductName => new(new("0cef7d53-fa64-11d1-a203-0000f81fedee"), 7);
        
        /// <summary>
        /// <b>System.Software.ProductVersion</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProductVersion => new(new("0cef7d53-fa64-11d1-a203-0000f81fedee"), 8);
        
        /// <summary>
        /// <b>System.Software.Publisher</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Publisher => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 2);
        
        /// <summary>
        /// <b>System.Software.ReadMeUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ReadMeUrl => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 13);
        
        /// <summary>
        /// <b>System.Software.RegisteredCompany</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RegisteredCompany => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 4);
        
        /// <summary>
        /// <b>System.Software.RegisteredOwner</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RegisteredOwner => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 3);
        
        /// <summary>
        /// <b>System.Software.SupportContactName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SupportContactName => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 12);
        
        /// <summary>
        /// <b>System.Software.SupportTelephone</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SupportTelephone => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 7);
        
        /// <summary>
        /// <b>System.Software.SupportUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SupportUrl => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 6);
        
        /// <summary>
        /// <b>System.Software.TasksFileUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY TasksFileUrl => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 17);
        
        /// <summary>
        /// <b>System.Software.TimesUsed</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY TimesUsed => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 15);
        
        /// <summary>
        /// <b>System.Software.UpdateInfoUrl</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY UpdateInfoUrl => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 14);
        
        public static class NullPreview
        {
            /// <summary>
            /// <b>System.Software.NullPreview.Subtitle</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Subtitle => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 998);
            
            /// <summary>
            /// <b>System.Software.NullPreview.Title</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Title => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 999);
            
            /// <summary>
            /// <b>System.Software.NullPreview.TotalSize</b> of <b>VT_UI8</b> type.
            /// </summary>
            public static PROPERTYKEY TotalSize => new(new("841e4f90-ff59-4d16-8947-e81bbffab36d"), 997);
        }
    }
    
    public static class StartMenu
    {
        /// <summary>
        /// <b>System.StartMenu.Group</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Group => new(new("4bd13b3d-e68b-44ec-89ee-7611789d4070"), 100);
        
        /// <summary>
        /// <b>System.StartMenu.GroupItem</b> of <b>VT_UI1, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY GroupItem => new(new("4bd13b3d-e68b-44ec-89ee-7611789d4070"), 103);
        
        /// <summary>
        /// <b>System.StartMenu.IncludeInScope</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IncludeInScope => new(new("4bd13b3d-e68b-44ec-89ee-7611789d4070"), 104);
        
        /// <summary>
        /// <b>System.StartMenu.Query</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Query => new(new("4bd13b3d-e68b-44ec-89ee-7611789d4070"), 102);
        
        /// <summary>
        /// <b>System.StartMenu.ResultSourceId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ResultSourceId => new(new("4bd13b3d-e68b-44ec-89ee-7611789d4070"), 105);
        
        /// <summary>
        /// <b>System.StartMenu.RunCommand</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RunCommand => new(new("4bd13b3d-e68b-44ec-89ee-7611789d4070"), 101);
    }
    
    public static class Storage
    {
        /// <summary>
        /// <b>System.Storage.Portable</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Portable => new(new("4d1ebee8-0803-4774-9842-b77db50265e9"), 2);
        
        /// <summary>
        /// <b>System.Storage.RemovableMedia</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY RemovableMedia => new(new("4d1ebee8-0803-4774-9842-b77db50265e9"), 3);
        
        /// <summary>
        /// <b>System.Storage.SystemCritical</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY SystemCritical => new(new("4d1ebee8-0803-4774-9842-b77db50265e9"), 4);
    }
    
    public static class StructuredQuery
    {
        /// <summary>
        /// <b>System.StructuredQuery.After</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY After => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 103);
        
        /// <summary>
        /// <b>System.StructuredQuery.Before</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY Before => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 102);
        
        /// <summary>
        /// <b>System.StructuredQuery.File</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY File => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 104);
        
        /// <summary>
        /// <b>System.StructuredQuery.Has</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY Has => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 110);
        
        /// <summary>
        /// <b>System.StructuredQuery.Is</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY Is => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 111);
        
        /// <summary>
        /// <b>System.StructuredQuery.Null</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Null => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 112);
        
        public static class CustomProperty
        {
            /// <summary>
            /// <b>System.StructuredQuery.CustomProperty.Boolean</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY Boolean => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 105);
            
            /// <summary>
            /// <b>System.StructuredQuery.CustomProperty.DateTime</b> of <b>VT_FILETIME</b> type.
            /// </summary>
            public static PROPERTYKEY DateTime => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 109);
            
            /// <summary>
            /// <b>System.StructuredQuery.CustomProperty.FloatingPoint</b> of <b>VT_R8</b> type.
            /// </summary>
            public static PROPERTYKEY FloatingPoint => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 107);
            
            /// <summary>
            /// <b>System.StructuredQuery.CustomProperty.Integer</b> of <b>VT_I4</b> type.
            /// </summary>
            public static PROPERTYKEY Integer => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 106);
            
            /// <summary>
            /// <b>System.StructuredQuery.CustomProperty.String</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY String => new(new("7036dcfc-69ab-4316-b5ac-50de702447b0"), 108);
        }
        
        public static class Virtual
        {
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.About</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY About => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 113);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.Artist</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Artist => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 118);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.Bcc</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Bcc => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 102);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.Cc</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Cc => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 103);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.DateCreated</b> of <b>VT_FILETIME</b> type.
            /// </summary>
            public static PROPERTYKEY DateCreated => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 110);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.From</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY From => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 104);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.IsEncrypted</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsEncrypted => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 116);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.IsRead</b> of <b>VT_BOOL</b> type.
            /// </summary>
            public static PROPERTYKEY IsRead => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 114);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.JournalDuration</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY JournalDuration => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 115);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.MessageSize</b> of <b>VT_I4</b> type.
            /// </summary>
            public static PROPERTYKEY MessageSize => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 112);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.OptionalAttendees</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY OptionalAttendees => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 108);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.Organizer</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Organizer => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 106);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.Phone</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Phone => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 111);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.RequiredAttendees</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY RequiredAttendees => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 107);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.Resources</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Resources => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 109);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.To</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY To => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 105);
            
            /// <summary>
            /// <b>System.StructuredQuery.Virtual.Type</b> of <b>VT_LPWSTR</b> type.
            /// </summary>
            public static PROPERTYKEY Type => new(new("e9edd392-0b4c-4cf2-82c0-b0d139666245"), 117);
        }
    }
    
    public static class Supplemental
    {
        /// <summary>
        /// <b>System.Supplemental.Album</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Album => new(new("0c73b141-39d6-4653-a683-cab291eaf95b"), 6);
        
        /// <summary>
        /// <b>System.Supplemental.AlbumID</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY AlbumID => new(new("0c73b141-39d6-4653-a683-cab291eaf95b"), 2);
        
        /// <summary>
        /// <b>System.Supplemental.Location</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Location => new(new("0c73b141-39d6-4653-a683-cab291eaf95b"), 5);
        
        /// <summary>
        /// <b>System.Supplemental.Person</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Person => new(new("0c73b141-39d6-4653-a683-cab291eaf95b"), 7);
        
        /// <summary>
        /// <b>System.Supplemental.ResourceId</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ResourceId => new(new("0c73b141-39d6-4653-a683-cab291eaf95b"), 3);
        
        /// <summary>
        /// <b>System.Supplemental.Tag</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Tag => new(new("0c73b141-39d6-4653-a683-cab291eaf95b"), 4);
    }
    
    public static class Sync
    {
        /// <summary>
        /// <b>System.Sync.Comments</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Comments => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 13);
        
        /// <summary>
        /// <b>System.Sync.ConflictCount</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ConflictCount => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 11);
        
        /// <summary>
        /// <b>System.Sync.ConflictDescription</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ConflictDescription => new(new("ce50c159-2fb8-41fd-be68-d3e042e274bc"), 4);
        
        /// <summary>
        /// <b>System.Sync.ConflictFirstLocation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ConflictFirstLocation => new(new("ce50c159-2fb8-41fd-be68-d3e042e274bc"), 6);
        
        /// <summary>
        /// <b>System.Sync.ConflictItemFullLocation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ConflictItemFullLocation => new(new("b180ad60-ed3f-4d16-bd43-f5b4fcf325a9"), 3);
        
        /// <summary>
        /// <b>System.Sync.ConflictItemShortLocation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ConflictItemShortLocation => new(new("b180ad60-ed3f-4d16-bd43-f5b4fcf325a9"), 2);
        
        /// <summary>
        /// <b>System.Sync.ConflictSecondLocation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ConflictSecondLocation => new(new("ce50c159-2fb8-41fd-be68-d3e042e274bc"), 7);
        
        /// <summary>
        /// <b>System.Sync.ConflictUnresolvable</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY ConflictUnresolvable => new(new("ce50c159-2fb8-41fd-be68-d3e042e274bc"), 10);
        
        /// <summary>
        /// <b>System.Sync.Connected</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Connected => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 16);
        
        /// <summary>
        /// <b>System.Sync.Context</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Context => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 19);
        
        /// <summary>
        /// <b>System.Sync.CopyIn</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CopyIn => new(new("328d8b21-7729-4bfc-954c-902b329d56b0"), 2);
        
        /// <summary>
        /// <b>System.Sync.DateSynchronized</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY DateSynchronized => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 7);
        
        /// <summary>
        /// <b>System.Sync.Enabled</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Enabled => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 14);
        
        /// <summary>
        /// <b>System.Sync.ErrorCount</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ErrorCount => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 12);
        
        /// <summary>
        /// <b>System.Sync.EventDescription</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EventDescription => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 4);
        
        /// <summary>
        /// <b>System.Sync.EventFlags</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY EventFlags => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 21);
        
        /// <summary>
        /// <b>System.Sync.EventLevel</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY EventLevel => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 20);
        
        /// <summary>
        /// <b>System.Sync.GlobalActivityMessage</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY GlobalActivityMessage => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 29);
        
        /// <summary>
        /// <b>System.Sync.HandlerCollectionID</b> of <b>VT_CLSID</b> type.
        /// </summary>
        public static PROPERTYKEY HandlerCollectionID => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 2);
        
        /// <summary>
        /// <b>System.Sync.HandlerID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HandlerID => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 3);
        
        /// <summary>
        /// <b>System.Sync.HandlerName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HandlerName => new(new("ce50c159-2fb8-41fd-be68-d3e042e274bc"), 2);
        
        /// <summary>
        /// <b>System.Sync.HandlerType</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY HandlerType => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 8);
        
        /// <summary>
        /// <b>System.Sync.HandlerTypeLabel</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HandlerTypeLabel => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 9);
        
        /// <summary>
        /// <b>System.Sync.Hidden</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY Hidden => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 15);
        
        /// <summary>
        /// <b>System.Sync.ItemID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ItemID => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 6);
        
        /// <summary>
        /// <b>System.Sync.ItemName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ItemName => new(new("ce50c159-2fb8-41fd-be68-d3e042e274bc"), 3);
        
        /// <summary>
        /// <b>System.Sync.ItemState</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ItemState => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 25);
        
        /// <summary>
        /// <b>System.Sync.ItemStatusAction</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ItemStatusAction => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 28);
        
        /// <summary>
        /// <b>System.Sync.ItemStatusDescription</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ItemStatusDescription => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 27);
        
        /// <summary>
        /// <b>System.Sync.ItemStatusText</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ItemStatusText => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 26);
        
        /// <summary>
        /// <b>System.Sync.LastSyncedMessage</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LastSyncedMessage => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 30);
        
        /// <summary>
        /// <b>System.Sync.Link</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Link => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 17);
        
        /// <summary>
        /// <b>System.Sync.Progress</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Progress => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 5);
        
        /// <summary>
        /// <b>System.Sync.ProgressPercentage</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY ProgressPercentage => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 23);
        
        /// <summary>
        /// <b>System.Sync.State</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY State => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 24);
        
        /// <summary>
        /// <b>System.Sync.Status</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Status => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 10);
        
        /// <summary>
        /// <b>System.Sync.SyncResults</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SyncResults => new(new("7bd5533e-af15-44db-b8c8-bd6624e1d032"), 22);
    }
    
    public static class Task
    {
        /// <summary>
        /// <b>System.Task.BillingInformation</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BillingInformation => new(new("d37d52c6-261c-4303-82b3-08b926ac6f12"), 100);
        
        /// <summary>
        /// <b>System.Task.CompletionStatus</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CompletionStatus => new(new("084d8a0a-e6d5-40de-bf1f-c8820e7c877c"), 100);
        
        /// <summary>
        /// <b>System.Task.Owner</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Owner => new(new("08c7cc5f-60f2-4494-ad75-55e3e0b5add0"), 100);
    }
    
    public static class Taskbar
    {
        /// <summary>
        /// <b>System.Taskbar.PinnedWebsite</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY PinnedWebsite => new(new("57086c23-86c6-478f-afb2-236188c8f47f"), 4);
        
        /// <summary>
        /// <b>System.Taskbar.PinnedWebsiteIconUri</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PinnedWebsiteIconUri => new(new("57086c23-86c6-478f-afb2-236188c8f47f"), 5);
        
        /// <summary>
        /// <b>System.Taskbar.PinnedWebsiteURL</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY PinnedWebsiteURL => new(new("57086c23-86c6-478f-afb2-236188c8f47f"), 6);
        
        /// <summary>
        /// <b>System.Taskbar.TabActive</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY TabActive => new(new("57086c23-86c6-478f-afb2-236188c8f47f"), 2);
        
        /// <summary>
        /// <b>System.Taskbar.TabList</b> of <b>VT_UI4, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY TabList => new(new("57086c23-86c6-478f-afb2-236188c8f47f"), 3);
    }
    
    public static class Tile
    {
        /// <summary>
        /// <b>System.Tile.Background</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Background => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 4);
        
        /// <summary>
        /// <b>System.Tile.BadgeLogoPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY BadgeLogoPath => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 15);
        
        /// <summary>
        /// <b>System.Tile.DisplayNameLanguage</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY DisplayNameLanguage => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 18);
        
        /// <summary>
        /// <b>System.Tile.EncodedTargetPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY EncodedTargetPath => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 23);
        
        /// <summary>
        /// <b>System.Tile.FencePost</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FencePost => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 21);
        
        /// <summary>
        /// <b>System.Tile.Flags</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Flags => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 14);
        
        /// <summary>
        /// <b>System.Tile.Foreground</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Foreground => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 5);
        
        /// <summary>
        /// <b>System.Tile.HoloBoundingBox</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HoloBoundingBox => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 25);
        
        /// <summary>
        /// <b>System.Tile.HoloContent</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY HoloContent => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 24);
        
        /// <summary>
        /// <b>System.Tile.InstallProgress</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY InstallProgress => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 22);
        
        /// <summary>
        /// <b>System.Tile.LongDisplayName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LongDisplayName => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 11);
        
        /// <summary>
        /// <b>System.Tile.ShortDisplayName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ShortDisplayName => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 26);
        
        /// <summary>
        /// <b>System.Tile.SmallLogoPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SmallLogoPath => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 2);
        
        /// <summary>
        /// <b>System.Tile.Square150x150LogoPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Square150x150LogoPath => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 12);
        
        /// <summary>
        /// <b>System.Tile.Square310x310LogoPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Square310x310LogoPath => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 19);
        
        /// <summary>
        /// <b>System.Tile.Square70x70LogoPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Square70x70LogoPath => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 20);
        
        /// <summary>
        /// <b>System.Tile.SuiteDisplayName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SuiteDisplayName => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 16);
        
        /// <summary>
        /// <b>System.Tile.SuiteSortName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY SuiteSortName => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 17);
        
        /// <summary>
        /// <b>System.Tile.Wide310x150LogoPath</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Wide310x150LogoPath => new(new("86d40b4d-9069-443c-819a-2a54090dccec"), 13);
    }
    
    public static class VersionControl
    {
        /// <summary>
        /// <b>System.VersionControl.CurrentFolderStatus</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY CurrentFolderStatus => new(new("73fd4f20-113b-4f94-b72c-93b71ded3eaa"), 8);
        
        /// <summary>
        /// <b>System.VersionControl.LastChangeAuthorEmail</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LastChangeAuthorEmail => new(new("73fd4f20-113b-4f94-b72c-93b71ded3eaa"), 4);
        
        /// <summary>
        /// <b>System.VersionControl.LastChangeAuthorName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LastChangeAuthorName => new(new("73fd4f20-113b-4f94-b72c-93b71ded3eaa"), 5);
        
        /// <summary>
        /// <b>System.VersionControl.LastChangeDate</b> of <b>VT_FILETIME</b> type.
        /// </summary>
        public static PROPERTYKEY LastChangeDate => new(new("73fd4f20-113b-4f94-b72c-93b71ded3eaa"), 3);
        
        /// <summary>
        /// <b>System.VersionControl.LastChangeID</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LastChangeID => new(new("73fd4f20-113b-4f94-b72c-93b71ded3eaa"), 6);
        
        /// <summary>
        /// <b>System.VersionControl.LastChangeMessage</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY LastChangeMessage => new(new("73fd4f20-113b-4f94-b72c-93b71ded3eaa"), 7);
        
        /// <summary>
        /// <b>System.VersionControl.Status</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Status => new(new("73fd4f20-113b-4f94-b72c-93b71ded3eaa"), 2);
    }
    
    public static class Video
    {
        /// <summary>
        /// <b>System.Video.Compression</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Compression => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 10);
        
        /// <summary>
        /// <b>System.Video.Director</b> of <b>VT_LPWSTR, VT_VECTOR</b> type.
        /// </summary>
        public static PROPERTYKEY Director => new(new("64440492-4c8b-11d1-8b70-080036b11a03"), 20);
        
        /// <summary>
        /// <b>System.Video.EncodingBitrate</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY EncodingBitrate => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 8);
        
        /// <summary>
        /// <b>System.Video.FourCC</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FourCC => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 44);
        
        /// <summary>
        /// <b>System.Video.FrameHeight</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FrameHeight => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 4);
        
        /// <summary>
        /// <b>System.Video.FrameRate</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FrameRate => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 6);
        
        /// <summary>
        /// <b>System.Video.FrameWidth</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY FrameWidth => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 3);
        
        /// <summary>
        /// <b>System.Video.HorizontalAspectRatio</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY HorizontalAspectRatio => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 42);
        
        /// <summary>
        /// <b>System.Video.IsSpherical</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsSpherical => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 100);
        
        /// <summary>
        /// <b>System.Video.IsStereo</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsStereo => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 98);
        
        /// <summary>
        /// <b>System.Video.Orientation</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Orientation => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 99);
        
        /// <summary>
        /// <b>System.Video.SampleSize</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY SampleSize => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 9);
        
        /// <summary>
        /// <b>System.Video.StreamName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY StreamName => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 2);
        
        /// <summary>
        /// <b>System.Video.StreamNumber</b> of <b>VT_UI2</b> type.
        /// </summary>
        public static PROPERTYKEY StreamNumber => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 11);
        
        /// <summary>
        /// <b>System.Video.TotalBitrate</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY TotalBitrate => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 43);
        
        /// <summary>
        /// <b>System.Video.TranscodedForSync</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY TranscodedForSync => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 46);
        
        /// <summary>
        /// <b>System.Video.VerticalAspectRatio</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY VerticalAspectRatio => new(new("64440491-4c8b-11d1-8b70-080036b11a03"), 45);
    }
    
    public static class Volume
    {
        /// <summary>
        /// <b>System.Volume.BitLockerCanChangePassphraseByProxy</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY BitLockerCanChangePassphraseByProxy => new(new("a63b464f-2ace-4d83-87ae-abaf011cc6ac"), 1720);
        
        /// <summary>
        /// <b>System.Volume.BitLockerCanChangePin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY BitLockerCanChangePin => new(new("c53e42a9-db3c-4bc7-b0f3-83a524adf0ec"), 1719);
        
        /// <summary>
        /// <b>System.Volume.BitLockerProtection</b> of <b>VT_I4</b> type.
        /// </summary>
        public static PROPERTYKEY BitLockerProtection => new(new("2d15a9a1-a556-4189-91ad-027458f11a07"), 1717);
        
        /// <summary>
        /// <b>System.Volume.BitLockerRequiresAdmin</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY BitLockerRequiresAdmin => new(new("ee31306c-fb9b-4d62-8621-3575d972a9f9"), 1718);
        
        /// <summary>
        /// <b>System.Volume.FileSystem</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY FileSystem => new(new("9b174b35-40ff-11d2-a27e-00c04fc30871"), 4);
        
        /// <summary>
        /// <b>System.Volume.IsMappedDrive</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsMappedDrive => new(new("149c0b69-2c2d-48fc-808f-d318d78c4636"), 2);
        
        /// <summary>
        /// <b>System.Volume.IsRoot</b> of <b>VT_BOOL</b> type.
        /// </summary>
        public static PROPERTYKEY IsRoot => new(new("9b174b35-40ff-11d2-a27e-00c04fc30871"), 10);
    }
    
    public static class Winx
    {
        /// <summary>
        /// <b>System.Winx.Hash</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Hash => new(new("fb8d2d7b-90d1-4e34-bf60-6eac09922bbf"), 2);
    }
    
    public static class Wireless
    {
        /// <summary>
        /// <b>System.Wireless.ConnectionMode</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ConnectionMode => new(new("9b34bbb9-949c-488d-9a6d-eeb47c847a2f"), 9);
        
        /// <summary>
        /// <b>System.Wireless.ProfileName</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY ProfileName => new(new("9b34bbb9-949c-488d-9a6d-eeb47c847a2f"), 2);
        
        /// <summary>
        /// <b>System.Wireless.RadioType</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY RadioType => new(new("9b34bbb9-949c-488d-9a6d-eeb47c847a2f"), 5);
        
        /// <summary>
        /// <b>System.Wireless.Security</b> of <b>VT_LPWSTR</b> type.
        /// </summary>
        public static PROPERTYKEY Security => new(new("9b34bbb9-949c-488d-9a6d-eeb47c847a2f"), 4);
    }
    
    public static class ZipFolder
    {
        /// <summary>
        /// <b>System.ZipFolder.CRC32</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY CRC32 => new(new("e88dcce0-b7b3-11d1-a9f0-00aa0060fa31"), 5);
        
        /// <summary>
        /// <b>System.ZipFolder.CompressedSize</b> of <b>VT_UI8</b> type.
        /// </summary>
        public static PROPERTYKEY CompressedSize => new(new("e88dcce0-b7b3-11d1-a9f0-00aa0060fa31"), 6);
        
        /// <summary>
        /// <b>System.ZipFolder.Encrypted</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Encrypted => new(new("e88dcce0-b7b3-11d1-a9f0-00aa0060fa31"), 2);
        
        /// <summary>
        /// <b>System.ZipFolder.Method</b> of <b>VT_NULL</b> type.
        /// </summary>
        public static PROPERTYKEY Method => new(new("e88dcce0-b7b3-11d1-a9f0-00aa0060fa31"), 3);
        
        /// <summary>
        /// <b>System.ZipFolder.Ratio</b> of <b>VT_UI4</b> type.
        /// </summary>
        public static PROPERTYKEY Ratio => new(new("e88dcce0-b7b3-11d1-a9f0-00aa0060fa31"), 4);
    }
}
