#nullable enable
namespace ShellN;

// https://learn.microsoft.com/windows/win32/api/searchapi/ne-searchapi-clusion_reason
public enum CLUSION_REASON
{
    CLUSIONREASON_UNKNOWNSCOPE = 0,
    CLUSIONREASON_DEFAULT = 1,
    CLUSIONREASON_USER = 2,
    CLUSIONREASON_GROUPPOLICY = 3,
    CLUSIONREASON_USAGE = 4,
}
