
namespace ShellN.Extensions.Utilities;

public class FileDialogOverwriteEventArgs(IFileDialog dlg, IShellItem folder) : FileDialogEventArgs(dlg, FileDialogEventType.Overwrite)
{
    public IShellItem Folder { get; } = folder;
    public FDE_OVERWRITE_RESPONSE Response { get; set; }
    public virtual bool Handled { get; set; }
}
