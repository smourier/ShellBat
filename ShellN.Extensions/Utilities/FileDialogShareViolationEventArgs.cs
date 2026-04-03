namespace ShellN.Extensions.Utilities;

public class FileDialogShareViolationEventArgs(IFileDialog dlg, IShellItem folder) : FileDialogEventArgs(dlg, FileDialogEventType.ShareViolation)
{
    public IShellItem Folder { get; } = folder;
    public FDE_SHAREVIOLATION_RESPONSE Response { get; set; }
    public virtual bool Handled { get; set; }
}
