namespace ShellN.Extensions.Utilities;

public class FileDialogFolderChangingEventArgs(IFileDialog dlg, IShellItem folder) : FileDialogEventArgs(dlg, FileDialogEventType.FolderChanging)
{
    public IShellItem Folder { get; } = folder;
}
