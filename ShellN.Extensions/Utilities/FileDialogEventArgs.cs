namespace ShellN.Extensions.Utilities;

public class FileDialogEventArgs(IFileDialog dialog, FileDialogEventType type) : CancelEventArgs
{
    public IFileDialog Dialog { get; } = dialog;
    public FileDialogEventType Type { get; } = type;
}
