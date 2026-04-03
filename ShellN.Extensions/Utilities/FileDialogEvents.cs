namespace ShellN.Extensions.Utilities;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class FileDialogEvents : IFileDialogEvents
{
    public event EventHandler<FileDialogEventArgs>? Event;

    protected virtual void OnEvent(object sender, FileDialogEventArgs e) => Event?.Invoke(sender, e);

    HRESULT IFileDialogEvents.OnFileOk(IFileDialog pfd)
    {
        OnEvent(this, new FileDialogEventArgs(pfd, FileDialogEventType.FileOk));
        return DirectN.Constants.S_OK;
    }

    HRESULT IFileDialogEvents.OnFolderChange(IFileDialog pfd)
    {
        OnEvent(this, new FileDialogEventArgs(pfd, FileDialogEventType.FolderChange));
        return DirectN.Constants.S_OK;
    }

    HRESULT IFileDialogEvents.OnFolderChanging(IFileDialog pfd, IShellItem psiFolder)
    {
        var e = new FileDialogFolderChangingEventArgs(pfd, psiFolder);
        OnEvent(this, e);
        if (e.Cancel)
            return DirectN.Constants.E_FAIL;

        return DirectN.Constants.S_OK;
    }

    HRESULT IFileDialogEvents.OnOverwrite(IFileDialog pfd, IShellItem psi, out FDE_OVERWRITE_RESPONSE pResponse)
    {
        var e = new FileDialogOverwriteEventArgs(pfd, psi);
        OnEvent(this, e);
        if (e.Handled)
        {
            pResponse = e.Response;
        }
        else
        {
            pResponse = FDE_OVERWRITE_RESPONSE.FDEOR_DEFAULT;
        }
        return DirectN.Constants.S_OK;
    }

    HRESULT IFileDialogEvents.OnSelectionChange(IFileDialog pfd)
    {
        OnEvent(this, new FileDialogEventArgs(pfd, FileDialogEventType.SelectionChange));
        return DirectN.Constants.S_OK;
    }

    HRESULT IFileDialogEvents.OnShareViolation(IFileDialog pfd, IShellItem psi, out FDE_SHAREVIOLATION_RESPONSE pResponse)
    {
        var e = new FileDialogShareViolationEventArgs(pfd, psi);
        OnEvent(this, e);
        if (e.Handled)
        {
            pResponse = e.Response;
        }
        else
        {
            pResponse = FDE_SHAREVIOLATION_RESPONSE.FDESVR_DEFAULT;
        }
        return DirectN.Constants.S_OK;
    }

    HRESULT IFileDialogEvents.OnTypeChange(IFileDialog pfd)
    {
        OnEvent(this, new FileDialogEventArgs(pfd, FileDialogEventType.TypeChange));
        return DirectN.Constants.S_OK;
    }
}
