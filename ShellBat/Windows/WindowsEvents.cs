namespace ShellBat.Windows;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class WindowsEvents : DispatchEvents
{
    private IComObject<IShellWindows>? _windows;

    public event EventHandler<WindowEventArgs>? WindowEvent;

    public WindowsEvents()
        : base(typeof(DShellWindowsEvents).GUID, ComObject.CoCreate<IShellWindows>(ShellN.Constants.ShellWindows)!)
    {
    }

    public IComObject<IShellWindows>? Windows => _windows;

    protected virtual void OnWindowEvent(object sender, WindowEventArgs e) => WindowEvent?.Invoke(sender, e);

    protected override void SetConnectionPointContainer(object connectionPointContainer)
    {
        _windows = (IComObject<IShellWindows>?)connectionPointContainer;
        base.SetConnectionPointContainer(connectionPointContainer);
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing)
        {
            Interlocked.Exchange(ref _windows, null)?.Dispose();
        }
    }

    protected override void OnEvent(object sender, DispatchEventArgs e)
    {
        const uint DISPID_WINDOWREGISTERED = 200;
        const uint DISPID_WINDOWREVOKED = 201;

        if (e.DispId == DISPID_WINDOWREGISTERED && _windows != null)
        {
            var cookie = Conversions.ChangeType<int>(e.Arguments.FirstOrDefault());
            OnWindowEvent(this, new WindowEventArgs(_windows, cookie, true));
        }
        else if (e.DispId == DISPID_WINDOWREVOKED && _windows != null)
        {
            var cookie = Conversions.ChangeType<int>(e.Arguments.FirstOrDefault());
            OnWindowEvent(this, new WindowEventArgs(_windows, cookie, false));
        }
    }
}
