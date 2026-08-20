namespace ShellBat.Model.Viewers;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class PreviewViewer(Entry entry) : Viewer(entry), IPreviewHandlerFrame
{
    private PreviewWindow? _window;
    private ShellBatWindow? _parentWindow;
    private IComObject<IPreviewHandler>? _previewHandler;
    private RECT? _lastRc;

    public override bool IsSupported
    {
        get
        {
            if (Entry.Extension.PreviewHandlerClsid is not null)
                return true;

            return base.IsSupported;
        }
    }

    public bool IsVisible => _window?.IsVisible == true;
    public bool IsHidden => !IsVisible;

    public string? GetCaptureFilePath()
    {
        using var bmp = _window?.GdiCapture();
        if (bmp == null)
            return null;

        var id = Guid.NewGuid().ToString("N");
        var cacheFileName = Path.Combine(Settings.HttpServerCachePath, UrlType.PreviewImage.ToString(), id + ".png");
        IOUtilities.FileEnsureDirectory(cacheFileName);
        bmp.Save(cacheFileName);
        return cacheFileName;
    }

    public virtual void Hide() => _window?.Hide();
    public virtual void Preview(JsonElement json, bool show)
    {
        if (_window?.IsDisposingOrDisposed == true)
            return; // happens (race condition)

        if (_parentWindow == null)
        {
            var handle = new HWND(json.GetValue<nint>("windowHandle"));
            _parentWindow = ShellBatApplication.ShellBatWindows.FirstOrDefault(w => w.Handle == handle);
            if (_parentWindow != null)
            {
                _parentWindow.PositionChanging += OnParentWindowPositionChanging;
            }
        }
        if (_parentWindow == null)
            return;

        _parentWindow.RegisterPreviewViewer(this);

        _previewHandler ??= CreateHandler(2);
        if (_previewHandler == null)
            return;

        var wrc = _parentWindow.WindowRect;
        var rc = json.GetDOMRect();
        _lastRc = rc;
        var entriesRc = json.GetDOMRect("e");

        _window ??= new PreviewWindow(_parentWindow);

        rc.left = (int)(rc.left * _parentWindow.Scale + wrc.left);
        rc.top = (int)(rc.top * _parentWindow.Scale + wrc.top);
        rc.right = (int)(rc.right * _parentWindow.Scale + wrc.left);
        rc.bottom = (int)(rc.bottom * _parentWindow.Scale + wrc.top);

        var right = Math.Min(rc.right, wrc.right);
        var bottom = Math.Min(rc.bottom, wrc.bottom - (wrc.Height - entriesRc.Height * _parentWindow.Scale));

        var newRC = RECT.Sized(
            (int)(rc.left + entriesRc.left),
            (int)(rc.top + entriesRc.top),
            (int)(right - rc.left),
            (int)(bottom - rc.top)
            );

        _window.ResizeAndMove(newRC);
        _window.Preview(_previewHandler, _window.ClientRect, show);
    }

    private void OnParentWindowPositionChanging(object? sender, ValueEventArgs<WINDOWPOS> e) => _ = Refresh();

    public async Task Refresh()
    {
        var window = _window;
        var handler = _previewHandler;
        var parentWindow = _parentWindow;
        if (_lastRc == null || window == null || handler == null || parentWindow == null)
            return;

        var entriesRc = await parentWindow.GetEntriesRect(false);
        if (entriesRc.Width == 0 || entriesRc.Height == 0)
            return; // window is probably dying

        var wrc = parentWindow.WindowRect;
        var rc = _lastRc.Value;
        rc.left = (int)(rc.left * parentWindow.Scale + wrc.left);
        rc.top = (int)(rc.top * parentWindow.Scale + wrc.top);
        rc.right = (int)(rc.right * parentWindow.Scale + wrc.left);
        rc.bottom = (int)(rc.bottom * parentWindow.Scale + wrc.top);

        var right = Math.Min(rc.right, wrc.right);
        var bottom = Math.Min(rc.bottom, wrc.bottom - (wrc.Height - entriesRc.Height * parentWindow.Scale));

        window.ResizeAndMove(
            (int)(rc.left + entriesRc.left),
            (int)(rc.top + entriesRc.top),
            (int)(right - rc.left),
            (int)(bottom - rc.top));
        window.Preview(handler, window.ClientRect, false);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (_parentWindow != null)
            {
                _parentWindow.PositionChanging -= OnParentWindowPositionChanging;
                _parentWindow.DisposePreviewViewer(this);
            }

            var handler = Interlocked.Exchange(ref _previewHandler, null);
            if (handler != null)
            {
                var hr = handler.Object.Unload(); // don't hard check for error
                var ows = handler.As<IObjectWithSite>();
                ows?.Object.SetSite(0);
                handler.Dispose();
            }

            // note we don't set _window to null here as we want to keep the fact it existed (avoid race conditions)
            _window?.Dispose();
        }
        base.Dispose(disposing);
    }

    protected override bool TryConvertArgument(MethodInfo method, int index, ParameterInfo parameter, object? value, out object? converted)
    {
        if (parameter.ParameterType == typeof(JsonElement) && value is string s)
        {
            converted = JsonDocument.Parse(s).RootElement;
            return true;
        }
        return base.TryConvertArgument(method, index, parameter, value, out converted);
    }

    public static void DeleteTempFiles()
    {
        var cacheDir = Path.Combine(Settings.HttpServerCachePath, UrlType.PreviewImage.ToString());
        IOUtilities.DirectoryDeleteFiles(cacheDir, true, throwOnError: false);
    }

    private IComObject<IPreviewHandler>? CreateHandler(int retryCount)
    {
        var hr = Entry.Extension.TryCreatePreviewHandler(out var handler);
        if (hr.IsError || handler == null)
        {
            Application.TraceError($"TryCreatePreviewHandler failed for '{Entry.ParsingName}': {hr}");
            if (hr == DirectN.Constants.RPC_E_SERVERCALL_RETRYLATER && retryCount > 0)
            {
                // preview handler server is busy, try kill it and retry
                Application.TraceWarning($"Preview handler server is busy for '{Entry.ParsingName}' retry {retryCount} cmdline: '{Entry.Extension.PreviewHandlerProcessCommandLine}'");
                if (TryKillPrevieHandlersProcesses() > 0)
                    return CreateHandler(retryCount - 1);
            }
            return null;
        }

        hr = Initialize(handler);
        if (hr.IsError)
        {
            Application.TraceError($"Initialize failed for '{Entry.ParsingName}': {hr}");
            handler.Dispose();
            return null;
        }

        return handler;
    }

    private int TryKillPrevieHandlersProcesses()
    {
        if (Entry.Extension.PreviewHandlerProcessCommandLine == null ||
            Entry.Extension.PreviewHandlerProcessCommandLineDiscriminant == null)
            return 0;

        var name = Path.GetFileNameWithoutExtension(Entry.Extension.PreviewHandlerProcessCommandLine);
        var count = 0;
        foreach (var process in Process.GetProcessesByName(name))
        {
            var cmdLine = process.GetCommandLine();
            // Application.TraceWarning($"PH process '{process.ProcessName}' (Id: {process.Id}) cmdLine: {cmdLine}");
            if (cmdLine == null)
                continue;

            if (!cmdLine.Contains(Entry.Extension.PreviewHandlerProcessCommandLineDiscriminant))
                continue;

            try
            {
                process.Kill();
                count++;
                //Application.TraceWarning($"PH process '{process.ProcessName}' (Id: {process.Id}) cmdLine: {cmdLine} was killed.");
            }
            catch (Exception ex)
            {
                Application.TraceError($"Failed to kill preview handler process '{process.ProcessName}' (Id: {process.Id}): {ex}");
            }
        }
        return count;
    }

    private HRESULT Initialize(IComObject<IPreviewHandler> handler)
    {
        // preferred and recommended but not always supported
        var iws = handler.As<IInitializeWithStream>();
        if (iws == null)
        {
            // the most common initialization interface
            var iwf = handler.As<IInitializeWithFile>();
            if (iwf == null)
            {
                // .msg files and others previewers only support IInitializeWithItem
                var iwi = handler.As<IInitializeWithItem>();
                if (iwi == null)
                    return DirectN.Constants.E_NOINTERFACE;

                var item = Entry.GetItem(Entry.ParsingName, ShellItemParsingOptions.DontThrowOnError);
                if (item == null)
                    return DirectN.Constants.E_FAIL;

                var hr = iwi.Object.Initialize(item.NativeObject, (uint)STGM.STGM_READ);
                if (hr.IsError)
                    return hr;
            }
            else
            {
                var hr = iwf.Object.Initialize(PWSTR.From(Entry.ParsingName), (uint)STGM.STGM_READ);
                if (hr.IsError)
                    return hr;
            }
        }
        else
        {
            var stream = Entry.OpenRead();
            if (stream == null)
                return DirectN.Constants.WINCODEC_ERR_STREAMNOTAVAILABLE; // not really the error code but close enough

            var mis = new ManagedIStream(stream, true);
            var hr = iws.Object.Initialize(mis, (uint)STGM.STGM_READ);
            if (hr.IsError)
                return hr;
        }

        var ows = handler.As<IObjectWithSite>();
        if (ows != null)
        {
            ComObject.WithComInstance(this, unk => ows?.Object.SetSite(unk), true);
        }
        return DirectN.Constants.S_OK;
    }

    HRESULT IPreviewHandlerFrame.GetWindowContext(out PREVIEWHANDLERFRAMEINFO info)
    {
        //Application.TraceVerbose("GetWindowContext called");
        info = new PREVIEWHANDLERFRAMEINFO();
        return DirectN.Constants.S_OK;
    }

    HRESULT IPreviewHandlerFrame.TranslateAccelerator(in MSG msg)
    {
        //Application.TraceVerbose($"msg:{MessageDecoder.Decode(msg)}");
        return DirectN.Constants.S_FALSE;
    }

    private sealed partial class PreviewWindow : CompositionWindow
    {
        private bool _setWindowCalled;
        private RECT _lastRc;
        private bool _lastShow;

        public PreviewWindow(Window parent) : base(
            parentHandle: parent.Handle,
            className: $"{Program.AppId}.{typeof(PreviewWindow).Name}{AssemblyUtilities.GetInformationalVersion()}",
            extendedStyle: WINDOW_EX_STYLE.WS_EX_NOACTIVATE | WINDOW_EX_STYLE.WS_EX_LAYERED,
            style: WINDOW_STYLE.WS_POPUP,
            instance: new HINSTANCE { Value = Application.ModuleHandle.Value }
            )
        {
            ParentWindow = parent;
            DestroyOnDispose = true;
            DirectN.Functions.SetLayeredWindowAttributes(Handle, 0, 255, LAYERED_WINDOW_ATTRIBUTES_FLAGS.LWA_ALPHA);
        }

        public Window ParentWindow { get; }

        public void Preview(IComObject<IPreviewHandler> handler, RECT rc, bool show)
        {
            if (_lastRc == rc && _lastShow == show)
                return;

            if (show)
            {
                Show();
            }

            HRESULT hr;
            if (_setWindowCalled)
            {
                hr = handler.Object.SetRect(rc);
                if (hr.IsError)
                {
                    Application.TraceError($"SetRect failed for {handler}: {hr}");
                    return;
                }
            }
            else
            {
                hr = handler.Object.SetWindow(Handle, rc);
                if (hr.IsError)
                {
                    Application.TraceError($"SetWindow failed for {handler}: {hr}");
                    return;
                }
                else
                {
                    _setWindowCalled = true;
                }

                hr = handler.Object.DoPreview();
                if (hr.IsError)
                {
                    Application.TraceError($"DoPreview failed for {handler}: {hr}");
                    return;
                }
            }

            _lastRc = rc;
            _lastShow = show;
        }
    }
}
