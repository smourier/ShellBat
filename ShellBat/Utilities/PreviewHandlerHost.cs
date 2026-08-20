namespace ShellBat.Utilities;

// this is not used:
// * it doesn't really work reliably with some preview handlers (Adobe PDF for example)
//   because of some timing issue (we don't know for sure when the preview handler finished rendering)
// * if the window is not visible, some preview handlers refuse to render (Adobe PDF for example)
// * it's not better using GraphicsCaptureItem (DXGI-based)
// * if the window is visible, it's ugly
public sealed partial class PreviewHandlerHost : IDisposable
{
    private Semaphore? _semaphore;
    private static int _count;

    public PreviewHandlerHost(int maximumThreads)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maximumThreads);
        _semaphore = new Semaphore(maximumThreads, maximumThreads);
    }

    public async Task<WicBitmapSource?> GetPreview(string parsingName, PreviewHandlerOptions? options = null)
    {
        var semaphore = _semaphore;
        if (semaphore == null)
            return null;

        var entry = Entry.Get(null, parsingName, ShellItemParsingOptions.DontThrowOnError);
        if (entry is null)
        {
            Application.TraceError($"Entry.Get failed for '{parsingName}'.");
            return null;
        }

        if (entry.Extension.PreviewHandlerClsid == null)
        {
            Application.TraceError($"No preview handler available for '{entry.ParsingName}'.");
            return null;
        }

        options ??= new PreviewHandlerOptions();
        options.FinalTimeout = Math.Max(20, options.SnapshotTimeout ?? 500);
        var semaphoreTimeout = options.FinalTimeout * 2;
        if (!semaphore.WaitOne(semaphoreTimeout))
        {
            Application.TraceError($"Timeout {semaphoreTimeout} waiting for a thread for previewing '{entry.ParsingName}'.");
            return null;
        }

        try
        {
            PreviewWindow? win = null;
            var windowReady = new AutoResetEvent(false);

            _ = TaskUtilities.RunWithNewSTAThread(() =>
            {
                try
                {
                    var count = Interlocked.Increment(ref _count);
                    Thread.CurrentThread.IsBackground = true;
                    Thread.CurrentThread.Name = $"{Program.AppId}.PreviewThread#{count}";
                    WindowSynchronizationContext.WithContext(() =>
                    {
                        using var dw = new ApplicationWithScheduler();
                        dw.ExitOnLastWindowRemoved = true;
                        var window = new PreviewWindow(count);
                        window.Move(DirectN.Constants.WHERE_NOONE_CAN_SEE_ME, DirectN.Constants.WHERE_NOONE_CAN_SEE_ME);
                        //window.ResizeAndMove(2560, 0, 100, 100);
                        window.Show();
                        win = window;
                        windowReady.Set();
                        dw.Run();
                    });
                }
                catch (Exception ex)
                {
                    Application.TraceError($"Preview thread failed for '{entry.ParsingName}': {ex}");
                }
            });

            var winReady = windowReady.WaitOne(semaphoreTimeout);
            if (!winReady)
            {
                win?.RunTaskOnUIThread(() => win?.Dispose());
                Application.TraceError($"Timeout {semaphoreTimeout} waiting for the preview window to be ready for '{entry.ParsingName}'.");
                return null;
            }

            windowReady.Dispose();
            var task = await win!.RunTaskOnUIThread(() => win.Preview(entry, options), true);
            var bmp = await task;
            win?.RunTaskOnUIThread(() => win?.Dispose());
            return bmp;
        }
        finally
        {
            semaphore.Release();
        }
    }

    public void Dispose()
    {
        Interlocked.Exchange(ref _semaphore, null)?.Dispose();
    }

    private sealed partial class PreviewWindow : Window
    {
        public PreviewWindow(int count) : base(
            className: $"{Program.AppId}.{typeof(PreviewWindow).Name}",
            title: $"{Program.AppId}.{typeof(PreviewWindow).Name}#{count}",
            style: WINDOW_STYLE.WS_POPUP,
            extendedStyle: WINDOW_EX_STYLE.WS_EX_NOACTIVATE)
        {
            DestroyOnDispose = true;
        }

        public async Task<WicBitmapSource?> Preview(Entry entry, PreviewHandlerOptions options)
        {
            //Application.TraceVerbose($"Previewing {entry.ParsingName}");
            var sw = Stopwatch.StartNew();
            using var handler = Handler.Create(entry);
            if (handler == null)
            {
                Application.TraceError($"Handler.Create failed for {handler}");
                return null;
            }

            var size = options.Size ?? PreviewHandlerOptions.GetDefaultSize();
            if (size.cx <= 0 || size.cy <= 0)
            {
                size = PreviewHandlerOptions.GetDefaultSize();
            }

            var phandler = handler.PreviewHandler;
            //ResizeAndMove(2560, 0, size.cx, size.cy);
            //Show();
            HRESULT hr;
            hr = phandler.Object.SetWindow(Handle, new RECT(POINT.Zero, size));
            if (hr.IsError)
            {
                Application.TraceError($"SetWindow failed for {handler}: {hr}");
                handler.Dispose();
                return null;
            }

            hr = phandler.Object.DoPreview();
            if (hr.IsError)
            {
                Application.TraceError($"DoPreview failed for {handler}: {hr}");
                handler.Dispose();
                return null;
            }

            // let the preview handler render during a bit of time (PDF with Adobe does that for example)
            var bmp = await Task.Run(() => CaptureGdi(options));
            //Application.TraceInfo($"Preview '{entry!.ParsingName}' ran {sw.Elapsed}");
            Dispose();
            return bmp;
        }

        private async Task<WicBitmapSource?> CaptureGdi(PreviewHandlerOptions options)
        {
            await Task.Delay(options.FinalTimeout);
            var hdc = DirectN.Functions.GetDC(Handle);
            if (hdc == 0)
            {
                Application.TraceError("GetDC failed.");
                return null;
            }

            var hdcMem = DirectN.Functions.CreateCompatibleDC(hdc);
            if (hdcMem == 0)
            {
                Application.TraceError("CreateCompatibleDC failed.");
                return null;
            }

            var rc = ClientRect;
            var bmp = DirectN.Functions.CreateCompatibleBitmap(hdc, rc.Width, rc.Height);
            if (DirectN.Functions.SelectObject(hdcMem, new(bmp.Value)) == 0)
            {
                Application.TraceError("SelectObject failed.");
                DirectN.Functions.DeleteObject(new(hdcMem));
                DirectN.Functions.DeleteObject(new(bmp));
                _ = DirectN.Functions.ReleaseDC(Handle, hdc);
                return null;
            }

            const int PW_RENDERFULLCONTENT = 2;
            DirectN.Functions.PrintWindow(Handle, hdcMem, PRINT_WINDOW_FLAGS.PW_CLIENTONLY | (PRINT_WINDOW_FLAGS)PW_RENDERFULLCONTENT);

            var wic = WicBitmapSource.FromHBITMAP(bmp);
            DirectN.Functions.DeleteObject(new(bmp));
            DirectN.Functions.DeleteObject(new(hdcMem));
            _ = DirectN.Functions.ReleaseDC(Handle, hdc);
            //Application.TraceInfo($"Bmp size:{wic?.Size.width} x {wic?.Size.height}.");
            return wic;
        }
    }

    [System.Runtime.InteropServices.Marshalling.GeneratedComClass]
    private sealed partial class Handler : IDisposable, IPreviewHandlerFrame//, ICustomQueryInterface
    {
        private readonly Entry _entry;
        private IComObject<IPreviewHandler>? _previewHandler;

        private Handler(Entry entry, IComObject<IPreviewHandler> previewHandler)
        {
            _entry = entry;
            _previewHandler = previewHandler;
            ThreadId = Environment.CurrentManagedThreadId;
        }

        public int ThreadId { get; }

        public IComObject<IPreviewHandler> PreviewHandler
        {
            get
            {
                var handler = _previewHandler;
                ObjectDisposedException.ThrowIf(handler == null, this);
                return handler;
            }
        }

        public static Handler? Create(Entry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);
            var hr = entry.Extension.TryCreatePreviewHandler(out var obj);
            if (hr.IsError)
            {
                Application.TraceError($"TryCreatePreviewHandler failed for '{entry.ParsingName}': {hr}");
                return null;
            }

            if (obj == null)
            {
                Application.TraceError($"There's no preview handler for '{entry.ParsingName}'");
                return null;
            }

            var handler = new Handler(entry, obj);
            hr = handler.Initialize(entry);
            if (hr.IsError)
            {
                Application.TraceError($"Initialize failed for '{entry.ParsingName}': {hr}");
                handler.Dispose();
                return null;
            }
            return handler;
        }

        public override string ToString() => $"{_entry.Extension} ({_entry.Extension.PreviewHandlerClsid:B}) '{_entry?.ParsingName}'";

        private HRESULT Initialize(Entry entry)
        {
            // preferred and recommended but not always supported
            var iws = PreviewHandler.As<IInitializeWithStream>();
            if (iws == null)
            {
                // the most common initialization interface
                var iwf = PreviewHandler.As<IInitializeWithFile>();
                if (iwf == null)
                {
                    // .msg files and others previewers only support IInitializeWithItem
                    var iwi = PreviewHandler.As<IInitializeWithItem>();
                    if (iwi == null)
                        return DirectN.Constants.E_NOINTERFACE;

                    var item = Entry.GetItem(entry.ParsingName, ShellItemParsingOptions.DontThrowOnError);
                    if (item == null)
                        return DirectN.Constants.E_FAIL;

                    var hr = iwi.Object.Initialize(item.NativeObject, (uint)STGM.STGM_READ);
                    if (hr.IsError)
                        return hr;
                }
                else
                {
                    var hr = iwf.Object.Initialize(PWSTR.From(entry.ParsingName), (uint)STGM.STGM_READ);
                    if (hr.IsError)
                        return hr;
                }
            }
            else
            {
                var stream = entry.OpenRead();
                if (stream == null)
                    return DirectN.Constants.WINCODEC_ERR_STREAMNOTAVAILABLE; // not really the error code but close enough

                var mis = new ManagedIStream(stream, true);
                var hr = iws.Object.Initialize(mis, (uint)STGM.STGM_READ);
                if (hr.IsError)
                    return hr;
            }

            var ows = PreviewHandler.As<IObjectWithSite>();
            if (ows != null)
            {
                ComObject.WithComInstance(this, unk => ows?.Object.SetSite(unk), true);
            }
            return DirectN.Constants.S_OK;
        }

        public void Dispose()
        {
            var handler = Interlocked.Exchange(ref _previewHandler, null);
            if (handler != null)
            {
                var ows = handler.As<IObjectWithSite>();
                if (ows != null)
                {
                    ComObject.WithComInstance(this, unk => ows?.Object.SetSite(0), true);
                }
                handler.Dispose();
            }
        }

        HRESULT IPreviewHandlerFrame.GetWindowContext(out PREVIEWHANDLERFRAMEINFO pinfo)
        {
            //Application.TraceInfo("GetWindowContext called");
            pinfo = new PREVIEWHANDLERFRAMEINFO();
            return DirectN.Constants.S_OK;
        }

        HRESULT IPreviewHandlerFrame.TranslateAccelerator(in MSG pmsg)
        {
            //Application.TraceInfo($"msg:{MessageDecoder.Decode(pmsg)}");
            return DirectN.Constants.S_FALSE;
        }

        //public CustomQueryInterfaceResult GetInterface(ref Guid iid, out nint ppv)
        //{
        //    var process = ComObject.GetCallContextProcess(false);

        //    Application.TraceInfo($"iid: {iid:B} process:'{process?.ProcessName}'");
        //    ppv = 0;
        //    return CustomQueryInterfaceResult.NotHandled;
        //}
    }
}
