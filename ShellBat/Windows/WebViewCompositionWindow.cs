namespace ShellBat.Windows;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class WebViewCompositionWindow : CompositionWindow, IDropTarget
{
    private readonly bool[] _capturedButtons = new bool[Enum.GetNames<MouseButton>().Length];
    private readonly HashSet<uint> _pointerIdsStartingInWebView = [];
    private readonly Dictionary<ulong, NavigationEventArgs> _navigationEvents = [];
    private IComObject<ICoreWebView2CompositionController>? _controller;
    private IComObject<ICoreWebView2CompositionController3>? _controller3;
    private ComObject<ICoreWebView2Environment3>? _environment;
    private IComObject<ICoreWebView2>? _webView;
    private CoreWebView2CompositionControllerEvents? _controllerEvents;
    private CoreWebView2Events? _webViewEvents;
    private bool _mouseTracking;
    private bool _isDropTarget;

    public event EventHandler<MouseEventArgs>? MouseMove;
    public event EventHandler<MouseEventArgs>? MouseLeave;
    public event EventHandler<MouseEventArgs>? MouseHover;
    public event EventHandler<MouseWheelEventArgs>? MouseWheel;
    public event EventHandler<MouseButtonEventArgs>? MouseButtonDown;
    public event EventHandler<MouseButtonEventArgs>? MouseButtonUp;
    public event EventHandler<MouseButtonEventArgs>? MouseButtonDoubleClick;
    public event EventHandler<PointerEnterEventArgs>? PointerEnter;
    public event EventHandler<PointerLeaveEventArgs>? PointerLeave;
    public event EventHandler<PointerWheelEventArgs>? PointerWheel;
    public event EventHandler<PointerPositionEventArgs>? PointerUpdate;
    public event EventHandler<PointerContactChangedEventArgs>? PointerContactChanged;
    public event EventHandler<KeyEventArgs>? KeyDown;
    public event EventHandler<KeyEventArgs>? KeyUp;
    public event EventHandler<KeyPressEventArgs>? KeyPress;
    public event EventHandler? MonitorChanged;
    public event EventHandler<NavigationEventArgs>? NavigationStarting;
    public event EventHandler<NavigationEventArgs>? NavigationCompleted;

    public WebViewCompositionWindow(
        string? title = null,
        WINDOW_STYLE style = WINDOW_STYLE.WS_THICKFRAME,
        WINDOW_EX_STYLE extendedStyle = WINDOW_EX_STYLE.WS_EX_NOREDIRECTIONBITMAP,
        RECT? rect = null)
        : base(title, style: style, extendedStyle: extendedStyle, rect: rect)
    {
        MonitorHandle = DirectN.Functions.MonitorFromWindow(Handle, MONITOR_FROM_FLAGS.MONITOR_DEFAULTTONULL);
        if (IsFullScreen)
        {
            SetCorner(DWM_WINDOW_CORNER_PREFERENCE.DWMWCP_DONOTROUND);
        }
        else
        {
            SetCorner(ShellBatInstance.Current.Settings.WindowCorner);
        }
        var options = GetEnvironmentOptions();
        using var userDataPathStr = new Pwstr(Settings.WebView2UserDataPath);
        WebView2.Functions.CreateCoreWebView2EnvironmentWithOptions(PWSTR.Null, userDataPathStr, options!,
            new CoreWebView2CreateCoreWebView2EnvironmentCompletedHandler((result, envObj) =>
            {
                options?.Dispose();
                var env3 = (ICoreWebView2Environment3)envObj;
                _environment = new ComObject<ICoreWebView2Environment3>(env3);
                _ = CreateControllerAsync(env3);
            }));
    }

    private async Task CreateControllerAsync(ICoreWebView2Environment3 environment)
    {
        try
        {
            _controller = await environment.CreateCoreWebView2CompositionControllerAsync(Handle) ?? throw new InvalidOperationException("The WebView2 composition controller could not be created.");
            _controller3 = ComExtensions.As<ICoreWebView2CompositionController3>(_controller);
            _controllerEvents = new CoreWebView2CompositionControllerEvents(_controller);
            _controllerEvents.CursorChanged += (sender, args) =>
            {
                if (CanChangeCursor && sender is ICoreWebView2CompositionController controller)
                {
                    Cursor = controller.Cursor;
                }
            };

            _controller.RootVisualTarget = RootVisual;

            var ctrl = (ICoreWebView2Controller)_controller.Object;
            ctrl.Bounds = ClientRect;
            _webView = ctrl.CoreWebView2 ?? throw new InvalidOperationException("The WebView2 controller has no WebView.");

            _webViewEvents = new CoreWebView2Events(_webView);
            _webViewEvents.NavigationStarting += (sender, args) =>
            {
                var id = args.NavigationId;
                var e = new NavigationEventArgs(id, args.Uri ?? string.Empty, args.IsUserInitiated, args.IsRedirected);
                _navigationEvents[id] = e;

                OnNavigationStarting(this, e);
                if (e.Cancel)
                {
                    args.Cancel = true;
                }
            };

            _webViewEvents.NavigationCompleted += (sender, args) =>
            {
                var id = args.NavigationId;
                if (_navigationEvents.TryGetValue(id, out var e))
                {
                    e.Type = NavigationEventType.NavigationCompleted;
                    e.WebErrorStatus = args.WebErrorStatus;
                    e.IsSuccess = args.IsSuccess;

                    _navigationEvents.Remove(id);
                    OnNavigationCompleted(this, e);
                }
            };

            ControllerCreated();
        }
        catch (Exception ex)
        {
            Application.AddError(ex);
        }
    }

    protected IComObject<ICoreWebView2CompositionController>? Controller => _controller;
    protected IComObject<ICoreWebView2>? WebView => _webView;
    protected ComObject<ICoreWebView2Environment3>? Environment => _environment;

    public HMONITOR MonitorHandle { get; private set; }
    public bool IsFullScreen => GetFullScreenBounds() == WindowRect;
    public virtual bool CanChangeCursor { get; set; }
    public bool IsDropTarget
    {
        get => _isDropTarget;
        set
        {
            if (value == _isDropTarget)
                return;

            if (value)
            {
                // we need to ensure this as STAThread doesn't always call it for some reason
                DirectN.Functions.OleInitialize(0); // don't check error
                var hr = DirectN.Functions.RegisterDragDrop(Handle, this);
                if (hr.IsError && hr != DirectN.Constants.DRAGDROP_E_ALREADYREGISTERED)
                    throw new Exception("Cannot enable drag & drop operations. Make sure the thread is initialized as an STA thread.", Marshal.GetExceptionForHR((int)hr)!);

                _isDropTarget = true;
            }
            else
            {
                var hr = DirectN.Functions.RevokeDragDrop(Handle);
                hr.ThrowOnErrorExcept(DirectN.Constants.DRAGDROP_E_NOTREGISTERED);
                _isDropTarget = false;
            }
        }
    }

    protected virtual CoreWebView2EnvironmentOptions? GetEnvironmentOptions() => null;
    protected virtual RECT? GetCaptionRect() => null;
    protected virtual void ControllerCreated()
    {
    }

    public virtual RECT GetFullScreenBounds()
    {
        var monitor = GetMonitor(MONITOR_FROM_FLAGS.MONITOR_DEFAULTTONEAREST)!;
        var bounds = monitor.Bounds;
        bounds = bounds.Inflate(1, 1, 1, 1); // work around a weird 1px gap that would otherwise appear on the right and bottom edges, probably due to some rounding issue in the DWM when using the exact monitor size
        return bounds;
    }

    protected override bool OnFocusChanged(bool setOrKill)
    {
        if (setOrKill)
        {
            _controller?.As<ICoreWebView2Controller>()?.Object.MoveFocus(COREWEBVIEW2_MOVE_FOCUS_REASON.COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC);
            return true;
        }
        return base.OnFocusChanged(setOrKill);
    }

    public virtual Task<T?> ExecuteScript<T>(string script, JsonTypeInfo<T> typeInfo, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(script);
        var webView = _webView ?? throw new InvalidOperationException();
        return webView.ExecuteScript(script, typeInfo, throwOnError: throwOnError);
    }

    public virtual Task<string?> ExecuteScriptAsJson(string script, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(script);
        var webView = _webView ?? throw new InvalidOperationException();
        return webView.ExecuteScriptAsJson(script, throwOnError);
    }

    public virtual HRESULT ExecuteScript(string script, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(script);
        var webView = _webView ?? throw new InvalidOperationException();
        return webView.ExecuteScript(script, throwOnError);
    }

    protected virtual void OnPointerWheel(object? sender, PointerWheelEventArgs e) => PointerWheel?.Invoke(sender, e);
    protected virtual void OnPointerLeave(object? sender, PointerLeaveEventArgs e) => PointerLeave?.Invoke(sender, e);
    protected virtual void OnPointerEnter(object? sender, PointerEnterEventArgs e) => PointerEnter?.Invoke(sender, e);
    protected virtual void OnPointerUpdate(object? sender, PointerPositionEventArgs e) => PointerUpdate?.Invoke(sender, e);
    protected virtual void OnPointerContactChanged(object? sender, PointerContactChangedEventArgs e) => PointerContactChanged?.Invoke(sender, e);
    protected virtual void OnMouseMove(object? sender, MouseEventArgs e) => MouseMove?.Invoke(sender, e);
    protected virtual void OnMouseLeave(object? sender, MouseEventArgs e) => MouseLeave?.Invoke(sender, e);
    protected virtual void OnMouseHover(object? sender, MouseEventArgs e) => MouseHover?.Invoke(sender, e);
    protected virtual void OnMouseWheel(object? sender, MouseWheelEventArgs e) => MouseWheel?.Invoke(sender, e);
    protected virtual void OnMouseButtonDown(object? sender, MouseButtonEventArgs e) => MouseButtonDown?.Invoke(sender, e);
    protected virtual void OnMouseButtonUp(object? sender, MouseButtonEventArgs e) => MouseButtonUp?.Invoke(sender, e);
    protected virtual void OnMouseButtonDoubleClick(object? sender, MouseButtonEventArgs e) => MouseButtonDoubleClick?.Invoke(sender, e);
    protected virtual void OnKeyDown(object? sender, KeyEventArgs e) => KeyDown?.Invoke(sender, e);
    protected virtual void OnKeyUp(object? sender, KeyEventArgs e) => KeyUp?.Invoke(sender, e);
    protected virtual void OnKeyPress(object? sender, KeyPressEventArgs e) => KeyPress?.Invoke(sender, e);
    protected virtual void OnMonitorChanged(object? sender, EventArgs e) => MonitorChanged?.Invoke(sender, e);
    protected virtual void OnNavigationStarting(object? sender, NavigationEventArgs e) => NavigationStarting?.Invoke(sender, e);
    protected virtual void OnNavigationCompleted(object? sender, NavigationEventArgs e) => NavigationCompleted?.Invoke(sender, e);

    protected override void OnExitSizeMove(object? sender, EventArgs e)
    {
        base.OnExitSizeMove(sender, e);
        UpdateMonitor();
    }

    protected override void OnPositionChanged(object? sender, ValueEventArgs<WINDOWPOS> e)
    {
        base.OnPositionChanged(sender, e);
        UpdateMonitor();
    }

    protected override void OnPositionChanging(object? sender, ValueEventArgs<WINDOWPOS> e)
    {
        base.OnPositionChanging(sender, e);
        UpdateMonitor();
    }

    protected virtual void UpdateMonitor()
    {
        var monitor = DirectN.Functions.MonitorFromWindow(Handle, MONITOR_FROM_FLAGS.MONITOR_DEFAULTTONULL);
        if (monitor.Value == MonitorHandle.Value)
            return;

        MonitorHandle = monitor;
        OnMonitorChanged(this, EventArgs.Empty);
    }

    protected unsafe internal virtual void SetCorner(DWM_WINDOW_CORNER_PREFERENCE corner)
    {
        // works only on Windows 11, does nothing on Windows 10, so we don't check error
        DirectN.Functions.DwmSetWindowAttribute(Handle, (uint)DWMWINDOWATTRIBUTE.DWMWA_WINDOW_CORNER_PREFERENCE, (nint)(&corner), 4);
    }

    protected virtual void OnAfterDragEnter(IDataObject dataObject, MODIFIERKEYS_FLAGS flags, POINTL point, DROPEFFECT effect) { }
    protected virtual HRESULT OnBeforeDragEnter(IDataObject dataObject, MODIFIERKEYS_FLAGS flags, POINTL point, ref DROPEFFECT effect, out bool handled)
    {
        handled = false;
        return DirectN.Constants.S_OK;
    }

    protected virtual void OnAfterDragOver(MODIFIERKEYS_FLAGS flags, POINTL point, DROPEFFECT effect) { }
    protected virtual HRESULT OnBeforeDragOver(MODIFIERKEYS_FLAGS flags, POINTL point, ref DROPEFFECT effect, out bool handled)
    {
        handled = false;
        return DirectN.Constants.S_OK;
    }

    protected virtual void OnAfterDragLeave() { }
    protected virtual HRESULT OnBeforeDragLeave(out bool handled)
    {
        handled = false;
        return DirectN.Constants.S_OK;
    }

    protected virtual void OnAfterDrop(IDataObject dataObject, MODIFIERKEYS_FLAGS flags, POINTL point, DROPEFFECT effect) { }
    protected virtual HRESULT OnBeforeDrop(IDataObject dataObject, MODIFIERKEYS_FLAGS flags, POINTL point, DROPEFFECT effect, out bool handled)
    {
        handled = false;
        return DirectN.Constants.S_OK;
    }

    protected virtual void ClearBrowsingDataAll()
    {
        if (_webView?.Object is not ICoreWebView2_13)
            return;

        using var profile = _webView.Profile;
        if (profile?.Object is not ICoreWebView2Profile2)
            return;

        _ = profile.ClearBrowsingDataAllAsync();
    }

    private void OnMouseMove(MouseEventArgs e)
    {
        OnMouseMove(this, e);
        if (e.Handled)
            return;

        var keys = WindowsExtensions.GetKeys(e.Keys, null);
        _controller?.SendMouseInput(COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_MOVE, keys, 0, e.Point);
    }

    private void OnMouseLeave(MouseEventArgs e)
    {
        OnMouseLeave(this, e);
        if (e.Handled)
            return;

        _controller?.SendMouseInput(
            COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_LEAVE,
            COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_NONE,
            0,
            POINT.Zero);
    }

    private void OnMouseButtonDown(MouseButtonEventArgs e)
    {
        OnMouseButtonDown(this, e);
        if (e.Handled)
            return;

        var keys = WindowsExtensions.GetKeys(e.Keys, e.Button);
        var kind = e.Button.GetKind(WindowsExtensions.ButtonAction.Down);
        _controller?.SendMouseInput(kind, keys, e.Button == MouseButton.X1 ? 1u : e.Button == MouseButton.X2 ? 2u : 0, e.Point);
    }

    private void OnMouseButtonUp(MouseButtonEventArgs e)
    {
        OnMouseButtonUp(this, e);
        if (e.Handled)
            return;

        var keys = WindowsExtensions.GetKeys(e.Keys, e.Button);
        var kind = e.Button.GetKind(WindowsExtensions.ButtonAction.Up);
        _controller?.SendMouseInput(kind, keys, e.Button == MouseButton.X1 ? 1u : e.Button == MouseButton.X2 ? 2u : 0, e.Point);
    }

    private void OnMouseButtonDoubleClick(MouseButtonEventArgs e)
    {
        OnMouseButtonDoubleClick(this, e);
        if (e.Handled)
            return;

        var keys = WindowsExtensions.GetKeys(e.Keys, e.Button);
        var kind = e.Button.GetKind(WindowsExtensions.ButtonAction.DoubleClick);
        _controller?.SendMouseInput(kind, keys, e.Button == MouseButton.X1 ? 1u : e.Button == MouseButton.X2 ? 2u : 0, e.Point);
    }

    private void OnMouseWheel(MouseWheelEventArgs e)
    {
        OnMouseWheel(this, e);
        if (e.Handled)
            return;

        var keys = WindowsExtensions.GetKeys(e.Keys, null);
        var kind = e.Orientation == Orientation.Horizontal
            ? COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_HORIZONTAL_WHEEL
            : COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_WHEEL;
        _controller?.SendMouseInput(kind, keys, (uint)(e.Delta * DirectN.Constants.WHEEL_DELTA), e.Point);
    }

    protected override LRESULT? WindowProc(HWND hwnd, uint msg, WPARAM wParam, LPARAM lParam)
    {
        //if (msg != MessageDecoder.WM_SETCURSOR && msg != MessageDecoder.WM_MOUSEMOVE &&
        //    msg != MessageDecoder.WM_NCHITTEST && msg != MessageDecoder.WM_NCMOUSEMOVE && msg != MessageDecoder.WM_GETICON)
        //{
        //    Application.TraceInfo("msg: " + MessageDecoder.Decode(hwnd, msg, wParam, lParam));
        //}

        MouseButton button;
        switch (msg)
        {
            //https://learn.microsoft.com/en-us/windows/win32/dwm/customframe#appendix-c-hittestnca-function
            case MessageDecoder.WM_ACTIVATE:

                MARGINS margins;
                margins.cxLeftWidth = -1;
                margins.cxRightWidth = -1;
                margins.cyBottomHeight = -1;
                margins.cyTopHeight = -1;
                DirectN.Functions.DwmExtendFrameIntoClientArea(Handle, margins);
                break;

            case MessageDecoder.WM_NCCALCSIZE:
                if (wParam.Value.ToUInt32() != 0)
                {
                    if (IsZoomed)
                    {
                        var monitor = GetMonitor();
                        if (monitor != null)
                        {
                            unsafe
                            {
                                // this is a NCCALCSIZE_PARAMS but we only need the first RECT
                                *(RECT*)lParam.Value = monitor.WorkingArea;
                            }
                        }
                    }
                    return 0;
                }
                break;

            case MessageDecoder.WM_CREATE:
                var rc = WindowRect;
                SetWindowPos(HWND.Null, rc.left, rc.top, rc.Width, rc.Height, SET_WINDOW_POS_FLAGS.SWP_FRAMECHANGED);
                break;

            case MessageDecoder.WM_NCHITTEST:
                var htn = HitTestWithoutWindowsFrame();
                if (htn.HasValue)
                    return new LRESULT((nint)htn.Value);

                HT? HitTestWithoutWindowsFrame()
                {
                    var rc = WindowRect;
                    var ncx = lParam.Value.SignedLOWORD();
                    var ncy = lParam.Value.SignedHIWORD();
                    var clix = ncx - rc.left;
                    var cliy = ncy - rc.top;
                    if (clix >= 0 && clix < rc.Width && cliy >= 0 && cliy <= rc.Height)
                    {
                        if (clix < WindowsExtensions.BorderWidth)
                        {
                            if (cliy <= WindowsExtensions.BorderHeight)
                                return HT.HTTOPLEFT;

                            if (cliy >= rc.Height - WindowsExtensions.BorderHeight)
                                return HT.HTBOTTOMLEFT;

                            return HT.HTLEFT;
                        }

                        if (clix > rc.Width - WindowsExtensions.BorderWidth)
                        {
                            if (cliy <= WindowsExtensions.BorderHeight)
                                return HT.HTTOPRIGHT;

                            if (cliy >= rc.Height - WindowsExtensions.BorderHeight)
                                return HT.HTBOTTOMRIGHT;

                            return HT.HTRIGHT;
                        }

                        if (cliy < WindowsExtensions.BorderHeight)
                            return HT.HTTOP;

                        if (cliy > rc.Height - WindowsExtensions.BorderHeight)
                            return HT.HTBOTTOM;

                        var caption = GetCaptionRect();
                        if (caption != null && caption.Value.Contains(clix, cliy))
                            return HT.HTCAPTION;
                    }
                    return null;
                }
                break;

            case MessageDecoder.WM_MOUSEMOVE:
                if (!_mouseTracking)
                {
                    unsafe
                    {
                        // https://learn.microsoft.com/en-us/windows/win32/learnwin32/other-mouse-operations#mouse-tracking-events-hover-and-leave
                        var tme = new TRACKMOUSEEVENT
                        {
                            cbSize = (uint)sizeof(TRACKMOUSEEVENT),
                            dwFlags = TRACKMOUSEEVENT_FLAGS.TME_LEAVE | TRACKMOUSEEVENT_FLAGS.TME_HOVER,
                            hwndTrack = hwnd,
                        };
                        _mouseTracking = DirectN.Functions.TrackMouseEvent(ref tme);
                    }
                }

                OnMouseMove(new MouseEventArgs(lParam.ToPOINT(), (MODIFIERKEYS_FLAGS)wParam.Value.LOWORD()));
                break;

            case MessageDecoder.WM_MOUSELEAVE:
                _mouseTracking = false;
                OnMouseLeave(new MouseEventArgs(POINT.Zero, 0));
                return 0;

            case MessageDecoder.WM_MOUSEHOVER:
                OnMouseHover(this, new MouseEventArgs(lParam.ToPOINT(), (MODIFIERKEYS_FLAGS)wParam.Value.LOWORD()));
                return 0;

            case MessageDecoder.WM_LBUTTONDOWN:
            case MessageDecoder.WM_RBUTTONDOWN:
            case MessageDecoder.WM_MBUTTONDOWN:
            case MessageDecoder.WM_XBUTTONDOWN:
                button = WindowsExtensions.MessageToButton(msg, wParam);
                _capturedButtons[(int)button] = true;
                DirectN.Functions.SetCapture(hwnd);
                OnMouseButtonDown(new MouseButtonEventArgs(lParam.ToPOINT(), (MODIFIERKEYS_FLAGS)wParam.Value.LOWORD(), button));
                break;

            case MessageDecoder.WM_LBUTTONUP:
            case MessageDecoder.WM_RBUTTONUP:
            case MessageDecoder.WM_MBUTTONUP:
            case MessageDecoder.WM_XBUTTONUP:
                button = WindowsExtensions.MessageToButton(msg, wParam);
                _capturedButtons[(int)button] = false;
                DirectN.Functions.ReleaseCapture();
                OnMouseButtonUp(new MouseButtonEventArgs(lParam.ToPOINT(), (MODIFIERKEYS_FLAGS)wParam.Value.LOWORD(), button));
                break;

            case MessageDecoder.WM_LBUTTONDBLCLK:
            case MessageDecoder.WM_RBUTTONDBLCLK:
            case MessageDecoder.WM_MBUTTONDBLCLK:
            case MessageDecoder.WM_XBUTTONDBLCLK:
                button = WindowsExtensions.MessageToButton(msg, wParam);
                var e3 = new MouseButtonEventArgs(lParam.ToPOINT(), (MODIFIERKEYS_FLAGS)wParam.Value.LOWORD(), button);
                OnMouseButtonDoubleClick(e3);
                break;

            case MessageDecoder.WM_MOUSEHWHEEL:
            case MessageDecoder.WM_MOUSEWHEEL:
                OnMouseWheel(new MouseWheelEventArgs(
                    lParam.ToPOINT().ScreenToClient(hwnd),
                    (MODIFIERKEYS_FLAGS)wParam.Value.LOWORD(),
                    wParam.Value.SignedHIWORD(),
                    msg == MessageDecoder.WM_MOUSEHWHEEL ? Orientation.Horizontal : Orientation.Vertical));
                break;

            case MessageDecoder.WM_POINTERACTIVATE:
            case MessageDecoder.WM_POINTERDOWN:
            case MessageDecoder.WM_POINTERENTER:
            case MessageDecoder.WM_POINTERLEAVE:
            case MessageDecoder.WM_POINTERUP:
            case MessageDecoder.WM_POINTERUPDATE:
                if (TryForwardPointerInput(msg, wParam, lParam))
                    return 0;
                break;

            case MessageDecoder.WM_CHAR:
            case MessageDecoder.WM_SYSCHAR:
                var e = new KeyPressEventArgs(wParam.Value.ToUInt32());
                OnKeyPress(this, e);
                if (e.Handled)
                    return null;

                break;

            case MessageDecoder.WM_KEYDOWN:
            case MessageDecoder.WM_KEYUP:
            case MessageDecoder.WM_SYSKEYDOWN:
            case MessageDecoder.WM_SYSKEYUP:
                var vk = (VIRTUAL_KEY)wParam.Value.ToUInt32();
                var e2 = new KeyEventArgs(vk, (uint)lParam.Value.ToInt64(), ((char)DirectN.Functions.MapVirtualKeyW((uint)vk, MAP_VIRTUAL_KEY_TYPE.MAPVK_VK_TO_CHAR)).ToString());
                if (e2.IsUp)
                {
                    OnKeyUp(this, e2);
                }
                else
                {
                    OnKeyDown(this, e2);
                }
                if (e2.Handled)
                    return null;

                break;
        }
        return base.WindowProc(hwnd, msg, wParam, lParam);
    }

    // from https://github.com/MicrosoftEdge/WebView2Samples/blob/main/SampleApps/WebView2APISample/ViewComponent.cpp
    protected virtual bool TryForwardPointerInput(uint msg, WPARAM wParam, LPARAM lParam)
    {
        if (Controller == null)
            return false;

        var pointerId = wParam.GetPointerId();
        var point = lParam.ToPOINT().ScreenToClient(Handle);
        var pointerStartedInWebView = _pointerIdsStartingInWebView.Contains(pointerId);
        if (!pointerStartedInWebView && !ClientRect.Contains(point))
            return false;

        if (!pointerStartedInWebView && (msg == MessageDecoder.WM_POINTERENTER || msg == MessageDecoder.WM_POINTERDOWN))
        {
            _pointerIdsStartingInWebView.Add(pointerId);
        }
        else if (msg == MessageDecoder.WM_POINTERLEAVE)
        {
            _pointerIdsStartingInWebView.Remove(pointerId);
        }

        var ctrl4 = Controller.As<ICoreWebView2ExperimentalCompositionController4>();
        if (ctrl4 == null)
            return false;

        var matrix = D2D_MATRIX_4X4_F.Identity();
        // this is needed to adjust pointer coordinates from screen to webview's root visual target, which may be different if the window is moved, etc.
        //matrix._41 += webViewBounds left
        //matrix._42 += m_webViewBounds top
        if (ctrl4.Object.CreateCoreWebView2PointerInfoFromPointerId(pointerId, Handle, matrix, out var infoObj).IsError)
            return false;

        var info = new ComObject<ICoreWebView2PointerInfo>(infoObj);
        Controller.Object.SendPointerInput((COREWEBVIEW2_POINTER_EVENT_KIND)msg, info.Object).ThrowOnError();
        return true;
    }

    protected override bool OnMoving(ref RECT rc)
    {
        if (_controller?.Object is ICoreWebView2Controller c)
        {
            c.NotifyParentWindowPositionChanged().ThrowOnError();
        }
        return base.OnMoving(ref rc);
    }

    protected override bool OnMoved()
    {
        if (_controller?.Object is ICoreWebView2Controller c)
        {
            c.NotifyParentWindowPositionChanged().ThrowOnError();
        }
        return base.OnMoved();
    }

    protected override bool OnResized(WindowResizedType type, SIZE size)
    {
        if (_controller?.Object is ICoreWebView2Controller c)
        {
            c.Bounds = ClientRect;
        }
        return base.OnResized(type, size);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            Interlocked.Exchange(ref _controllerEvents, null)?.Dispose();
            Interlocked.Exchange(ref _webViewEvents, null)?.Dispose();

            var controller = Interlocked.Exchange(ref _controller, null);
            if (controller?.Object is ICoreWebView2Controller coreController)
            {
                coreController.Close();
            }

            _environment?.Dispose();
            controller?.Dispose();
        }
        base.Dispose(disposing);
    }

    HRESULT IDropTarget.DragEnter(IDataObject pDataObj, MODIFIERKEYS_FLAGS grfKeyState, POINTL pt, ref DROPEFFECT pdwEffect)
    {
        if (_controller3 == null)
            return DirectN.Constants.E_NOTIMPL;

        var hr = OnBeforeDragEnter(pDataObj, grfKeyState, pt, ref pdwEffect, out var handled);
        if (hr.IsError)
            return hr;

        // note: if handled is true, we must not call other DragEnter/DragOver/Drop methods on the webview2 or
        // it will mess up the internal state and we'll get errors like this:
        // 0x8007139F: The group or resource is not in the correct state to perform the requested operation.
        if (handled)
            return DirectN.Constants.S_OK;

        var effect = (uint)pdwEffect;
        hr = _controller3.Object.DragEnter(pDataObj, (uint)grfKeyState, ScreenToClient(new POINT(pt.x, pt.y)), ref effect);
        if (hr.IsSuccess)
        {
            pdwEffect = (DROPEFFECT)effect;
            OnAfterDragEnter(pDataObj, grfKeyState, pt, pdwEffect);
        }
        return hr;
    }

    HRESULT IDropTarget.DragOver(MODIFIERKEYS_FLAGS grfKeyState, POINTL pt, ref DROPEFFECT pdwEffect)
    {
        if (_controller3 == null)
            return DirectN.Constants.E_NOTIMPL;

        var hr = OnBeforeDragOver(grfKeyState, pt, ref pdwEffect, out var handled);
        if (hr.IsError)
            return hr;

        if (handled)
            return DirectN.Constants.S_OK;

        var effect = (uint)pdwEffect;
        hr = _controller3.Object.DragOver((uint)grfKeyState, ScreenToClient(new POINT(pt.x, pt.y)), ref effect);
        if (hr.IsSuccess)
        {
            pdwEffect = (DROPEFFECT)effect;
            OnAfterDragOver(grfKeyState, pt, pdwEffect);
        }
        return hr;
    }

    HRESULT IDropTarget.DragLeave()
    {
        if (_controller3 == null)
            return DirectN.Constants.E_NOTIMPL;

        var hr = OnBeforeDragLeave(out var handled);
        if (hr.IsError)
            return hr;

        if (handled)
            return DirectN.Constants.S_OK;

        hr = _controller3.Object.DragLeave();
        if (hr.IsSuccess)
        {
            OnAfterDragLeave();
        }
        return hr;
    }

    HRESULT IDropTarget.Drop(IDataObject pDataObj, MODIFIERKEYS_FLAGS grfKeyState, POINTL pt, ref DROPEFFECT pdwEffect)
    {
        if (_controller3 == null)
            return DirectN.Constants.E_NOTIMPL;

        var hr = OnBeforeDrop(pDataObj, grfKeyState, pt, pdwEffect, out var handled);
        if (hr.IsError)
            return hr;

        if (handled)
            return DirectN.Constants.S_OK;

        var effect = (uint)pdwEffect;
        hr = _controller3.Object.Drop(pDataObj, (uint)grfKeyState, ScreenToClient(new POINT(pt.x, pt.y)), ref effect);
        if (hr.IsSuccess)
        {
            pdwEffect = (DROPEFFECT)effect;
            OnAfterDrop(pDataObj, grfKeyState, pt, pdwEffect);
        }
        return hr;
    }
}
