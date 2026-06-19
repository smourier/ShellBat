namespace ShellBat.Windows;

public static class WindowsExtensions
{
    private readonly static Lazy<int> _borderWidth = new(() => DirectN.Functions.GetSystemMetrics(SYSTEM_METRICS_INDEX.SM_CXFRAME) + DirectN.Functions.GetSystemMetrics(SYSTEM_METRICS_INDEX.SM_CXPADDEDBORDER));
    public static int BorderWidth => _borderWidth.Value;

    private readonly static Lazy<int> _borderHeight = new(() => DirectN.Functions.GetSystemMetrics(SYSTEM_METRICS_INDEX.SM_CYFRAME) + DirectN.Functions.GetSystemMetrics(SYSTEM_METRICS_INDEX.SM_CXPADDEDBORDER));
    public static int BorderHeight => _borderHeight.Value;

    private readonly static Lazy<int> _buttonWidth = new(() => DirectN.Functions.GetSystemMetrics(SYSTEM_METRICS_INDEX.SM_CXSIZE));
    public static int ButtonWidth => _buttonWidth.Value;

    private readonly static Lazy<int> _buttonHeight = new(() => DirectN.Functions.GetSystemMetrics(SYSTEM_METRICS_INDEX.SM_CYSIZE));
    public static int ButtonHeight => _buttonHeight.Value;

    public static POINT ToPOINT(this LPARAM lParam) => new(lParam.Value.SignedLOWORD(), lParam.Value.SignedHIWORD());
    public static POINT ScreenToClient(this POINT pt, HWND hwnd) { DirectN.Functions.ScreenToClient(hwnd, ref pt); return pt; }
    public static POINT ClientToScreen(this POINT pt, HWND hwnd) { DirectN.Functions.ClientToScreen(hwnd, ref pt); return pt; }

    public static COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS GetKeys(POINTER_MOD vk, MouseButton? button) => GetKeys((MODIFIERKEYS_FLAGS)vk, button);
    public static COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS GetKeys(MODIFIERKEYS_FLAGS vk, MouseButton? button)
    {
        var keys = COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_NONE;
        if (vk.HasFlag(MODIFIERKEYS_FLAGS.MK_CONTROL))
        {
            keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_CONTROL;
        }

        if (vk.HasFlag(MODIFIERKEYS_FLAGS.MK_SHIFT))
        {
            keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_SHIFT;
        }

        if (vk.HasFlag(MODIFIERKEYS_FLAGS.MK_LBUTTON))
        {
            keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_LEFT_BUTTON;
        }

        if (vk.HasFlag(MODIFIERKEYS_FLAGS.MK_RBUTTON))
        {
            keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_RIGHT_BUTTON;
        }

        if (vk.HasFlag(MODIFIERKEYS_FLAGS.MK_MBUTTON))
        {
            keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_MIDDLE_BUTTON;
        }

        if (vk.HasFlag(MODIFIERKEYS_FLAGS.MK_XBUTTON1))
        {
            keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_X_BUTTON1;
        }

        if (vk.HasFlag(MODIFIERKEYS_FLAGS.MK_XBUTTON2))
        {
            keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_X_BUTTON2;
        }

        if (button != null)
        {
            switch (button.Value)
            {
                case MouseButton.Left:
                    keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_LEFT_BUTTON;
                    break;

                case MouseButton.Right:
                    keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_RIGHT_BUTTON;
                    break;

                case MouseButton.Middle:
                    keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_MIDDLE_BUTTON;
                    break;

                case MouseButton.X1:
                    keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_X_BUTTON1;
                    break;

                case MouseButton.X2:
                    keys |= COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS.COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_X_BUTTON2;
                    break;
            }
        }
        return keys;
    }

    public enum ButtonAction
    {
        Down,
        Up,
        DoubleClick,
    }

    public static COREWEBVIEW2_MOUSE_EVENT_KIND GetKind(this MouseButton button, ButtonAction action)
    {
        switch (button)
        {
            case MouseButton.Left:
                switch (action)
                {
                    case ButtonAction.Down: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_LEFT_BUTTON_DOWN;
                    case ButtonAction.Up: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_LEFT_BUTTON_UP;
                    case ButtonAction.DoubleClick: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_LEFT_BUTTON_DOUBLE_CLICK;
                }
                break;

            case MouseButton.Right:
                switch (action)
                {
                    case ButtonAction.Down: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_RIGHT_BUTTON_DOWN;
                    case ButtonAction.Up: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_RIGHT_BUTTON_UP;
                    case ButtonAction.DoubleClick: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_RIGHT_BUTTON_DOUBLE_CLICK;
                }
                break;

            case MouseButton.Middle:
                switch (action)
                {
                    case ButtonAction.Down: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_MIDDLE_BUTTON_DOWN;
                    case ButtonAction.Up: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_MIDDLE_BUTTON_UP;
                    case ButtonAction.DoubleClick: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_MIDDLE_BUTTON_DOUBLE_CLICK;
                }
                break;

            case MouseButton.X1:
            case MouseButton.X2:
                switch (action)
                {
                    case ButtonAction.Down: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_X_BUTTON_DOWN;
                    case ButtonAction.Up: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_X_BUTTON_UP;
                    case ButtonAction.DoubleClick: return COREWEBVIEW2_MOUSE_EVENT_KIND.COREWEBVIEW2_MOUSE_EVENT_KIND_X_BUTTON_DOUBLE_CLICK;
                }
                break;
        }
        throw new NotSupportedException();
    }

    public static MouseButton MessageToButton(uint msg, WPARAM wParam)
    {
        switch (msg)
        {
            case MessageDecoder.WM_LBUTTONDOWN:
            case MessageDecoder.WM_LBUTTONUP:
            case MessageDecoder.WM_LBUTTONDBLCLK:
            case MessageDecoder.WM_NCLBUTTONDOWN:
            case MessageDecoder.WM_NCLBUTTONUP:
            case MessageDecoder.WM_NCLBUTTONDBLCLK:
                return MouseButton.Left;

            case MessageDecoder.WM_RBUTTONDOWN:
            case MessageDecoder.WM_RBUTTONUP:
            case MessageDecoder.WM_RBUTTONDBLCLK:
            case MessageDecoder.WM_NCRBUTTONDOWN:
            case MessageDecoder.WM_NCRBUTTONUP:
            case MessageDecoder.WM_NCRBUTTONDBLCLK:
                return MouseButton.Right;

            case MessageDecoder.WM_MBUTTONDOWN:
            case MessageDecoder.WM_MBUTTONUP:
            case MessageDecoder.WM_MBUTTONDBLCLK:
            case MessageDecoder.WM_NCMBUTTONDOWN:
            case MessageDecoder.WM_NCMBUTTONUP:
            case MessageDecoder.WM_NCMBUTTONDBLCLK:
                return MouseButton.Middle;

            case MessageDecoder.WM_XBUTTONDOWN:
            case MessageDecoder.WM_XBUTTONUP:
            case MessageDecoder.WM_XBUTTONDBLCLK:
            case MessageDecoder.WM_NCXBUTTONDOWN:
            case MessageDecoder.WM_NCXBUTTONUP:

                var xb = wParam.Value.HIWORD();
                if (xb == 1)
                    return MouseButton.X1;

                if (xb == 2)
                    return MouseButton.X2;

                break;
        }
        throw new NotSupportedException();
    }

    public static void Open(string? fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            return;

        Process.Start(new ProcessStartInfo
        {
            FileName = fileName,
            UseShellExecute = true,
        });
    }

    public unsafe static void EnableAcrylicBlurBehind(HWND hwnd)
    {
        var accent = new ACCENT_POLICY
        {
            AccentState = ACCENT_STATE.ACCENT_ENABLE_ACRYLICBLURBEHIND
        };

        //accent.GradientColor = 0xFF00FF00;
        var data = new WINDOWCOMPOSITIONATTRIBDATA
        {
            dwAttrib = WINDOWCOMPOSITIONATTRIB.WCA_ACCENT_POLICY,
            cbData = sizeof(ACCENT_POLICY),
            pvData = (nint)(&accent)
        };
        DirectN.Functions.SetWindowCompositionAttribute(hwnd, ref data).ThrowOnError();
    }

    public unsafe static void EnableBlurBehind(HWND hwnd)
    {
        var accent = new ACCENT_POLICY
        {
            AccentState = ACCENT_STATE.ACCENT_ENABLE_BLURBEHIND
        };

        //accent.GradientColor = 0xFF00FF00;
        var data = new WINDOWCOMPOSITIONATTRIBDATA
        {
            dwAttrib = WINDOWCOMPOSITIONATTRIB.WCA_ACCENT_POLICY,
            cbData = sizeof(ACCENT_POLICY),
            pvData = (nint)(&accent)
        };
        DirectN.Functions.SetWindowCompositionAttribute(hwnd, ref data).ThrowOnError();
    }

    public static WicBitmapSource? GdiCapture(this Window window)
    {
        var hdc = DirectN.Functions.GetDC(window.Handle);
        if (hdc == 0)
            return null;

        var hdcMem = DirectN.Functions.CreateCompatibleDC(hdc);
        if (hdcMem == 0)
            return null;

        var rc = window.ClientRect;
        var bmp = DirectN.Functions.CreateCompatibleBitmap(hdc, rc.Width, rc.Height);
        if (DirectN.Functions.SelectObject(hdcMem, new(bmp.Value)) == 0)
        {
            DirectN.Functions.DeleteObject(new(hdcMem));
            DirectN.Functions.DeleteObject(new(bmp));
            _ = DirectN.Functions.ReleaseDC(window.Handle, hdc);
            return null;
        }

        DirectN.Functions.PrintWindow(window.Handle, hdcMem, (PRINT_WINDOW_FLAGS)3);

        var wic = WicBitmapSource.FromHBITMAP(bmp);
        DirectN.Functions.DeleteObject(new(bmp));
        DirectN.Functions.DeleteObject(new(hdcMem));
        _ = DirectN.Functions.ReleaseDC(window.Handle, hdc);
        return wic;
    }
}
