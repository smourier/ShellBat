namespace ShellBat;

public partial class ShellBatUpdateWindow : D2DRenderWindow
{
    private const int _padding = 10;
    private const int _height = 90;
    private static readonly D3DCOLORVALUE _progressColor = D3DCOLORVALUE.Goldenrod;

    private readonly CancellationTokenSource _cts = new();
    private float _progressStep = 0.02f;
    private float _progress;
    private bool _closingQuestion;
    private bool _updateCompleted;

    private ShellBatUpdateWindow(string updateAuthId)
        : base(Res.UpdateTitle)
    {
        Style &= ~(WINDOW_STYLE.WS_MAXIMIZEBOX | WINDOW_STYLE.WS_MINIMIZEBOX);
        BackgroundColor = D3DCOLORVALUE.Transparent;

        _ = Task.Run(async () =>
        {
            try
            {
                await ShellBatInstance.UpdateToLatest(updateAuthId, _cts.Token);
                _updateCompleted = true;
            }
            catch (Exception ex)
            {
                _updateCompleted = true;
                ShellBatInstance.LogError($"Update failed: {ex}");
                if (!_closingQuestion)
                {
                    MessageBox.Show(Handle, $"{Res.UpdateFailed} {ex.GetAllMessagesWithDots()}", Res.UpdateTitle, style: MESSAGEBOX_STYLE.MB_ICONERROR | MESSAGEBOX_STYLE.MB_OK);
                }
            }

            _ = RunTaskOnUIThread(Close);
        });
    }

    protected override Icon? LoadCreationIcon() => Program.LoadIcon(32);

    protected override void OnPositionChanging(object? sender, ValueEventArgs<WINDOWPOS> e)
    {
        var pos = e.Value;
        e.Cancel = pos.cy != _height;
        pos.cy = _height;
        e.Value = pos;
        base.OnPositionChanging(sender, e);
    }

    protected override void Render(IComObject<ID2D1HwndRenderTarget> renderTarget)
    {
        if (_closingQuestion)
            return;

        base.Render(renderTarget);
        var cr = ClientRect;
        var props = new D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES
        {
            startPoint = new D2D_POINT_2F { x = _padding, y = (cr.Height - 2 * _padding) / 2 },
            endPoint = new D2D_POINT_2F { x = cr.Width - _padding, y = (cr.Height - 2 * _padding) / 2 },
        };

        using var stops = renderTarget.CreateGradientStopCollection((D2D1_GRADIENT_STOP[])[
            new D2D1_GRADIENT_STOP { position = 0, color = D3DCOLORVALUE.Transparent },
            new D2D1_GRADIENT_STOP { position = _progress, color = _progressColor },
            new D2D1_GRADIENT_STOP { position = 1, color = D3DCOLORVALUE.Transparent},
        ]);

        using var brush = renderTarget.CreateLinearGradientBrush(props, stops);
        renderTarget.FillRectangle(D2D_RECT_F.Sized(_padding, _padding, cr.Width - _padding, cr.Height - 2 * _padding), brush);
        IncrementProgress();
    }

    private void IncrementProgress()
    {
        _progress += _progressStep;
        if (_progress <= 0 || _progress >= 1)
        {
            _progress -= _progressStep;
            _progressStep = -_progressStep;
        }
        Invalidate();
    }

    protected override void OnClosing(object? sender, CancelEventArgs e)
    {
        if (!_updateCompleted)
        {
            MESSAGEBOX_RESULT result;
            _closingQuestion = true;
            try
            {
                result = MessageBox.Show(Handle, Res.ConfirmCancel, Res.UpdateTitle, style: MESSAGEBOX_STYLE.MB_ICONQUESTION | MESSAGEBOX_STYLE.MB_YESNO);
            }
            finally
            {
                _closingQuestion = false;
            }

            if (result != MESSAGEBOX_RESULT.IDYES)
            {
                _cts?.Cancel();
                e.Cancel = true;
                Invalidate();
                return;
            }
        }
        base.OnClosing(sender, e);
    }

    internal static void ShowProgress(string updateAuthId)
    {
        using var app = new ShellBatApplication();
        WindowSynchronizationContext.Install();

        using var win = new ShellBatUpdateWindow(updateAuthId);

        var width = 300;
        var monitor = WebScreen.Get(CommandLine.Current.GetNullifiedArgument(Program.ScreenDevicePathArgumentName))?.Monitor ??
            win.GetMonitor(MONITOR_FROM_FLAGS.MONITOR_DEFAULTTONEAREST) ??
            DirectN.Extensions.Utilities.Monitor.Primary;
        if (monitor != null)
        {
            var area = monitor.WorkingArea;
            win.ResizeAndMove(RECT.Sized(
                area.left + (area.Width - width) / 2,
                area.top + (area.Height - _height) / 2,
                width,
                _height));
        }
        else
        {
            win.ResizeClient(width, _height);
        }

        win.Center();
        win.Show();
        win.SetForeground();
        app.Run();
    }
}
