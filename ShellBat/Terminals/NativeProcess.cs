namespace ShellBat.Terminals;

public sealed partial class NativeProcess(STARTUPINFOEX startupInfo, PROCESS_INFORMATION processInfo) : IDisposable
{
    private bool _disposedValue;

    public STARTUPINFOEX StartupInfo { get; } = startupInfo;
    public PROCESS_INFORMATION ProcessInfo { get; } = processInfo;
    public int ProcessId => (int)ProcessInfo.dwProcessId;
    public bool IsDisposed => _disposedValue;

    public AutoResetEvent GetWaitForExitEvent() => new(false) { SafeWaitHandle = new SafeWaitHandle(ProcessInfo.hProcess, ownsHandle: false) };

    ~NativeProcess() { Dispose(); }
    public void Dispose()
    {
        var disposed = Interlocked.Exchange(ref _disposedValue, true);
        if (!disposed)
        {
            if (StartupInfo.lpAttributeList != 0)
            {
                Interop.Functions.DeleteProcThreadAttributeList(StartupInfo.lpAttributeList);
                Marshal.FreeHGlobal(StartupInfo.lpAttributeList);
            }

            if (ProcessInfo.hProcess != 0)
            {
                DirectN.Functions.CloseHandle(ProcessInfo.hProcess);
            }

            if (ProcessInfo.hThread != 0)
            {
                DirectN.Functions.CloseHandle(ProcessInfo.hThread);
            }
        }
        GC.SuppressFinalize(this);
    }

    public static NativeProcess? Start(string commandLine, string? currentDirectory, nint attributes, nint hPC, bool throwOnError = true) =>
        Start(commandLine, currentDirectory, attributes, hPC, out _, throwOnError);

    public static unsafe NativeProcess? Start(string commandLine, string? currentDirectory, nint attributes, nint hPC, out Exception? error, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(commandLine);
        error = null;
        var startupInfo = ConfigureProcessThread(hPC, attributes);

        DirectN.Functions.SetErrorMode(THREAD_ERROR_MODE.SEM_FAILCRITICALERRORS);
        if (!Interop.Functions.CreateProcessW(
            PWSTR.Null,
            PWSTR.From(commandLine),
            new SECURITY_ATTRIBUTES { nLength = (uint)sizeof(SECURITY_ATTRIBUTES) },
            new SECURITY_ATTRIBUTES { nLength = (uint)sizeof(SECURITY_ATTRIBUTES) },
            false,
            Interop.Constants.EXTENDED_STARTUPINFO_PRESENT,
            0,
            PWSTR.From(currentDirectory),
            startupInfo,
            out var info))
        {
            DirectN.Functions.SetErrorMode(0);
            if (throwOnError)
                throw new Win32Exception(Marshal.GetLastPInvokeError());

            error = new Win32Exception(Marshal.GetLastPInvokeError());
            return null;
        }

        DirectN.Functions.SetErrorMode(0);
        return new NativeProcess(startupInfo, info);
    }

    // cf https://docs.microsoft.com/en-us/windows/console/creating-a-pseudoconsole-session#preparing-for-creation-of-the-child-process
    private unsafe static STARTUPINFOEX ConfigureProcessThread(HPCON hPC, nint attributes)
    {
        nint size = 0;
        Interop.Functions.InitializeProcThreadAttributeList(0, 1, 0, ref size);

        var startupInfo = new STARTUPINFOEX();
        startupInfo.StartupInfo.cb = (uint)sizeof(STARTUPINFOEX);
        startupInfo.lpAttributeList = Marshal.AllocHGlobal(size);

        if (!Interop.Functions.InitializeProcThreadAttributeList(startupInfo.lpAttributeList, 1, 0, ref size))
            throw new Win32Exception(Marshal.GetLastPInvokeError());

        if (!Interop.Functions.UpdateProcThreadAttribute(startupInfo.lpAttributeList, 0, attributes, hPC, nint.Size, 0, 0))
            throw new Win32Exception(Marshal.GetLastPInvokeError());

        return startupInfo;
    }
}
