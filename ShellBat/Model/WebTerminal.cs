namespace ShellBat.Model;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class WebTerminal : DispatchObject, IDisposable
{
#pragma warning disable CA1822 // Mark members as static

    private const int _oscCommandNotFound = 9001;
    private const string _oscCommandNotFoundToken = "CmdNotFound;";
    private static readonly string _windowsDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);

    private readonly ShellBatWindow? _window;
    private readonly TerminalEntry _terminalEntry;
    private Terminal? _terminal;
    private string? _currentWorkingDirectory;
    private bool _shellBatSync;
    private bool _terminalSync;
    private readonly BlindTerminal _blindTerminal = new();

    public event EventHandler? WorkingDirectoryChanged;

    public WebTerminal(TerminalEntry terminalEntry, Entry entry, ShellBatWindow? window, string? id)
    {
        ArgumentNullException.ThrowIfNull(terminalEntry);
        ArgumentNullException.ThrowIfNull(entry);
        _terminalEntry = terminalEntry;
        Entry = entry;
        FolderEntry = GetFolderEntry();
        _window = window;
        Id = id;
        if (_window != null)
        {
            _terminal = new Terminal();
        }

        _shellBatSync = SupportsShellBatSync;
        _terminalSync = SupportsTerminalSync;
    }

    public Entry Entry { get; }
    public Entry FolderEntry { get; }
    public ShellBatWindow? Window => _window;
    public bool SupportsShellBatSync => (nint.Size == 4 && !ProcessUtilities.IsWow64(null)) || _terminalEntry.SupportsShellBatSync;
    public bool SupportsTerminalSync => (nint.Size == 4 && !ProcessUtilities.IsWow64(null)) || _terminalEntry.SupportsTerminalSync;
    public bool SupportsCmdNotFound => _terminalEntry.SupportsCmdNotFound;
    public bool IsWsl => _terminalEntry.IsWsl;
    public string? Id { get; }
    public string? ProcessName { get; private set; }
    public string Key => _terminalEntry.Key;
    public string? DisplayName => _terminalEntry.DisplayName.Nullify() ?? Res.ResourceManager.GetString(Key).Nullify() ?? Key;
    public string Icon => _terminalEntry.FinalIcon;

    public string? CommandLine
    {
        get
        {
            var cmd = _terminalEntry.CommandLine;
            if (cmd == null)
                return null;

            var arg = IsWsl ? FolderEntry.WslFileSystemPath : FolderEntry.ParsingName;
            cmd = Environment.ExpandEnvironmentVariables(string.Format(cmd, arg));
            return cmd;
        }
    }

    public bool GetShellBatSync() => _shellBatSync;
    public void SetShellBatSync(bool value)
    {
        if (_shellBatSync == value)
            return;

        _shellBatSync = value;
        if (_shellBatSync && _currentWorkingDirectory != null)
        {
            _window?.RunTaskOnUIThread(() => _window?.Navigate(_currentWorkingDirectory, false));
        }
    }

    public bool GetTerminalSync() => _terminalSync;
    public void SetTerminalSync(bool value)
    {
        if (_terminalSync == value)
            return;

        _terminalSync = value;
    }

    public string? CurrentWorkingDirectory
    {
        get => _currentWorkingDirectory;
        private set
        {
            value = value.Nullify();
            if (value == null)
                return;

            if (value.EndsWith(Path.DirectorySeparatorChar))
            {
                value = value[..^1];
            }
            if (_currentWorkingDirectory.EqualsIgnoreCase(value))
                return;

            _currentWorkingDirectory = value;
            OnWorkingDirectoryChanged(this, EventArgs.Empty);

            // many console app navigate to Windows, we don't follow
            if (_shellBatSync && _currentWorkingDirectory != null && !_currentWorkingDirectory.EqualsIgnoreCase(_windowsDir))
            {
                _window?.RunTaskOnUIThread(() => _window?.Navigate(_currentWorkingDirectory, false));
            }
        }
    }

    override public string ToString() => $"{Key}/{Id} - {CommandLine}";

    protected virtual void OnWorkingDirectoryChanged(object sender, EventArgs e) => WorkingDirectoryChanged?.Invoke(sender, e);
    protected virtual Entry GetFolderEntry()
    {
        if (IsValidFolder(Entry))
            return Entry;

        var parent = Entry.Parent;
        if (IsValidFolder(parent))
            return parent!;

        var userProfile = Entry.UserProfile;
        if (IsValidFolder(userProfile))
            return userProfile;

        return Entry.Get(null, Environment.CurrentDirectory)!;
    }

    private static bool IsValidFolder(Entry? entry) => entry is not null && entry.IsFolder && entry.FileSystemPath is not null && IOUtilities.PathIsDirectory(entry.FileSystemPath);

    public virtual void OnReady(string? commandLine)
    {
        var terminal = _terminal;
        if (terminal is null)
            return;

        var cmdLine = commandLine.Nullify() ?? CommandLine;
        if (cmdLine is null)
        {
            Application.TraceInfo("No command line defined");
            Dispose();
            return;
        }

        terminal.ProcessExited += (s, e) =>
        {
            if (Settings.Current.CloseExitedTerminalProcesses)
            {
                //Application.TraceInfo("Process exited");
                Dispose();
            }
            else
            {
                //Application.TraceInfo("Process exited, notifying webview");
                var text = string.Format(Res.ProcessExited, ProcessName ?? terminal.Process?.ProcessId.ToString());
                WriteProcessExitedToWebView(text);
                //Application.TraceInfo("Notified webview of process exit");
            }
        };

        terminal.OutputReady += (s, e) => new Thread(ReadCharactersWithState) { Name = $"xTerm:{typeof(Program).Namespace}:{Environment.TickCount64}", IsBackground = true }.Start();

        CurrentWorkingDirectory = FolderEntry.FileSystemPath;
        try
        {
            var np = terminal.Start(cmdLine, CurrentWorkingDirectory, throwOnError: false);
            if (np != null)
            {
                try
                {
                    var process = Process.GetProcessById(np.ProcessId);
                    ProcessName = process.ProcessName;
                }
                catch
                {
                    // too bad, continue
                }
            }
            else if (terminal.ProcessStartError != null)
            {
                var text = $"Failed to start terminal process using '{cmdLine}' command line: {terminal.ProcessStartError.GetAllMessagesWithDots()}";
                Application.TraceError(text);
                WriteLineToWebView(text);

                text = string.Format(Res.ProcessExited, cmdLine);
                WriteProcessExitedToWebView(text);
            }
        }
        catch (Exception ex)
        {
            var text = $"Failed to start terminal process using '{cmdLine}' command line: {ex.GetAllMessagesWithDots()}";
            Application.TraceError(text);
            WriteLineToWebView(text);
            return;
        }
    }

    public virtual void OnResize(int cols, int rows)
    {
        //Application.TraceInfo($"Resize: {cols} x {rows}");
        _terminal?.Console.Resize((short)cols, (short)rows);
    }

    public virtual void OnData(string data)
    {
        //Application.TraceVerbose("data: " + data);
        WriteToTerminal(data);
    }

    // return false to eat the event
    public virtual bool OnKeyEvent(object? obj)
    {
        if (obj is not DirectN.IDispatch disp)
            return true;

        var keyEvent = disp.AsWebKeyEvent();
        if (keyEvent == null)
            return true;

        var result = _blindTerminal.Add(keyEvent);
        if (result.EnterPressed)
        {
            if (result.Command != null)
            {
                //Application.TraceVerbose($"running term:'{_blindTerminal}' cmd:{result.Command}");
                if (result.Command.Run(this))
                {
                    InputToWebView("\x1B", false);
                    return false;
                }
            }

            if (keyEvent.Type != "keydown")
                return false;
        }
        return true;
    }

    public void WriteProcessExitedToWebView(string text)
    {
        if (Id == null || _window == null)
            return;

        var script = $"terminalProcessExited(\"{Id}\", \"{HttpUtility.JavaScriptStringEncode(text)}\");";
        _window.RunTaskOnUIThread(() => _window.ExecuteScript(script));
    }

    public void WriteBytesToWebView(string data)
    {
        ArgumentNullException.ThrowIfNull(data);
        WriteBytesToWebView(Encoding.UTF8.GetBytes(data));
    }

    public void WriteBytesToWebView(Span<byte> data)
    {
        if (Id == null || _window == null)
            return;

        var script = $"writeBytesTerminal(\"{Id}\", \"{Convert.ToBase64String(data)}\");";
        _window.RunTaskOnUIThread(() => _window.ExecuteScript(script));
    }

    public void WriteToWebView(string text)
    {
        if (Id == null || _window == null)
            return;

        var script = $"writeTerminal(\"{Id}\", \"{HttpUtility.JavaScriptStringEncode(text)}\");";
        _window.RunTaskOnUIThread(() => _window.ExecuteScript(script));
    }

    public void WriteLineToWebView(string text)
    {
        if (Id == null || _window == null)
            return;

        var script = $"writeLnTerminal(\"{Id}\", \"{HttpUtility.JavaScriptStringEncode(text)}\");";
        _window.RunTaskOnUIThread(() => _window.ExecuteScript(script));
    }

    public void ChangeDirectoryInWebView(Entry entry)
    {
        ArgumentNullException.ThrowIfNull(entry);
        if (Id == null || _window == null)
            return;

        var path = IsWsl ? entry.WslFileSystemPath : entry.FileSystemPath;
        ChangeDirectoryInWebView(path);
    }

    public void ChangeDirectoryInWebView(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;

        var command = _terminalEntry.ChangeDirectoryCommand.Nullify();
        if (command == null)
            return;

        var cd = string.Format(command, path);

        if (IsWsl)
        {
            PasteToWebView(cd);
            InputToWebView("\n", false);
        }
        else
        {
            if (!cd.EndsWith("\r\n"))
            {
                cd += "\r\n";
            }
            PasteToWebView(cd);
        }
    }

    public void PasteToWebView(string text)
    {
        if (Id == null || _window == null)
            return;

        var script = $"pasteTerminal(\"{Id}\", \"{HttpUtility.JavaScriptStringEncode(text)}\");";
        _window.RunTaskOnUIThread(() => _window.ExecuteScript(script));
    }

    public void InputToWebView(string text, bool wasUserInput = true)
    {
        if (Id == null || _window == null)
            return;

        var script = $"inputTerminal(\"{Id}\", \"{HttpUtility.JavaScriptStringEncode(text)}\", {wasUserInput.ToString().ToLower()});";
        _window.RunTaskOnUIThread(() => _window.ExecuteScript(script));
    }

    public void ClearWebView()
    {
        if (Id == null || _window == null)
            return;

        var script = $"clearTerminal(\"{Id}\");";
        _window.RunTaskOnUIThread(() => _window.ExecuteScript(script));
    }

    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    protected virtual void Dispose(bool disposing)
    {
        if (Id != null && _window != null)
        {
            _window.DisposeTerminal(Id);
        }
        Interlocked.Exchange(ref _terminal, null)?.Dispose();
    }

    private void ReadCharactersWithState(object? state)
    {
        try
        {
            ReadCharacters();
        }
        catch (Exception ex)
        {
            Application.TraceError($"Error: {ex.GetAllMessagesWithDots()}");
        }
    }

    private void ReadCharacters()
    {
        int read;
        var buffer = new byte[16384];
        var bufferStart = 0;

        // this call waits indefinitely until data is available
        // disposing the terminal will replace the stream by a null stream which
        // will return 0 bytes read and allow the loop to exit
        while ((read = _terminal?.ReadStream!.Read(buffer.AsSpan()[bufferStart..]) ?? 0) != 0)
        {
            CurrentWorkingDirectory = _terminal?.WorkingDirectory;

            // we want complete buffers
            if (read == buffer.Length - bufferStart)
            {
                //Application.TraceVerbose($"Need more data to complete sequences, from {buffer.Length} => {buffer.Length * 2}");
                bufferStart = buffer.Length;
                var newBuf = new byte[buffer.Length * 2];
                Buffer.BlockCopy(buffer, 0, newBuf, 0, buffer.Length);
                buffer = newBuf;
                continue;
            }

            var slice = buffer[..(bufferStart + read)];
            bufferStart = 0;
            //Application.TraceVerbose(read + Environment.NewLine + slice.ToArray().ToHexaDump(read));

            var writeToWebView = true;

            // only cmd seems to support cmdNotFound thing
            if (SupportsCmdNotFound)
            {
                var reader = new TerminalReader(slice);
                do
                {
                    var sequence = reader.Read();
                    if (sequence == null)
                        break;

                    if (sequence is OscSequence seq &&
                        seq.Number == _oscCommandNotFound &&
                        seq.Value.StartsWith(_oscCommandNotFoundToken, StringComparison.OrdinalIgnoreCase))
                    {
                        var cmd = seq.Value[_oscCommandNotFoundToken.Length..];
                        if (HandleCmdNotFound(buffer, cmd))
                        {
                            writeToWebView = false;
                            break;
                        }
                    }

                    //var hex = slice[sequence.Range].AsSpan().ToHexaDump(new HexaDumpOptions { AddHeader = false });
                    //Application.TraceVerbose($"TerminalReader:{sequence}{Environment.NewLine}{hex}");
                }
                while (true);
            }

            if (writeToWebView)
            {
                WriteBytesToWebView(slice);
            }
        }
    }

    private bool HandleCmdNotFound(byte[] buffer, string cmd)
    {
        //Application.TraceVerbose($"cmdNotFound: {cmd}");
        var workingDir = _terminal?.WorkingDirectory;
        if (workingDir == null)
            return false;

        var dir = Path.Combine(workingDir, cmd);
        if (!IOUtilities.PathIsDirectory(dir))
            return false;

        // consume the rest of the cmd not found sequence
        // unfortunately, this will draw empty spaces/lines in the terminal
        // don't know how to do differently right now
        _terminal?.ReadStream?.Read(buffer);
        ChangeDirectoryInWebView(dir);
        return true;
    }

    private void WriteToTerminal(string data)
    {
        var stream = _terminal?.WriteStream;
        if (stream == null)
            return;

        var bytes = Encoding.UTF8.GetBytes(data);
        //Application.TraceVerbose($"write:{Environment.NewLine}{bytes.ToHexaDump()}");
        stream.Write(bytes, 0, bytes.Length);
        stream.Flush();
    }

#pragma warning restore CA1822 // Mark members as static
}
