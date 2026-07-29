namespace ShellBat;

[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)]
public class ShellBatCommand(string name)
{
    // file commands
    public static ShellBatCommand Copy { get; } = new(nameof(Copy)) { Key = new(VIRTUAL_KEY.VK_C) { RequiresCtrl = true }, IsReadOnly = true, CategoryName = Res.Builtin };
    public static ShellBatCommand Cut { get; } = new(nameof(Cut)) { Key = new(VIRTUAL_KEY.VK_X) { RequiresCtrl = true }, IsReadOnly = true, CategoryName = Res.Builtin };
    public static ShellBatCommand Paste { get; } = new(nameof(Paste)) { Key = new(VIRTUAL_KEY.VK_V) { RequiresCtrl = true }, IsReadOnly = true, CategoryName = Res.Builtin };

    // general commands
    public static ShellBatCommand OpenSettings { get; } = new(nameof(OpenSettings)) { Key = new(VIRTUAL_KEY.VK_G) { RequiresCtrl = true } };
    public static ShellBatCommand OpenInstanceSettings { get; } = new(nameof(OpenInstanceSettings)) { Key = new(VIRTUAL_KEY.VK_H) { RequiresCtrl = true } };
    public static ShellBatCommand OpenWithExplorer { get; } = new(nameof(OpenWithExplorer)) { Key = new(VIRTUAL_KEY.VK_E) { RequiresAlt = true } };
    public static ShellBatCommand OpenDesktop { get; } = new(nameof(OpenDesktop)) { Key = new(VIRTUAL_KEY.VK_D) { RequiresAlt = true } };
    public static ShellBatCommand OpenThisPC { get; } = new(nameof(OpenThisPC)) { Key = new(VIRTUAL_KEY.VK_P) { RequiresAlt = true } };
    public static ShellBatCommand OpenRecycleBin { get; } = new(nameof(OpenRecycleBin)) { Key = new(VIRTUAL_KEY.VK_R) { RequiresAlt = true } };
    public static ShellBatCommand FindStrings { get; } = new(nameof(FindStrings)) { Key = new(VIRTUAL_KEY.VK_F) { RequiresAlt = true } };
    public static ShellBatCommand OpenSearch { get; } = new(nameof(OpenSearch)) { Key = new(VIRTUAL_KEY.VK_S) { RequiresAlt = true } };
    public static ShellBatCommand CopyFullPath { get; } = new(nameof(CopyFullPath)) { Key = new(VIRTUAL_KEY.VK_F) { RequiresCtrl = true } };
    public static ShellBatCommand OpenTerminal { get; } = new(nameof(OpenTerminal)) { Key = new(VIRTUAL_KEY.VK_T) { RequiresAlt = true } };
    public static ShellBatCommand CloseWindow { get; } = new(nameof(CloseWindow)) { Key = new(VIRTUAL_KEY.VK_W) { RequiresCtrl = true } };
    public static ShellBatCommand MinimizeWindow { get; } = new(nameof(MinimizeWindow)) { Key = new(VIRTUAL_KEY.VK_M) { RequiresAlt = true } };
    public static ShellBatCommand MaximizeWindow { get; } = new(nameof(MaximizeWindow)) { Key = new(VIRTUAL_KEY.VK_X) { RequiresAlt = true } };
    public static ShellBatCommand RestartAsAdministrator { get; } = new(nameof(RestartAsAdministrator)) { Key = new(VIRTUAL_KEY.VK_R) { RequiresCtrl = true, RequiresShift = true } };
    public static ShellBatCommand OpenViewers { get; } = new(nameof(OpenViewers)) { Key = new(VIRTUAL_KEY.VK_V) { RequiresAlt = true } };

    // navigation commands
    public static ShellBatCommand GoBack { get; } = new(nameof(GoBack)) { Key = new(VIRTUAL_KEY.VK_LEFT) { RequiresAlt = true }, CategoryName = Res.Navigation };
    public static ShellBatCommand GoForward { get; } = new(nameof(GoForward)) { Key = new(VIRTUAL_KEY.VK_RIGHT) { RequiresAlt = true }, CategoryName = Res.Navigation };
    public static ShellBatCommand GoUp { get; } = new(nameof(GoUp)) { Key = new(VIRTUAL_KEY.VK_UP) { RequiresAlt = true }, CategoryName = Res.Navigation };
    public static ShellBatCommand OpenHistory { get; } = new(nameof(OpenHistory)) { Key = new(VIRTUAL_KEY.VK_H) { RequiresAlt = true }, CategoryName = Res.Navigation };
    public static ShellBatCommand OpenFavorites { get; } = new(nameof(OpenFavorites)) { Key = new(VIRTUAL_KEY.VK_A) { RequiresAlt = true }, CategoryName = Res.Navigation };

    // instance commands
    public static ShellBatCommand OpenNewInstance { get; } = new(nameof(OpenNewInstance)) { Key = new(VIRTUAL_KEY.VK_I) { RequiresAlt = true }, CategoryName = Res.Instances };
    public static ShellBatCommand OpenNewInstanceOnOtherScreen { get; } = new(nameof(OpenNewInstanceOnOtherScreen)) { Key = new(VIRTUAL_KEY.VK_I) { RequiresAlt = true, RequiresShift = true }, CategoryName = Res.Instances };

    // readonly commands
    public static ShellBatCommand OpenKeyboardShortcuts { get; } = new(nameof(OpenKeyboardShortcuts)) { Key = new(VIRTUAL_KEY.VK_K) { RequiresCtrl = true }, IsReadOnly = true, CategoryName = Res.Builtin };
    public static ShellBatCommand CloseApplication { get; } = new(nameof(CloseApplication)) { Key = new(VIRTUAL_KEY.VK_F4) { RequiresAlt = true }, IsReadOnly = true, CategoryName = Res.Builtin };
    public static ShellBatCommand Reload { get; } = new(nameof(Reload)) { Key = new(VIRTUAL_KEY.VK_F5), IsReadOnly = true, CategoryName = Res.Builtin };
    public static ShellBatCommand EditAddress { get; } = new(nameof(EditAddress)) { Key = new(VIRTUAL_KEY.VK_F6), IsReadOnly = true, CategoryName = Res.Builtin };
    public static ShellBatCommand Rename { get; } = new(nameof(Rename)) { Key = new(VIRTUAL_KEY.VK_F2), IsReadOnly = true, CategoryName = Res.Builtin };
    public static ShellBatCommand ToggleFullScreen { get; } = new(nameof(ToggleFullScreen)) { Key = new(VIRTUAL_KEY.VK_F11), IsReadOnly = true, CategoryName = Res.Builtin };
    public static ShellBatCommand OpenDevTools { get; } = new(nameof(OpenDevTools)) { Key = new(VIRTUAL_KEY.VK_F12), IsReadOnly = true, IsHidden = true };

    public string Name { get; } = name ?? throw new ArgumentNullException(nameof(name));
    public string DisplayName => Res.ResourceManager.GetString(Name).Nullify() ?? Conversions.Decamelize(Name);

    public virtual bool IsReadOnly { get; set; }
    public virtual bool IsHidden { get; set; }
    public virtual CommandKey Key { get; private set; } = null!;
    public virtual string? CategoryName { get; set; }

    public override string ToString() => $"{Name} '{DisplayName}' [{Key}]";

    public static IReadOnlyList<ShellBatCommand> All => Commands._allCommands.Value.AllCommands;
    public static IReadOnlyDictionary<string, ShellBatCommand> AllByName => Commands._allCommands.Value.AllByName;
    public static IReadOnlyDictionary<CommandKey, ShellBatCommand> AllByKey => Commands._allCommands.Value.AllByKey;
    public static IReadOnlyDictionary<string, string> DefaultCommandTexts => Commands._allCommands.Value.DefaultCommandTexts;

    internal static void ProcessSettings(HWND handle) => Commands._allCommands.Value.ProcessSettings(handle);
    internal static void SetCommandKey(WebCommandKey wc)
    {
        var cmd = All.FirstOrDefault(c => c.Name.EqualsIgnoreCase(wc.CommandName));
        if (cmd == null)
            return;

        var key = CommandKey.Parse(wc.CommandText);
        if (key == null || cmd.Key.Equals(key))
            return;

        cmd.Key = key;
        var settings = Settings.Current.Shortcuts ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (settings.TryGetValue(cmd.Name, out var existing) && existing.EqualsIgnoreCase(key.ToString()))
            return;

        if (cmd.IsReadOnly)
            return;

        settings[cmd.Name] = key.ToString();

        // removed default shortcuts from settings
        foreach (var kv in Commands._allCommands.Value.DefaultCommandTexts)
        {
            if (settings.TryGetValue(kv.Key, out var val) && val.EqualsIgnoreCase(kv.Value))
            {
                settings.Remove(kv.Key);
            }
        }
        Settings.Current.SerializeToConfiguration();
    }

    internal static string? GetDefaultCommandText(string commandName)
    {
        DefaultCommandTexts.TryGetValue(commandName, out var text);
        return text;
    }

    internal static ShellBatCommand? GetCommandByKey(KeyEventArgs args)
    {
        var key = GetKey(args);
        if (key == null)
            return null;

        AllByKey.TryGetValue(key, out var command);
        return command;
    }

    private static string? GetInvalidKeyErrorText(CommandKey key)
    {
        if (!key.HasNonModifierVirtualKey || !key.PreventsDefault)
            return null;

        return string.Format(Res.InvalidShortcutKey, key.ToString());
    }

    private static CommandKey? GetKey(KeyEventArgs e)
    {
        CommandKey key;
        if (e.Key == 0)
        {
            // dead keys, weird stuff, etc.
            if (e.Character == null || e.Character.Length != 1)
                return null;

            key = CommandKey.FromCharacter(e.Character![0]);
        }
        else
        {
            // the browser keeps US layout for letters, so we must convert from char
            if (e.Key.IsLetter() && e.Character != null && e.Character.Length == 1)
            {
                key = CommandKey.FromCharacter(e.Character![0]);
            }
            else
            {
                key = new CommandKey(e.Key);
            }
        }

        key.RequiresCtrl = e.WithControl;
        key.RequiresShift = e.WithShift;
        key.RequiresAlt = e.WithMenu;
        return key;
    }

    internal static WebCommandKey Validate(HWND handle, KeyEventArgs e, string commandName)
    {
        var key = GetKey(e);
        if (key == null)
        {
            key = new CommandKey(e.Key)
            {
                RequiresCtrl = e.WithControl,
                RequiresShift = e.WithShift,
                RequiresAlt = e.WithMenu
            };

            return new WebCommandKey()
            {
                Key = key.Key,
                DefaultCommandText = GetDefaultCommandText(commandName),
                CommandText = key.ToString(),
                Ctrl = key.RequiresCtrl,
                Shift = key.RequiresShift,
                Alt = key.RequiresAlt,
                IsValid = false,
                ErrorText = GetInvalidKeyErrorText(key),
                PreventDefault = key.PreventsDefault
            };
        }

        var wc = new WebCommandKey
        {
            Key = key.Key,
            DefaultCommandText = GetDefaultCommandText(commandName),
            CommandText = key.ToString(),
            Ctrl = key.RequiresCtrl,
            Shift = key.RequiresShift,
            Alt = key.RequiresAlt
        };

        // check it's valid or already assigned
        if (!key.IsValid(handle) || (AllByKey.TryGetValue(key, out var command) && command.Name != commandName))
        {
            wc.ErrorText = GetInvalidKeyErrorText(key);
            wc.PreventDefault = key.PreventsDefault;
        }
        else
        {
            wc.PreventDefault = true;
            wc.IsValid = true;
        }
        return wc;
    }

    private sealed class Commands
    {
        public static readonly Lazy<Commands> _allCommands = new(() => new());

        public Commands()
        {
            AllCommands = (ShellBatCommand[])[.. typeof(ShellBatCommand).GetProperties(BindingFlags.Public | BindingFlags.Static)
                .Where(p => p.PropertyType == typeof(ShellBatCommand))
                .Select(p => (ShellBatCommand)p.GetValue(null)!)
                .Where(p => p.Key != null)];

            AllByName = AllCommands.ToDictionary(c => c.Name, StringComparer.OrdinalIgnoreCase);
            AllByKey = AllCommands.ToDictionary(c => c.Key, c => c);

            // remember default shortcuts
            var defShortcuts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var cmd in AllCommands)
            {
                defShortcuts[cmd.Name] = cmd.Key.ToString();
            }
            DefaultCommandTexts = defShortcuts.AsReadOnly();
        }

        public IReadOnlyList<ShellBatCommand> AllCommands { get; }
        public IReadOnlyDictionary<string, ShellBatCommand> AllByName { get; }
        public IReadOnlyDictionary<CommandKey, ShellBatCommand> AllByKey { get; }
        public IReadOnlyDictionary<string, string> DefaultCommandTexts { get; }

        public void ProcessSettings(HWND handle)
        {
            var dic = Settings.Current.Shortcuts;
            if (dic == null)
                return;

            foreach (var kv in dic)
            {
                if (!AllByName.TryGetValue(kv.Key, out var command))
                    continue;

                var key = CommandKey.Parse(kv.Value);
                if (key == null)
                {
                    Application.TraceWarning($"Shortcut key '{kv.Value}' for command '{command.Name}' cannot be parsed. Skipping.");
                    continue;
                }

                if (!key.IsValid(handle))
                {
                    Application.TraceWarning($"Shortcut key '{key}' for command '{command.Name}' is not valid. Skipping.");
                    continue;
                }

                if (AllByKey.TryGetValue(key, out var cmd) && !cmd.Name.EqualsIgnoreCase(command.Name))
                {
                    Application.TraceWarning($"Shortcut key '{key}' for command '{command.Name}' is already assigned to command '{cmd.Name}'. Skipping.");
                    continue;
                }

                command.Key = key;
            }
        }
    }
}
