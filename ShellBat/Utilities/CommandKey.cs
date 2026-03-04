namespace ShellBat.Utilities;

public class CommandKey(VirtualKey key) : IEquatable<CommandKey>
{
    // virtual keys that are generally invalid as command keys
    private static readonly ConcurrentDictionary<VIRTUAL_KEY, object?> _invalidKeys = GetInvalidKeys();
    private static ConcurrentDictionary<VIRTUAL_KEY, object?> GetInvalidKeys()
    {
        var dict = new ConcurrentDictionary<VIRTUAL_KEY, object?>();
        dict[VIRTUAL_KEY.VK_ESCAPE] = null;
        dict[VIRTUAL_KEY.VK_PRINT] = null;
        dict[VIRTUAL_KEY.VK_NUMLOCK] = null;
        dict[VIRTUAL_KEY.VK_RETURN] = null;
        dict[VIRTUAL_KEY.VK_BACK] = null;
        dict[VIRTUAL_KEY.VK_CAPITAL] = null;
        dict[VIRTUAL_KEY.VK_CLEAR] = null;

        dict[VIRTUAL_KEY.VK_MENU] = null;
        dict[VIRTUAL_KEY.VK_LMENU] = null;
        dict[VIRTUAL_KEY.VK_RMENU] = null;

        dict[VIRTUAL_KEY.VK_CONTROL] = null;
        dict[VIRTUAL_KEY.VK_LCONTROL] = null;
        dict[VIRTUAL_KEY.VK_RCONTROL] = null;

        dict[VIRTUAL_KEY.VK_SHIFT] = null;
        dict[VIRTUAL_KEY.VK_LSHIFT] = null;
        dict[VIRTUAL_KEY.VK_RSHIFT] = null;
        return dict;
    }

    // keys that are invalid as shortcuts because they conflict with common shortcuts
    private static readonly ConcurrentDictionary<CommandKey, object?> _invalidShortcutKeys = GetInvalidShortcutKeys();
    private static ConcurrentDictionary<CommandKey, object?> GetInvalidShortcutKeys()
    {
        var dict = new ConcurrentDictionary<CommandKey, object?>();
        dict[new(VIRTUAL_KEY.VK_C) { RequiresCtrl = true }] = null;
        dict[new(VIRTUAL_KEY.VK_V) { RequiresCtrl = true }] = null;
        dict[new(VIRTUAL_KEY.VK_X) { RequiresCtrl = true }] = null;
        dict[new(VIRTUAL_KEY.VK_Z) { RequiresCtrl = true }] = null;

        dict[new(VIRTUAL_KEY.VK_LEFT) { RequiresCtrl = true }] = null;
        dict[new(VIRTUAL_KEY.VK_UP) { RequiresCtrl = true }] = null;
        dict[new(VIRTUAL_KEY.VK_RIGHT) { RequiresCtrl = true }] = null;
        dict[new(VIRTUAL_KEY.VK_DOWN) { RequiresCtrl = true }] = null;

        dict[new(VIRTUAL_KEY.VK_TAB)] = null;
        dict[new(VIRTUAL_KEY.VK_TAB) { RequiresShift = true }] = null;
        dict[new(VIRTUAL_KEY.VK_TAB) { RequiresAlt = true }] = null;
        dict[new(VIRTUAL_KEY.VK_TAB) { RequiresShift = true, RequiresAlt = true }] = null;

        dict[new(VIRTUAL_KEY.VK_SPACE)] = null;
        return dict;
    }

    // keys that allow navigation and thus should not prevent default behavior
    private static readonly ConcurrentDictionary<CommandKey, object?> _dontPreventDefaultKeys = GetDontPreventDefaultKeys();
    private static ConcurrentDictionary<CommandKey, object?> GetDontPreventDefaultKeys()
    {
        var dict = new ConcurrentDictionary<CommandKey, object?>();
        dict[new(VIRTUAL_KEY.VK_TAB)] = null;
        dict[new(VIRTUAL_KEY.VK_TAB) { RequiresShift = true }] = null;
        dict[new(VIRTUAL_KEY.VK_ESCAPE)] = null;
        dict[new(VIRTUAL_KEY.VK_RETURN)] = null;
        return dict;
    }

    // keys that are valid without modifiers
    private static readonly ConcurrentDictionary<VIRTUAL_KEY, object?> _validWithoutModifiersKeys = GetValidWithoutModifiersKeys();
    private static ConcurrentDictionary<VIRTUAL_KEY, object?> GetValidWithoutModifiersKeys()
    {
        var dict = new ConcurrentDictionary<VIRTUAL_KEY, object?>();
        dict[VIRTUAL_KEY.VK_APPS] = null;
        for (var vk = VIRTUAL_KEY.VK_F1; vk <= VIRTUAL_KEY.VK_F24; vk++)
        {
            dict[vk] = null;
        }
        return dict;
    }

    public static VirtualKey Ctrl { get; } = VIRTUAL_KEY.VK_CONTROL;
    public static VirtualKey Shift { get; } = VIRTUAL_KEY.VK_SHIFT;
    public static VirtualKey Alt { get; } = VIRTUAL_KEY.VK_MENU;

    public VirtualKey Key { get; } = key;
    public string? KeyDisplayName => Key.GetName(HotKeyModifiers);

    public virtual bool RequiresCtrl { get; set; }
    public virtual bool RequiresShift { get; set; }
    public virtual bool RequiresAlt { get; set; }
    public bool HasNonModifierVirtualKey => Key.Key != 0 && !Key.Key.IsControlKey() && !Key.Key.IsShiftKey() && !Key.Key.IsAltKey();
    public bool PreventsDefault => !_dontPreventDefaultKeys.ContainsKey(this);
    public HOT_KEY_MODIFIERS HotKeyModifiers
    {
        get
        {
            HOT_KEY_MODIFIERS mods = 0;
            if (RequiresCtrl)
            {
                mods |= HOT_KEY_MODIFIERS.MOD_CONTROL;
            }

            if (RequiresShift)
            {
                mods |= HOT_KEY_MODIFIERS.MOD_SHIFT;
            }

            if (RequiresAlt)
            {
                mods |= HOT_KEY_MODIFIERS.MOD_ALT;
            }
            return mods;
        }
    }

    public bool IsHotkeyAvailable(HWND hwnd) => Key.Key.IsHotkeyAvailable(HotKeyModifiers, hwnd);
    public bool IsValid(HWND handle)
    {
        //Application.TraceVerbose($"Validating command key: {this} k:{Key}");
        if (_invalidKeys.ContainsKey(Key.Key))
            return false;

        if (_invalidShortcutKeys.ContainsKey(this))
            return false;

        if (!RequiresAlt && !RequiresCtrl)
        {
            if (!_validWithoutModifiersKeys.ContainsKey(Key.Key))
                return false;
        }

        if (!IsHotkeyAvailable(handle))
            return false;

        return true;
    }

    public override bool Equals(object? obj) => obj is CommandKey other && Equals(other);
    public override int GetHashCode() => HashCode.Combine(Key, RequiresCtrl, RequiresShift, RequiresAlt);
    public bool Equals(CommandKey? other) => other != null && other.Key == Key && other.RequiresCtrl == RequiresCtrl && other.RequiresShift == RequiresShift && other.RequiresAlt == RequiresAlt;
    public override string ToString()
    {
        var parts = new List<string>();
        if (RequiresCtrl)
        {
            parts.Add(Ctrl.DisplayOrSimplifiedName);
        }

        if (RequiresAlt)
        {
            parts.Add(Alt.DisplayOrSimplifiedName);
        }

        if (RequiresShift)
        {
            parts.Add(Shift.DisplayOrSimplifiedName);
        }

        if (HasNonModifierVirtualKey)
        {
            parts.Add(Key.DisplayOrSimplifiedName);
        }

        return string.Join(" + ", parts);
    }

    public static CommandKey? Parse(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return null;

        var parts = text.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        // last is virtual key
        var lastPart = parts[^1];
        if (!VirtualKey.AllByNames.TryGetValue(lastPart, out var vk))
        {
            // support "Ctrl + VK_K" notation
            if (Enum.TryParse<VIRTUAL_KEY>(lastPart, true, out var vkEnum))
            {
                vk = [vkEnum];
            }
            else
                return null;
        }

        if (vk[0].IsModifierKey())
            return null;

        var commandKey = new CommandKey(vk[0]);
        for (var i = 0; i < parts.Length - 1; i++)
        {
            // support localized names and generic names
            var part = parts[i];
            if (part.EqualsIgnoreCase(Ctrl.DisplayOrSimplifiedName) || part.EqualsIgnoreCase("Ctrl"))
            {
                commandKey.RequiresCtrl = true;
            }
            else if (part.EqualsIgnoreCase(Alt.DisplayOrSimplifiedName) || part.EqualsIgnoreCase("Alt"))
            {
                commandKey.RequiresAlt = true;
            }
            else if (part.EqualsIgnoreCase(Shift.DisplayOrSimplifiedName) || part.EqualsIgnoreCase("Shift"))
            {
                commandKey.RequiresShift = true;
            }
        }
        return commandKey;
    }

    public static CommandKey FromCharacter(char character)
    {
        var scan = DirectN.Functions.VkKeyScanW(character);
        var vk = (VIRTUAL_KEY)(scan & 0xFF);
        var shift = (scan & 0x100) != 0;
        var alt = (scan & 0x200) != 0;
        var ctrl = (scan & 0x400) != 0;
        return new CommandKey(vk)
        {
            RequiresShift = shift,
            RequiresAlt = alt,
            RequiresCtrl = ctrl,
        };
    }
}
