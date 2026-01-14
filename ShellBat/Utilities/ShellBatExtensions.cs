namespace ShellBat.Utilities;

public static class ShellBatExtensions
{
    [return: NotNullIfNotNull(nameof(exception))]
    public static string? GetAllMessagesWithDots(this Exception? exception) => exception.GetAllMessages(".");

    public static void FillToDictionary(this IComObject<ICoreWebView2HttpHeadersCollectionIterator>? iterator, IDictionary<string, string?> dictionary) => FillToDictionary(iterator?.Object, dictionary);
    public static void FillToDictionary(this ICoreWebView2HttpHeadersCollectionIterator? iterator, IDictionary<string, string?> dictionary)
    {
        ArgumentNullException.ThrowIfNull(dictionary);
        if (iterator == null)
            return;

        BOOL currentHeader = false;
        while (iterator.get_HasCurrentHeader(ref currentHeader).IsSuccess && currentHeader)
        {
            iterator.GetCurrentHeader(out var namePtr, out var valuePtr).ThrowOnError();
            using var name = (Pwstr)namePtr;
            using var value = (Pwstr)valuePtr;
            dictionary[name.ToString() ?? string.Empty] = value.ToString();
            iterator.MoveNext(ref currentHeader).ThrowOnError();
            if (!currentHeader)
                break;
        }
    }

    public static char NibbleToChar(this int value, bool upper = true) => upper ? NibbleToCharUpper(value) : NibbleToCharLower(value);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static char NibbleToCharUpper(this int value)
    {
        value &= 0xF;
        value += '0';
        if (value > '9')
        {
            value += 'A' - ('9' + 1);
        }
        return (char)value;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static char NibbleToCharLower(this int value)
    {
        value &= 0xF;
        value += '0';
        if (value > '9')
        {
            value += 'a' - ('9' + 1);
        }
        return (char)value;
    }

    // equivalent to Tools.encodeFilePathForUrl
    [return: NotNullIfNotNull(nameof(path))]
    public static string? EncodePathForWebView2(this string? path) => path?.Replace("%", "%25").Replace("#", "%23");

    public static unsafe bool TryReceiveCopyData<T>(this LPARAM lParam, nuint expectedType, out T value) where T : unmanaged
    {
        value = default;
        if (lParam == 0)
            return false;

        var cds = (COPYDATASTRUCT*)lParam.Value;
        if (cds->dwData != expectedType || cds->cbData != (uint)sizeof(T))
            return false;

        value = *(T*)cds->lpData;
        return true;
    }

    public static T ReceiveCopyData<T>(this LPARAM lParam, nuint expectedType, T defaultValue) where T : unmanaged
    {
        if (!TryReceiveCopyData<T>(lParam, expectedType, out var value))
            return defaultValue;

        return value;
    }

    public static unsafe LRESULT SendCopyDataTo(this HWND hwnd, nuint type)
    {
        var cds = new COPYDATASTRUCT { dwData = type };
        return DirectN.Functions.SendMessageW(hwnd, MessageDecoder.WM_COPYDATA, 0, new((nint)(&cds)));
    }

    public static unsafe LRESULT SendCopyDataTo<T>(this HWND hwnd, nuint type, T value) where T : unmanaged
    {
        var cds = new COPYDATASTRUCT
        {
            lpData = (nint)(&value),
            cbData = (uint)sizeof(T),
            dwData = type
        };
        return DirectN.Functions.SendMessageW(hwnd, MessageDecoder.WM_COPYDATA, 0, new((nint)(&cds)));
    }

    public static Color ToColor(this D3DCOLORVALUE value) => Color.FromArgb(value.BA, value.BR, value.BG, value.BB);
    public static Size GetWinSize(this IComObject<IWICBitmapSource> bitmap) => GetWinSize(bitmap?.Object!);
    public static Size GetWinSize(this IWICBitmapSource bitmap)
    {
        ArgumentNullException.ThrowIfNull(bitmap);
        bitmap.GetSize(out var width, out var height);
        return new Size(width, height);
    }

    public static bool IsImage(this ShellItem? item, out bool isWebView2NativeImage, out bool isWicImage)
    {
        isWebView2NativeImage = false;
        isWicImage = false;
        if (item == null)
            return false;

        var extension = item.Extension.Nullify();
        if (extension == null)
            return false;

        var fe = FileExtension.GetByExtension(extension);
        if (fe == null)
            return false;

        isWebView2NativeImage = fe.IsWebView2NativeImage;
        isWicImage = fe.IsWicImage;
        return fe.IsImage;
    }

    public static bool IsFilePasteEnabled(this DataObject? dataObject)
    {
        if (dataObject == null)
            return false;

        var formats = dataObject.GetFormats();
        return formats.Any(f => f.EqualsIgnoreCase("FileDrop") ||
            f.EqualsIgnoreCase(ShellN.Constants.CFSTR_FILEDESCRIPTORW) ||
            f.EqualsIgnoreCase(Clipboard.CFSTR_SHELLIDLIST));
    }

    [return: NotNullIfNotNull(nameof(winRTObject))]
    public static IComObject<T>? AsComObject<T>(this object? winRTObject, CreateObjectFlags flags = CreateObjectFlags.UniqueInstance)
    {
        if (winRTObject == null)
            return null;

        var ptr = MarshalInspectable<object>.FromManaged(winRTObject);
        var obj = ComObject.FromPointer<T>(ptr, flags);
        return obj ?? throw new InvalidCastException($"Object of type '{winRTObject.GetType().FullName}' is not of type '{typeof(T).FullName}'.");
    }

    public static IComObject<T>? AsDxgiComObject<T>(this object? winRTObject, CreateObjectFlags flags = CreateObjectFlags.UniqueInstance)
    {
        if (winRTObject == null)
            return null;

        using var access = winRTObject.AsComObject<IDirect3DDxgiInterfaceAccess>(flags) ??
            throw new InvalidCastException($"Object of type '{winRTObject.GetType().FullName}' is not of type '{nameof(IDirect3DDxgiInterfaceAccess)}'.");

        access.Object.GetInterface(typeof(T).GUID, out var ptr).ThrowOnError();
        var obj = ComObject.FromPointer<T>(ptr, flags);
        return obj ?? throw new InvalidCastException($"Object of type '{winRTObject.GetType().FullName}' is not of type '{typeof(T).FullName}'.");
    }

    public static Size GetScaleFactor(this Size availableSize, Size contentSize)
    {
        var scaleX = contentSize.Width == 0 ? 0 : availableSize.Width / contentSize.Width;
        var scaleY = contentSize.Height == 0 ? 0 : availableSize.Height / contentSize.Height;
        var minscale = scaleX < scaleY ? scaleX : scaleY;
        scaleX = scaleY = minscale;
        return new Size(scaleX, scaleY);
    }

    public static string? GetEnumCommonPrefix(Type enumType)
    {
        var enumNames = Enum.GetNames(enumType);
        return GetCommonPrefix(enumNames);
    }

    private static string? GetCommonPrefix(string[] strings)
    {
        if (strings.Length == 0)
            return null;

        var prefix = strings[0];
        for (var i = 1; i < strings.Length; i++)
        {
            var str = strings[i];
            var j = 0;
            while (j < prefix.Length && j < str.Length && prefix[j] == str[j])
            {
                j++;
            }
            prefix = prefix[..j];
            if (string.IsNullOrEmpty(prefix))
                break;
        }
        return prefix.Nullify();
    }

    public static string FormatByteSize(long size)
    {
        using var pwstr = new AllocPwstr(128 * 2);
        DirectN.Functions.StrFormatByteSizeW(size, pwstr, pwstr.SizeInChars);
        return pwstr.ToString() ?? size.ToString();
    }

    public static string? GetNullifiedString(this NameValueCollection? collection, string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (collection == null)
            return null;

        return collection.Get(name).Nullify();
    }

    public static T? GetValue<T>(this NameValueCollection? collection, string name, IFormatProvider? provider = null, T? defaultValue = default)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (collection == null)
            return defaultValue;

        var str = collection.Get(name);
        if (str == null)
            return defaultValue;

        return Conversions.ChangeType(str, defaultValue, provider);
    }

    public static bool TryGetValue<T>(this NameValueCollection? collection, string name, out T? value) => TryGetValue(collection, name, null, out value);
    public static bool TryGetValue<T>(this NameValueCollection? collection, string name, IFormatProvider? provider, out T? value)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (collection == null)
        {
            value = default;
            return false;
        }

        var str = collection.Get(name);
        if (str == null)
        {
            value = default;
            return false;
        }

        return Conversions.TryChangeType(str, provider, out value);
    }

    public static string? GetNullifiedString(this IDictionary<string, string?>? dictionary, string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (dictionary == null)
            return null;

        if (!dictionary.TryGetValue(name, out var str))
            return null;

        return str.Nullify();
    }

    public static T? GetValue<T>(this IDictionary<string, string?>? dictionary, string name, IFormatProvider? provider = null, T? defaultValue = default)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (dictionary == null || !dictionary.TryGetValue(name, out var str))
            return defaultValue;

        return Conversions.ChangeType(str, defaultValue, provider);
    }

    public static bool TryGetValue<T>(this IDictionary<string, string?>? dictionary, string name, out T? value) => TryGetValue(dictionary, name, null, out value);
    public static bool TryGetValue<T>(this IDictionary<string, string?>? dictionary, string name, IFormatProvider? provider, out T? value)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (dictionary == null || !dictionary.TryGetValue(name, out var str))
        {
            value = default;
            return false;
        }

        return Conversions.TryChangeType(str, provider, out value);
    }

    public static T? GetValue<T>(this RegistryKey? key, string name, T? defaultValue = default, IFormatProvider? provider = null)
    {
        if (!TryGetValue<T>(key, name, provider, out var value))
            return defaultValue;

        return value;
    }

    public static string? GetNullifiedValue(this RegistryKey? key, string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return (key?.GetValue(name) as string).Nullify();
    }

    public static bool TryGetValue<T>(this RegistryKey? key, string name, out T? value) => TryGetValue(key, name, null, out value);
    public static bool TryGetValue<T>(this RegistryKey? key, string name, IFormatProvider? provider, out T? value)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (key is null)
        {
            value = default!;
            return false;
        }

        var regValue = key.GetValue(name);
        return Conversions.TryChangeType(regValue, provider, out value);
    }

    public static bool IsPressed(this VIRTUAL_KEY vk, bool async = true) => (async ? DirectN.Functions.GetAsyncKeyState((int)vk) : DirectN.Functions.GetKeyState((int)vk)) < 0;
    public static bool IsDigit(this VIRTUAL_KEY vk) => (vk >= VIRTUAL_KEY.VK_0 && vk <= VIRTUAL_KEY.VK_9) || IsNumericKeypadKey(vk);
    public static bool IsLetter(this VIRTUAL_KEY vk) => vk >= VIRTUAL_KEY.VK_A && vk <= VIRTUAL_KEY.VK_Z;
    public static bool IsControlKey(this VIRTUAL_KEY vk) => vk == VIRTUAL_KEY.VK_CONTROL || vk == VIRTUAL_KEY.VK_LCONTROL || vk == VIRTUAL_KEY.VK_RCONTROL;
    public static bool IsShiftKey(this VIRTUAL_KEY vk) => vk == VIRTUAL_KEY.VK_SHIFT || vk == VIRTUAL_KEY.VK_LSHIFT || vk == VIRTUAL_KEY.VK_RSHIFT;
    public static bool IsAltKey(this VIRTUAL_KEY vk) => vk == VIRTUAL_KEY.VK_MENU || vk == VIRTUAL_KEY.VK_LMENU || vk == VIRTUAL_KEY.VK_RMENU;
    public static bool IsModifierKey(this VIRTUAL_KEY vk) => vk.IsControlKey() || vk.IsShiftKey() || vk.IsAltKey();
    public static bool IsFunctionKey(this VIRTUAL_KEY vk) => vk >= VIRTUAL_KEY.VK_F1 && vk <= VIRTUAL_KEY.VK_F24;
    public static bool IsNavigationKey(this VIRTUAL_KEY vk) => vk == VIRTUAL_KEY.VK_LEFT || vk == VIRTUAL_KEY.VK_RIGHT ||
               vk == VIRTUAL_KEY.VK_UP || vk == VIRTUAL_KEY.VK_DOWN ||
               vk == VIRTUAL_KEY.VK_HOME || vk == VIRTUAL_KEY.VK_END ||
               vk == VIRTUAL_KEY.VK_PRIOR || vk == VIRTUAL_KEY.VK_NEXT;
    public static bool IsWhitespaceKey(this VIRTUAL_KEY vk) => vk == VIRTUAL_KEY.VK_SPACE || vk == VIRTUAL_KEY.VK_TAB || vk == VIRTUAL_KEY.VK_RETURN;
    public static bool IsEnterKey(this VIRTUAL_KEY vk) => vk == VIRTUAL_KEY.VK_RETURN;
    public static bool IsBackspaceKey(this VIRTUAL_KEY vk) => vk == VIRTUAL_KEY.VK_BACK || vk == VIRTUAL_KEY.VK_DELETE;
    public static bool IsEscapeKey(this VIRTUAL_KEY vk) => vk == VIRTUAL_KEY.VK_ESCAPE;
    public static bool IsNumericKeypadKey(this VIRTUAL_KEY vk) => vk >= VIRTUAL_KEY.VK_NUMPAD0 && vk <= VIRTUAL_KEY.VK_NUMPAD9 ||
               vk == VIRTUAL_KEY.VK_ADD || vk == VIRTUAL_KEY.VK_SUBTRACT ||
               vk == VIRTUAL_KEY.VK_MULTIPLY || vk == VIRTUAL_KEY.VK_DIVIDE ||
               vk == VIRTUAL_KEY.VK_DECIMAL || vk == VIRTUAL_KEY.VK_NUMLOCK;

    public static bool IsHotkeyAvailable(this VIRTUAL_KEY vk, HOT_KEY_MODIFIERS modifiers) => IsHotkeyAvailable(vk, modifiers, HWND.Null);
    public static bool IsHotkeyAvailable(this VIRTUAL_KEY vk, HOT_KEY_MODIFIERS modifiers, HWND hwnd)
    {
        const int HOTKEY_ID = 1;
        if (DirectN.Functions.RegisterHotKey(hwnd, HOTKEY_ID, modifiers, (uint)vk))
        {
            DirectN.Functions.UnregisterHotKey(hwnd, HOTKEY_ID);
            return true;
        }
        return false;
    }

    public static void OpenRecycleBin() => Process.Start(new ProcessStartInfo
    {
        FileName = $"shell:::{ShellN.Constants.CLSID_RecycleBin:B}",
        UseShellExecute = true,
    });

    public static long ToUnixTimeMilliseconds(this DateTime dateTime)
    {
        try
        {
            return new DateTimeOffset(dateTime).ToUnixTimeMilliseconds();
        }
        catch
        {
            return 0;
        }
    }

    public static IReadOnlyList<string> Shorten(this IEnumerable<string> paths, string sep = @" \ ")
    {
        ArgumentNullException.ThrowIfNull(paths);

        var split = paths
            .Select(p => p.Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar], StringSplitOptions.RemoveEmptyEntries))
            .ToList();

        var results = new List<string>();
        for (var i = 0; i < split.Count; i++)
        {
            var current = split[i];
            var take = 1;

            while (take <= current.Length)
            {
                var suffix = current.Skip(current.Length - take);
                var duplicates = split.Count(s => s.Skip(s.Length - take).SequenceEqual(suffix));
                if (duplicates == 1)
                    break;

                take++;
            }

            results.Add(string.Join(sep, current.Skip(current.Length - take)));
        }
        return results.AsReadOnly();
    }
}
