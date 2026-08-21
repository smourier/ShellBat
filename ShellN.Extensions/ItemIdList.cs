namespace ShellN.Extensions;

public sealed class ItemIdList : IDisposable, IEnumerable<ItemId>, IEquatable<ItemIdList>, IItemWithAbsoluteIdList
{
    private static readonly Lazy<ItemIdList> _desktop = new(() => ShellFolder.Desktop.GetIdList()!);
    public static ItemIdList Desktop => _desktop.Value;

    // don't own, don't dispose
    public static ItemIdList? FromPointer(nint pidl, bool owned) => pidl == 0 ? null : new ItemIdList(pidl, owned);

    // owned, disposable
    public static ItemIdList? CloneFromPointer(nint pidl) => pidl == 0 ? null : new ItemIdList(Functions.ILClone(pidl), true);

    private readonly IntPtrBuffer _buffer;
    private readonly Lazy<int> _hashcode;

    public ItemIdList(nint pidl, bool owned)
    {
        if (pidl == 0)
            throw new ArgumentException(null, nameof(pidl));

        var size = Functions.ILGetSize(pidl);
        _buffer = new IntPtrBuffer(pidl, size, owned);
        _hashcode = new Lazy<int>(ComputeHashCode, true);
    }

    ItemIdList IItemWithAbsoluteIdList.AbsoluteIdList => this;
    ItemId? IItemWithId.Id => LastId;
    public IntPtrBuffer Buffer => _buffer;
    public nint Pointer => _buffer.DangerousGetHandle();
    public uint Size => (uint)_buffer.ByteLength;

    public unsafe nint RelativePointer
    {
        get
        {
            nint pointer;
            Functions.SHBindToParent(Pointer, typeof(IShellFolder).GUID, out var parent, (nint)(&pointer));
            return pointer;
        }
    }

    public string? DisplayName
    {
        get
        {
            if (IsTerminator)
                return null;

            Functions.SHGetNameFromIDList(Pointer, SIGDN.SIGDN_NORMALDISPLAY, out var str);
            return str.ToString();
        }
    }

    public string? FullParsingName
    {
        get
        {
            if (IsTerminator)
                return null;

            Functions.SHGetNameFromIDList(Pointer, SIGDN.SIGDN_DESKTOPABSOLUTEPARSING, out var str);
            return str.ToString();
        }
    }

    public string? FileSystemPath
    {
        get
        {
            if (IsTerminator)
                return null;

            Functions.SHGetNameFromIDList(Pointer, SIGDN.SIGDN_FILESYSPATH, out var str);
            return str.ToString();
        }
    }

    public bool IsTerminator
    {
        get
        {
            var pointer = Pointer;
            if (pointer == 0)
                return false;

            return Marshal.ReadInt16(pointer) == 0;
        }
    }

    public ItemIdList? Next
    {
        get
        {
            if (IsTerminator || Pointer == 0)
                return null;

            var next = Functions.ILGetNext(Pointer);
            if (next == 0)
                return null;

            return new ItemIdList(next, false);
        }
    }

    public ItemId? FirstId => this.FirstOrDefault();
    public ItemId? LastId
    {
        get
        {
            ItemId? last = null;
            foreach (var idl in this)
            {
                if (idl.IsTerminator)
                    return last;

                last = idl;
            }
            return last;
        }
    }

#if DEBUG
    public string? HexaDump => IsTerminator ? null : Conversions.ToHexaDump(Pointer, (int)Size);
#endif

    public ShellItem? GetItem(bool noFolder = false, bool owned = true, bool throwOnError = true) => ShellItem.FromItemIdList(this, noFolder, owned, throwOnError);

#if DEBUG
    public override string ToString()
    {
        if (Pointer == 0)
            return "<null>";

        if (IsTerminator)
            return "<term>";

        return string.Join(@"\", this);
    }
#else
    public override string ToString() => Pointer == IntPtr.Zero ? "<null>" : IsTerminator ? "<term>" : Pointer.ToHex();
#endif

    public byte[] ToByteArray() => _buffer.ToByteArray();
    public bool Equals(ItemIdList? other) => other is not null && _buffer.MemoryEqualsTo(other._buffer);
    public void Dispose() => _buffer.Dispose();

    private int ComputeHashCode()
    {
        var pointer = _buffer.DangerousGetHandle();
        if (pointer == 0 || _buffer.ByteLength == 0)
            return 0;

        return (int)Crc32.HashToUInt32(_buffer.GetSpan(17));
    }

    public override int GetHashCode() => _hashcode.Value;
    public override bool Equals(object? obj) => Equals(obj as ItemIdList);
    public static bool operator ==(ItemIdList idList1, ItemIdList idList2)
    {
        if (idList1 is null)
            return idList2 is null;

        return idList1.Equals(idList2);
    }

    public static bool operator !=(ItemIdList idList1, ItemIdList idList2)
    {
        if (idList1 is null)
            return idList2 is not null;

        return !idList1.Equals(idList2);
    }

    // includes the terminator
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    public IEnumerator<ItemId> GetEnumerator()
    {
        var current = Pointer;
        while (current != 0)
        {
            var size = ItemId.GetSize(current);
            var id = new ItemId(current, size, false);
            yield return id;
            if (size == 0)
                break;

            current = Functions.ILGetNext(current);
        }
    }
}
