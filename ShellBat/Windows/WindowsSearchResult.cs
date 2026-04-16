namespace ShellBat.Windows;

public sealed partial class WindowsSearchResult : InterlockedComObject<IRowset>
{
    private readonly IComObject<IGetRow> _getRow;

    public WindowsSearchResult(IComObject<IRowset> rowset)
        : base(rowset)
    {
        // tripoli does that for us (IPropertyStore) so we don't need to do too much of dusty OLE DB :-)
        _getRow = ComObject.As<IGetRow>(throwOnError: true)!;
    }

    public Exception? EnumerationError { get; private set; }

    public IEnumerable<WindowsSearchRow> EnumerateRows(int batchSize = 1)
    {
        batchSize = Math.Max(1, batchSize);
        var accessor = ComObject.As<IAccessor>()!;
        if (accessor == null)
            yield break;

        var index = 0;
        while (true)
        {
            var hr = NativeObject.GetNextRows(0, index, batchSize, out var rowsObtained, out var rows);
            if (hr.IsError)
            {
                EnumerationError = ComError.GetError(hr) ?? Marshal.GetExceptionForHR(hr);
                yield break;
            }

            if (rowsObtained == 0)
                yield break;

            try
            {
                for (var idx = 0; idx < (int)rowsObtained; idx++)
                {
                    var hrow = GetHRow(rows, idx);

                    _getRow.Object.GetRowFromHROW(idx, hrow, typeof(ShellN.IPropertyStore).GUID, out var psUnk);
                    var store = new PropertyStore(DirectN.Extensions.Com.ComObject.FromPointer<ShellN.IPropertyStore>(psUnk)!);
                    var row = new WindowsSearchRow(index++, store);
                    yield return row;
                }
            }
            finally
            {
                var hrows = GetHRowsPointer(rows, (int)rowsObtained);
                NativeObject.ReleaseRows(rowsObtained, hrows, null!, null!, null!).ThrowOnError();
            }
        }
    }

    private static unsafe nuint GetHRow(nint rows, int idx) => *((nuint*)rows + idx);
    private static unsafe nuint[] GetHRowsPointer(nint rows, int count)
    {
        var hrows = new nuint[count];
        fixed (nuint* pHRows = hrows)
        {
            Buffer.MemoryCopy((void*)rows, pHRows, count * sizeof(nuint), count * sizeof(nuint));
        }
        return hrows;
    }
}
