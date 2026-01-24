using System.Data;
using System.Runtime.CompilerServices;

namespace DataToExcel.Utilities;

internal sealed class BufferedAsyncRecordEnumerator
{
    private readonly IAsyncEnumerator<IDataRecord> _inner;
    private bool _hasBuffered;
    public IDataRecord? Current { get; private set; }

    public BufferedAsyncRecordEnumerator(IAsyncEnumerator<IDataRecord> inner)
        => _inner = inner;

    public async Task<bool> TryGetNextAsync()
    {
        if (_hasBuffered)
        {
            _hasBuffered = false;
            return true;
        }

        if (await _inner.MoveNextAsync())
        {
            Current = _inner.Current;
            return true;
        }

        Current = null;
        return false;
    }

    public async Task<bool> TryPeekNextAsync()
    {
        if (_hasBuffered)
            return true;

        if (await _inner.MoveNextAsync())
        {
            Current = _inner.Current;
            _hasBuffered = true;
            return true;
        }

        Current = null;
        return false;
    }
}

internal static class AsyncEnumerableHelpers
{
    public static async IAsyncEnumerable<IDataRecord> ToAsyncEnumerable(IEnumerable<IDataRecord> data,
        [EnumeratorCancellation] CancellationToken ct)
    {
        foreach (var record in data)
        {
            ct.ThrowIfCancellationRequested();
            yield return record;
            await Task.Yield();
        }
    }
}
