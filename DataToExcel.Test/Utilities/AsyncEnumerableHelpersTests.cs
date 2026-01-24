using System.Data;
using DataToExcel.Utilities;
using Xunit;

namespace DataToExcel.Test.Utilities;

public class AsyncEnumerableHelpersTests
{
    [Fact]
    public async Task ToAsyncEnumerableReturnsItemsInOrder()
    {
        var items = new List<IDataRecord>
        {
            new FakeDataRecord("A"),
            new FakeDataRecord("B")
        };

        var results = new List<string>();
        await foreach (var record in AsyncEnumerableHelpers.ToAsyncEnumerable(items, CancellationToken.None))
        {
            results.Add(record.GetString(0));
        }

        Assert.Equal(new[] { "A", "B" }, results);
    }

    [Fact]
    public async Task ToAsyncEnumerableHonorsCancellation()
    {
        var items = new List<IDataRecord> { new FakeDataRecord("A") };
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(async () =>
        {
            await foreach (var _ in AsyncEnumerableHelpers.ToAsyncEnumerable(items, cts.Token))
            {
            }
        });
    }

    [Fact]
    public async Task BufferedAsyncRecordEnumeratorPeeksAndConsumes()
    {
        var items = new List<IDataRecord>
        {
            new FakeDataRecord("A"),
            new FakeDataRecord("B")
        };

        await using var enumerator = AsyncEnumerableHelpers.ToAsyncEnumerable(items, CancellationToken.None)
            .GetAsyncEnumerator();
        var buffered = new BufferedAsyncRecordEnumerator(enumerator);

        Assert.True(await buffered.TryPeekNextAsync());
        Assert.Equal("A", buffered.Current?.GetString(0));

        Assert.True(await buffered.TryGetNextAsync());
        Assert.Equal("A", buffered.Current?.GetString(0));

        Assert.True(await buffered.TryGetNextAsync());
        Assert.Equal("B", buffered.Current?.GetString(0));

        Assert.False(await buffered.TryGetNextAsync());
        Assert.Null(buffered.Current);
    }

    private sealed class FakeDataRecord : IDataRecord
    {
        private readonly string _value;

        public FakeDataRecord(string value)
            => _value = value;

        public int FieldCount => 1;
        public object this[int i] => _value;
        public object this[string name] => _value;
        public bool GetBoolean(int i) => false;
        public byte GetByte(int i) => 0;
        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) => 0;
        public char GetChar(int i) => _value[0];
        public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) => 0;
        public IDataReader GetData(int i) => throw new NotSupportedException();
        public string GetDataTypeName(int i) => "string";
        public DateTime GetDateTime(int i) => DateTime.MinValue;
        public decimal GetDecimal(int i) => 0;
        public double GetDouble(int i) => 0;
        public Type GetFieldType(int i) => typeof(string);
        public float GetFloat(int i) => 0;
        public Guid GetGuid(int i) => Guid.Empty;
        public short GetInt16(int i) => 0;
        public int GetInt32(int i) => 0;
        public long GetInt64(int i) => 0;
        public string GetName(int i) => "Value";
        public int GetOrdinal(string name) => 0;
        public string GetString(int i) => _value;
        public object GetValue(int i) => _value;
        public int GetValues(object[] values)
        {
            values[0] = _value;
            return 1;
        }
        public bool IsDBNull(int i) => false;
    }
}
