using System.Data.Common;
using AdoNet.Specification.Tests;

namespace SimpleXlsxAdo.Specification.Tests;

public class XlsxReaderTests : DataReaderTestBase<XlsxSelectValueDeleteFixture>
{
    public XlsxReaderTests(XlsxSelectValueDeleteFixture fixture) : base(fixture)
    {
    }

    [Fact]
    public override void Depth_returns_zero()
    {
        using var connection = CreateOpenConnection();
        using var command = connection.CreateCommand();
        using var reader = command.ExecuteReader();
        Assert.Equal(0, reader.Depth);
    }

    [Fact]
    public override void Dispose_command_before_reader()
    {
        using var connection = CreateOpenConnection();
        DbDataReader reader;
        using (var command = connection.CreateCommand())
        {
            command.CommandText = "Table1";
            reader = command.ExecuteReader();
        }

        Assert.True(reader.Read());
        Assert.Equal("test", reader.GetString(0));
        Assert.False(reader.Read());
    }

    [Fact]
    public override void FieldCount_throws_when_closed()
        => X_throws_when_closed(
            r =>
            {
                var x = r.FieldCount;
            });

    [Fact]
    public override void FieldCount_works()
    {
        using var connection = CreateOpenConnection();
        using var command = connection.CreateCommand();
        using var reader = command.ExecuteReader();
        Assert.Equal(1, reader.FieldCount);
    }

    #region ExecuteScalar is not supported

    [Fact]
    public override void ExecuteScalar_returns_null_when_empty() { }

    #endregion

    void X_throws_when_closed(Action<DbDataReader> action)
    {
        using var connection = CreateOpenConnection();
        using var command = connection.CreateCommand();
        var reader = command.ExecuteReader();
        ((IDisposable) reader).Dispose();

        Assert.Throws<ObjectDisposedException>(() => action(reader));
    }
}
