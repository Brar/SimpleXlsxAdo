using System.Data;
using AdoNet.Specification.Tests;

namespace SimpleXlsxAdo.Specification.Tests;

public class XlsxCommandTests : CommandTestBase<XlsxDbFactoryFixture>
{
    public XlsxCommandTests(XlsxDbFactoryFixture fixture) : base(fixture) { }

    #region CommandText is optional

    [Fact]
    public override void ExecuteReader_throws_when_no_command_text() { }

    [Fact]
    public override Task ExecuteReaderAsync_throws_when_no_command_text() => Task.CompletedTask;

    #endregion

    #region CommandText can only be a table name

    [Fact]
    public override void CommandText_throws_when_set_when_open_reader()
    {
        using var connection = CreateOpenConnection();
        using var command = connection.CreateCommand();
        using var reader = command.ExecuteReader();

        Assert.Throws<InvalidOperationException>(() => command.CommandText = "Table1");
    }

    [Fact]
    public override void Connection_throws_when_set_to_null_when_open_reader()
    {
        using var connection = CreateOpenConnection();
        using var command = connection.CreateCommand();
        using var reader = command.ExecuteReader();

        Assert.Throws<InvalidOperationException>(() => command.Connection = null);
    }

    [Fact]
    public override void Connection_throws_when_set_when_open_reader()
    {
        using var connection = CreateOpenConnection();
        using var command = connection.CreateCommand();
        using var reader = command.ExecuteReader();

        Assert.Throws<InvalidOperationException>(() => command.Connection = CreateConnection());
    }

    [Fact]
    public override void ExecuteReader_HasRows_is_false_for_comment() { }


    [Fact]
    public override void ExecuteReader_supports_CloseConnection()
    {
        using var connection = CreateOpenConnection();
        using var command = connection.CreateCommand();
        using (command.ExecuteReader(CommandBehavior.CloseConnection))
        {
        }
        Assert.Equal(ConnectionState.Closed, connection.State);
    }
    
    [Fact]
    public override void ExecuteReader_throws_when_reader_open()
    {
        using var connection = CreateOpenConnection();
        using var command = connection.CreateCommand();
        using var reader = command.ExecuteReader();
        Assert.ThrowsAny<InvalidOperationException>(() => command.ExecuteReader());
    }

    [Fact]
    public override void ExecuteReader_works_when_trailing_comments() { }

    #endregion

    #region Default CommandType is TableDirect not Text

    [Fact]
    public override void CommandType_text_by_default()
    {
        using var command = Fixture.Factory.CreateCommand();
        Assert.Equal(CommandType.TableDirect, command!.CommandType);
    }

    [Fact]
    public override void CommandType_does_not_throw_when_disposed()
    {
        var command = Fixture.Factory.CreateCommand();
        command!.Dispose();
        Assert.Equal(CommandType.TableDirect, command.CommandType);
    }

    [Fact]
    public override Task ExecuteScalarAsync_throws_when_no_connection() => Task.CompletedTask;

    [Fact]
    public override Task ExecuteScalarAsync_throws_when_connection_closed() => Task.CompletedTask;

    [Fact]
    public override Task ExecuteScalarAsync_throws_when_no_command_text() => Task.CompletedTask;

    [Fact]
    public override Task ExecuteScalarAsync_throws_on_error() => Task.CompletedTask;

    #endregion

    #region ExecuteNonQuery is not supported

    [Fact]
    public void ExecuteNonQuery_throws()
    {
	    using var connection = CreateOpenConnection();
	    using var command = connection.CreateCommand();
	    Assert.Throws<NotSupportedException>(() => command.ExecuteNonQuery());
    }

    [Fact]
    public override void ExecuteNonQuery_returns_negative_one_for_SELECT() { }

    [Fact]
    public override void ExecuteNonQuery_throws_when_connection_closed() { }

    [Fact]
    public override void ExecuteNonQuery_throws_when_no_command_text() { }

    [Fact]
    public override void ExecuteNonQuery_throws_when_no_connection() { }

    [Fact]
    public override Task ExecuteNonQueryAsync_is_canceled() => Task.CompletedTask;

    #endregion

    #region ExecuteScalar is not supported

    [Fact]
    public void ExecuteScalar_throws()
    {
	    using var connection = CreateOpenConnection();
	    using var command = connection.CreateCommand();
	    Assert.Throws<NotSupportedException>(() => command.ExecuteScalar());
    }

    [Fact]
    public override void Execute_throws_for_unknown_ParameterValue_type() { }

	[Fact]
	public override void ExecuteScalar_throws_when_no_connection() { }

	[Fact]
	public override void ExecuteScalar_throws_when_connection_closed() { }

	[Fact]
	public override void ExecuteScalar_throws_when_no_command_text() { }

	[Fact]
	public override void ExecuteScalar_returns_integer() { }

	[Fact]
	public override void ExecuteScalar_returns_real() { }

	[Fact]
	public override void ExecuteScalar_returns_string_when_text() { }

	[Fact]
	public override void ExecuteScalar_returns_DBNull_when_null() { }

	[Fact]
	public override void ExecuteScalar_returns_first_when_batching() { }

	[Fact]
	public override void ExecuteScalar_returns_first_when_multiple_columns() { }

	[Fact]
	public override void ExecuteScalar_returns_first_when_multiple_rows() { }

    #endregion

    #region Parameters are not supported

    [Fact]
    public void Parameters_get_throws()
    {
	    using var connection = CreateOpenConnection();
	    using var command = connection.CreateCommand();
	    Assert.Throws<NotSupportedException>(() => command.Parameters.ToString());
    }

    [Fact]
    public void CreateParameter_throws()
    {
	    using var connection = CreateOpenConnection();
	    using var command = connection.CreateCommand();
	    Assert.Throws<NotSupportedException>(() => command.CreateParameter());
    }

    [Fact]
    public override void CreateParameter_is_not_null() { }

    [Fact]
    public override void ExecuteReader_binds_parameters() { }

    [Fact]
    public override void Parameters_is_not_null() { }

    [Fact]
    public override void Parameters_returns_same_object() { }

    [Fact]
    public override void Parameters_does_not_throw_when_disposed() { }

    #endregion

    #region Transactions are not supported

    [Fact]
    public void Transaction_get_throws()
    {
	    using var connection = CreateOpenConnection();
	    using var command = connection.CreateCommand();
	    Assert.Throws<NotSupportedException>(() => command.Transaction?.ToString());
    }

    [Fact]
    public void Transaction_set_throws()
    {
	    using var connection = CreateOpenConnection();
	    using var command = connection.CreateCommand();
	    Assert.Throws<NotSupportedException>(() => command.Transaction = null);
    }

    [Fact]
    public override void ExecuteReader_throws_when_transaction_required() { }

    [Fact]
    public override void ExecuteReader_throws_when_transaction_mismatched() { }
    
    #endregion

    #region Prepare is not supported

    [Fact]
    public void Prepare_throws()
    {
	    using var connection = CreateOpenConnection();
	    using var command = connection.CreateCommand();
	    Assert.Throws<NotSupportedException>(() => command.Prepare());
    }

    [Fact]
    public override void Prepare_throws_when_no_connection() { }

    [Fact]
    public override void Prepare_throws_when_connection_closed() { }

    [Fact]
    public override void Prepare_throws_when_no_command_text() { }

    #endregion
}
