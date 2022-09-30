using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;

namespace SimpleXlsxAdo;

public sealed class XlsxCommand : DbCommand
{
    XlsxConnection? _xlsxConnection;
    XlsxReader? _reader;
    string? _commandText;

    internal XlsxCommand(XlsxConnection? xlsxConnection = null, string? commandText = null)
    {
        _xlsxConnection = xlsxConnection;
        _commandText = commandText;
    }

    public override void Cancel() { }

    public override int ExecuteNonQuery() => throw new NotSupportedException();

    public override object? ExecuteScalar() => throw new NotSupportedException();

    public override void Prepare() => throw new NotSupportedException();

    [AllowNull]
    public override string CommandText
    {
        get => _commandText ?? string.Empty;
        set
        {
            if (_reader != null)
                throw new InvalidOperationException($"Cannot set {nameof(CommandText)} while a reader is open.");
            _commandText = value;
        }
    }

    public override int CommandTimeout { get; set; }

    public override CommandType CommandType
    {
        get => CommandType.TableDirect;
        set
        {
            if (value != CommandType.TableDirect)
                throw new NotSupportedException(
                    $"Only {nameof(CommandType)}.{nameof(CommandType.TableDirect)} is supported.");
        }
    }

    public override UpdateRowSource UpdatedRowSource
    {
        get => UpdateRowSource.None;
        set
        {
            if (value != UpdateRowSource.None)
                throw new NotSupportedException($"Setting {nameof(UpdatedRowSource)} is not supported");
        }
    }

    protected override DbConnection? DbConnection
    {
        get => _xlsxConnection;
        set
        {
            if (_reader != null)
                throw new InvalidOperationException($"Cannot set {nameof(DbConnection)} while a reader is open.");
            _xlsxConnection = (XlsxConnection)value!;
        }
    }

    protected override DbParameterCollection DbParameterCollection
        => throw new NotSupportedException("Queries and parameters are not supported.");

    protected override DbTransaction? DbTransaction
    {
        get => throw new NotSupportedException("Transactions are not supported");
        set => throw new NotSupportedException("Transactions are not supported");
    }

    public override bool DesignTimeVisible { get; set; }

    protected override DbParameter CreateDbParameter()
        => throw new NotSupportedException("Parameters are not supported");

    protected override DbDataReader ExecuteDbDataReader(CommandBehavior behavior)
        => ExecuteReader(behavior);

    public new XlsxReader ExecuteReader(CommandBehavior behavior = CommandBehavior.Default)
    {
        if (_reader != null)
            throw new InvalidOperationException($"A reader is already open.");

        _reader = new(
            _xlsxConnection ?? throw new InvalidOperationException("You need to set a connection to start reading."),
            new()
            {
                Header = _xlsxConnection.Header,
                NullString = _xlsxConnection.NullString,
                WorksheetName = _commandText,
                CommandBehavior = behavior
            });
        return _reader;
    }
}