using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Packaging;

namespace SimpleXlsxAdo;
public sealed class XlsxConnection : DbConnection
{
    XlsxConnectionString _settings;
    FileStream? _documentStream;
    SpreadsheetDocument? _document;

    public XlsxConnection(string? connectionString = null)
        => _settings = new XlsxConnectionString(connectionString);

    [AllowNull]
    public override string ConnectionString
    {
        get => _settings.ConnectionString;
        set
        {
            if (_document != null)
                throw new InvalidOperationException("Setting the connection string is not possible when the connection is open.");

            _settings = new XlsxConnectionString(value);
        }
    }

    public override string Database => _settings.Path ?? string.Empty;

    static readonly StateChangeEventArgs ClosedToOpenEventArgs = new(ConnectionState.Closed, ConnectionState.Open);
    static readonly StateChangeEventArgs OpenToClosedEventArgs = new(ConnectionState.Open, ConnectionState.Closed);
    public override ConnectionState State => _document == null ? ConnectionState.Closed : ConnectionState.Open;
    public override string DataSource => _settings.Path ?? string.Empty;
    public override string ServerVersion
    {
        get
        {
            if (_document == null)
                throw new InvalidOperationException("Connection is not open");

            var props = _document.ExtendedFilePropertiesPart?.Properties;
            if (props == null)
                return string.Empty;

            return (props.Application?.Text ?? string.Empty) + " " + (props.ApplicationVersion?.Text ?? string.Empty);
        }
    }

    protected override DbTransaction BeginDbTransaction(IsolationLevel isolationLevel)
        => throw new NotSupportedException("Transactions are not supported");

    public override void ChangeDatabase(string databaseName)
        => throw new NotSupportedException("Changing the database is not supported");

    public override void Close()
    {
        Dispose();
        OnStateChange(OpenToClosedEventArgs);
    }

    protected override void Dispose(bool disposing)
    {
        var doc = _document;
        var docStream = _documentStream;
        _document = null;
        _documentStream = null;
        doc?.Dispose();
        docStream?.Dispose();
    }

    public override void Open()
    {
        if (string.IsNullOrEmpty(_settings.ConnectionString))
            throw new InvalidOperationException("Please set a connection string before opening the connection.");
        if (_document != null)
            throw new InvalidOperationException("A connection cannot be opened twice.");
        if (!File.Exists(_settings.Path))
            throw new InvalidOperationException("Specify the path to your xlsx file in the connection string.");

        _documentStream = new(_settings.Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        _document = SpreadsheetDocument.Open(_documentStream, false, new OpenSettings{ AutoSave = false});
        OnStateChange(ClosedToOpenEventArgs);
    }

    public XlsxCommand CreateCommand(string? commandText = null)
        => new(this, commandText);

    protected override DbCommand CreateDbCommand()
        => CreateCommand();

    internal SpreadsheetDocument GetDocument()
        => _document ?? throw new InvalidOperationException("Connection is not open");

    internal bool Header => _settings.Header;
    internal string? NullString => _settings.NullString;

    protected override DbProviderFactory DbProviderFactory => XlsxProviderFactory.Instance;
}
