namespace SimpleXlsxAdo;

struct XlsxConnectionString
{
    readonly string? _connectionString;
    readonly XlsxConnectionStringBuilder _builder;

    public XlsxConnectionString(string? connectionString)
    {
        _connectionString = connectionString;
        _builder = new XlsxConnectionStringBuilder(connectionString);
    }

    public string ConnectionString => _connectionString ?? string.Empty;
    public string Path => _builder.Path;
    public bool Header => _builder.Header;
    public string? NullString => _builder.NullString;
}