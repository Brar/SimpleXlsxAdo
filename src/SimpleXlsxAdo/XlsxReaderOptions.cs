using System.Data;

namespace SimpleXlsxAdo;

public record XlsxReaderOptions
{
    public string? WorksheetName { get; init; }
    public bool Header { get; init; }
    public string? NullString { get; init; } = string.Empty;
    public CommandBehavior CommandBehavior { get; init; } = CommandBehavior.CloseConnection;
}