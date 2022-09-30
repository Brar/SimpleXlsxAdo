using System.Data.Common;

namespace SimpleXlsxAdo;

public class XlsxProviderFactory : DbProviderFactory
{
    XlsxProviderFactory() { }
    public static readonly XlsxProviderFactory Instance = new();
    public override DbConnectionStringBuilder CreateConnectionStringBuilder() => new XlsxConnectionStringBuilder();
    public override DbConnection CreateConnection() => new XlsxConnection();
    public override DbCommand CreateCommand() => new XlsxCommand();
}