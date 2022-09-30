using System.Data;
using System.Data.Common;
using AdoNet.Specification.Tests;

namespace SimpleXlsxAdo.Specification.Tests;

public class XlsxSelectValueDeleteFixture : XlsxDbFactoryFixture, ISelectValueFixture, IDeleteFixture
{
    public string CreateSelectSql(DbType dbType, ValueKind kind)
    {
        return kind == ValueKind.Null ? kind.ToString() : $"{dbType.ToString()}_{kind.ToString()}";
    }

    public string CreateSelectSql(byte[] value)
    {
        return "";
    }

    public IReadOnlyCollection<DbType> SupportedDbTypes { get; } = new[] { DbType.Double, DbType.String, DbType.Boolean, DbType.DateTime }.AsReadOnly();
    public string SelectNoRows => "EmptyTable";
    public Type NullValueExceptionType { get; } = typeof(ArgumentNullException);
    public string DeleteNoRows => "EmptyTable";
}