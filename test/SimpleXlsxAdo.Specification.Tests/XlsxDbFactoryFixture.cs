using System.Data.Common;
using AdoNet.Specification.Tests;

namespace SimpleXlsxAdo.Specification.Tests;

public class XlsxDbFactoryFixture : IDbFactoryFixture
{
    const string DefaultConnectionString = "Path=../../../SpecificationTests.xlsx";
    public DbProviderFactory Factory => XlsxProviderFactory.Instance;
    public string ConnectionString => DefaultConnectionString;
}