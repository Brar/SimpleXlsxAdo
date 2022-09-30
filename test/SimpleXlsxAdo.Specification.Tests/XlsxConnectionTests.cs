using AdoNet.Specification.Tests;

namespace SimpleXlsxAdo.Specification.Tests;

public class XlsxConnectionTests : ConnectionTestBase<XlsxDbFactoryFixture>
{
    public XlsxConnectionTests(XlsxDbFactoryFixture fixture) : base(fixture)
    { }

    #region Transactions are not supported

    [Fact]
    public void BeginTransaction_throws()
    {
        using var connection = CreateOpenConnection();
        Assert.Throws<NotSupportedException>(() => { using var transaction = connection.BeginTransaction(); });

    }

    [Fact]
    public override void CreateCommand_does_not_set_Transaction_property() { }

    #endregion
}
