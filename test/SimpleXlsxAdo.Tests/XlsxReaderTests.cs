using System.Data;

namespace SimpleXlsxAdo.Tests;

public class XlsxReaderTests
{
    [Test]
    public void HasRows([Values]bool hasRows)
    {
        using var reader = OpenReader(1, hasRows? "ABC" : "123");
        Assert.That(reader.HasRows, Is.EqualTo(hasRows));
    }

    [Test]
    public void Read([Values]bool hasRows)
    {
        using var reader = OpenReader(1, hasRows? "ABC" : "123");
        Assert.That(reader.Read(), Is.EqualTo(hasRows));
    }

    [Test]
    public void NextResult([Values]bool multipleWorkSheets)
    {
        using var reader = OpenReader(1, multipleWorkSheets? null : "ABC");
        Assert.That(reader.NextResult(), Is.EqualTo(multipleWorkSheets));
    }

    [Test]
    public void FieldCount()
    {
        using var reader = OpenReader(1, "ABC");
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.FieldCount, Is.EqualTo(3));
    }

    [Test]
    public void GetName([Values] bool header)
    {
        const int documentId = 1;
        const string workSheetName = "ABC";
        var expectedHeaders = GetExpectedHeaders(documentId, workSheetName, header);
        using var reader = OpenReader(documentId, workSheetName, header);
        Assert.That(reader.Read(), Is.True);
        for (var i = reader.FieldCount - 1; i >= 0; i--)
            Assert.That(reader.GetName(i), Is.EqualTo(expectedHeaders[i]));
        Assert.That(() => reader.GetName(reader.FieldCount), Throws.Exception.TypeOf<IndexOutOfRangeException>());
        Assert.That(reader.Read(), Is.True);
        for (var i = reader.FieldCount - 1; i >= 0; i--)
            Assert.That(reader.GetName(i), Is.EqualTo(expectedHeaders[i]));
        Assert.That(() => reader.GetName(reader.FieldCount), Throws.Exception.TypeOf<IndexOutOfRangeException>());
    }
    
    [Test]
    public void GetName_With_Complex_Header()
    {
        const int documentId = 1;
        const string workSheetName = "xyz";
        var expectedHeaders = GetExpectedHeaders(documentId, workSheetName, header: true);
        using var reader = OpenReader(documentId, workSheetName, header: true);
        Assert.That(reader.Read(), Is.False);
        for (var i = 2; i >= 0; i--)
            Assert.That(reader.GetName(i), Is.EqualTo(expectedHeaders[i]));
        Assert.That(() => reader.GetName(3), Throws.Exception.TypeOf<IndexOutOfRangeException>());
    }

    [Test]
    public void GetOrdinal([Values] bool header)
    {
        const int documentId = 1;
        const string workSheetName = "ABC";
        var expectedHeaders = GetExpectedHeaders(documentId, workSheetName, header);
        using var reader = OpenReader(documentId, workSheetName, header);
        Assert.That(reader.Read(), Is.True);
        for (var i = reader.FieldCount - 1; i >= 0; i--)
            Assert.That(reader.GetOrdinal(expectedHeaders[i]), Is.EqualTo(i));
        Assert.That(() => reader.GetOrdinal("Non-Existing"), Throws.Exception.TypeOf<IndexOutOfRangeException>());
        Assert.That(reader.Read(), Is.True);
        for (var i = reader.FieldCount - 1; i >= 0; i--)
            Assert.That(reader.GetOrdinal(expectedHeaders[i]), Is.EqualTo(i));
        Assert.That(() => reader.GetOrdinal("Non-Existing"), Throws.Exception.TypeOf<IndexOutOfRangeException>());
    }

    [Test]
    public void ReadUnknownWorkBook()
    {
        using var reader = OpenReader(1);
        do
        {
            Console.WriteLine($"WorkBook: \"{reader.WorkSheetName}\"");
            int rowNumber = 1;
            while (reader.Read())
            {
                Console.WriteLine($"\tRow: {rowNumber++}");

                for (var i = 0; i < reader.FieldCount; i++)
                {
                    Console.WriteLine($"\t\tColumn: {(i + 1)}");
                    var value = reader.GetValue(i);
                    var type = reader.GetFieldType(i);
                    var columnName = reader.GetName(i);
                    if (reader.IsDBNull(i))
                    {
                        Assert.That(value, Is.EqualTo(DBNull.Value));
                        Assert.That(type, Is.EqualTo(typeof(DBNull)));
                    }
                    else
                    {
                        switch (value)
                        {
                            case double doubleValue:
                                Assert.That(type, Is.EqualTo(typeof(double)));
                                var doubleResult = reader.GetDouble(i);
                                Assert.That(doubleValue, Is.EqualTo(doubleResult));
                                break;
                            case string stringValue:
                                Assert.That(type, Is.EqualTo(typeof(string)));
                                var stringResult = reader.GetString(i);
                                Assert.That(stringValue, Is.EqualTo(stringResult));
                                break;
                            case bool boolValue:
                                Assert.That(type, Is.EqualTo(typeof(bool)));
                                var boolResult = reader.GetBoolean(i);
                                Assert.That(boolValue, Is.EqualTo(boolResult));
                                break;
                            case Exception exceptionValue:
                                Assert.That(type, Is.EqualTo(typeof(Exception)));
                                var exceptionResult = reader.GetFieldValue<Exception>(i);
                                Assert.That(exceptionValue.ToString(), Is.EqualTo(exceptionResult.ToString()));
                                break;
                            case DateTime dateTimeValue:
                                Assert.That(type, Is.AssignableTo(typeof(DateTime)));
                                var dateTimeResult = reader.GetDateTime(i);
                                Assert.That(dateTimeValue, Is.EqualTo(dateTimeResult));
                                break;
                        }
                    }
                }
            }
        } while (reader.NextResult());
    }

    static string[] GetExpectedHeaders(int documentId, string sheetName, bool header)
        => header
            ? documentId switch
            {
                1 => sheetName switch
                {
                    "ABC" => new[] {"Header 1", "Header 2", "Header 3"},
                    "xyz" => new[] {"Header A", "Header B", "Header C"},
                    _ => throw new ArgumentOutOfRangeException(nameof(sheetName), sheetName, null)
                },
                _ => throw new ArgumentOutOfRangeException(nameof(documentId), documentId, null)
            }
            : new[] {"A", "B", "C"};

    static XlsxReader OpenReader(int documentId, string? workSheetName = null, bool header = false)
    {
        var path = Path.GetFullPath($"../../../TestDocument_{documentId:000}.xlsx");
        var conn = new XlsxConnection($"Path={path};Header={header}");
        conn.Open();
        Console.WriteLine($"Application version: {conn.ServerVersion}");
        return conn.CreateCommand(workSheetName).ExecuteReader(CommandBehavior.CloseConnection);
    }
}