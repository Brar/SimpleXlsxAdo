using System.Collections;
using System.Data;
using System.Data.Common;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SimpleXlsxAdo;

public class XlsxReader : DbDataReader
{
    readonly XlsxConnection? _connection;
    readonly SpreadsheetDocument _document;
    readonly CommandBehavior _commandBehavior;
    readonly bool _header;
    readonly string? _nullString;
    readonly string[]? _sharedStringTable;
    readonly IEnumerator<WorksheetPart>? _worksheetPartEnumerator;
    readonly string? _workSheetName;
    readonly Dictionary<string, string> _workSheetNamesById = new();
    OpenXmlReader? _worksheetReader;
    string[]? _headers;
    Cell[]? _cells;
    bool _disposed;

    public XlsxReader(Stream documentStream, XlsxReaderOptions options)
        : this(SpreadsheetDocument.Open(documentStream, false, new() { AutoSave = false}), options) { }

    internal XlsxReader(XlsxConnection connection, XlsxReaderOptions options)
        : this(connection.GetDocument(), options)
    {
        if (options.CommandBehavior.HasFlag(CommandBehavior.CloseConnection))
            _connection = connection;
    }


    XlsxReader(SpreadsheetDocument document, XlsxReaderOptions options)
    {
        _document = document;
        _commandBehavior = options.CommandBehavior;
        _header = options.Header;
        _nullString = options.NullString;
        var wb = document.WorkbookPart;
        if (wb == null)
            throw new ArgumentException("The file does not contain a workbook part.");

        if (wb.SharedStringTablePart == null)
            _sharedStringTable = null;
        else
        {
            using var sstReader = OpenXmlReader.Create(wb.SharedStringTablePart);
            var table = new List<string>();
            while (sstReader.Read())
            {
                if (sstReader.ElementType != typeof(SharedStringItem))
                    continue;

                var text = new System.Text.StringBuilder();
                var sharedStringFound = false;
                if (!sstReader.ReadFirstChild())
                    throw new FormatException("Failed to read shared string table.");
                do
                {
                    if (sstReader.ElementType.IsAssignableTo(typeof(OpenXmlLeafTextElement)))
                    {
                        sharedStringFound = true;
                        text.Append(sstReader.GetText());
                    }
                    else if (sstReader.ElementType == typeof(Run)
                             || sstReader.ElementType == typeof(PhoneticRun)
                             || sstReader.ElementType == typeof(PhoneticProperties))
                    {
                        if (!sstReader.ReadFirstChild())
                            throw new FormatException("Failed to read shared string table.");
                        do
                        {
                            if (!sstReader.ElementType.IsAssignableTo(typeof(OpenXmlLeafTextElement)))
                                continue;

                            sharedStringFound = true;
                            text.Append(sstReader.GetText());
                        } while (sstReader.ReadNextSibling());
                    }
                } while (sstReader.ReadNextSibling());

                if (!sharedStringFound)
                    throw new FormatException("Failed to read shared string table.");

                table.Add(text.ToString());
            }

            _sharedStringTable = table.ToArray();
        }

        var worksheetName = options.WorksheetName;
        using var wbReader = OpenXmlReader.Create(wb);
        while (wbReader.Read())
        {
            if (wbReader.ElementType != typeof(Sheet))
                continue;

            var sheet = (Sheet)wbReader.LoadCurrentElement()!;
            if (worksheetName == null && sheet.Id?.Value != null && sheet.Name?.Value != null)
                _workSheetNamesById[sheet.Id.Value] = sheet.Name.Value;
            else
            {
                if (sheet.Name?.Value != worksheetName)
                    continue;

                var part = wb.GetPartById(sheet.Id?.Value ??
                                          throw new FormatException($"Sheet {worksheetName} does not have an id."));
                _worksheetReader = InitializeWorksheetReader(part as WorksheetPart ??
                                                             throw new FormatException(
                                                                 $"Part {sheet.Id.Value} is not a {nameof(WorksheetPart)}."));
                _workSheetName = worksheetName;
                return;
            }
        }

        if (worksheetName != null)
            throw new XlsxException($"The {nameof(Worksheet)} \"{worksheetName}\" was not found.");

        var worksheetPartEnumerator = wb.WorksheetParts.GetEnumerator();
        if (!worksheetPartEnumerator.MoveNext())
            throw new ArgumentException("The file does not contain a worksheet part.");
        _worksheetReader = InitializeWorksheetReader(worksheetPartEnumerator.Current);
        _worksheetPartEnumerator = worksheetPartEnumerator;
    }

    public override bool GetBoolean(int ordinal)
    {
        var cellText = GetCellText(ordinal);

        if (cellText == null || cellText == _nullString)
            throw new InvalidCastException($"Cannot cast {nameof(DBNull)} to {nameof(Boolean)}.");

        return cellText == "1";
    }

    public override byte GetByte(int ordinal)
    {
        throw new NotImplementedException();
    }

    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length)
    {
        throw new NotImplementedException();
    }

    public override char GetChar(int ordinal)
    {
        throw new NotImplementedException();
    }

    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length)
    {
        throw new NotImplementedException();
    }

    public override string GetDataTypeName(int ordinal)
    {
        throw new NotImplementedException();
    }

    public override DateTime GetDateTime(int ordinal)
    {
        throw new NotImplementedException();
    }

    public override decimal GetDecimal(int ordinal)
    {
        throw new NotImplementedException();
    }

    public override double GetDouble(int ordinal)
    {
        var cellText = GetCellText(ordinal);

        if (cellText == null || cellText == _nullString)
            throw new InvalidCastException($"Cannot cast {nameof(DBNull)} to {nameof(Double)}.");

        return double.Parse(cellText);
    }

    public override Type GetFieldType(int ordinal)
    {
        var cell = GetCell(ordinal);
        switch (cell.DataType?.Value)
        {
            case CellValues.Boolean:
                return typeof(bool);
            case CellValues.Error:
                return cell.CellValue?.Text == "#N/A"
                    ? typeof(DBNull)
                    : typeof(Exception);
            case CellValues.SharedString:
            case CellValues.String:
            case CellValues.InlineString:
                return typeof(string);
            case CellValues.Date:
                return typeof(DateTime);
            case CellValues.Number:
                return typeof(double);
            default:
                var cellText = GetCellText(cell);
                if (cellText == null || cellText == _nullString)
                    return typeof(DBNull);

                return typeof(double);
        }
    }

    public override float GetFloat(int ordinal)
    {
        throw new NotImplementedException();
    }

    public override Guid GetGuid(int ordinal)
    {
        throw new NotImplementedException();
    }

    public override short GetInt16(int ordinal)
    {
        throw new NotImplementedException();
    }

    public override int GetInt32(int ordinal)
    {
        throw new NotImplementedException();
    }

    public override long GetInt64(int ordinal)
    {
        throw new NotImplementedException();
    }

    public override string GetName(int ordinal)
    {
        CheckDisposed();
        var headers = _headers;
        if (headers == null)
            throw new InvalidOperationException();

        return headers[ordinal];
    }

    public override int GetOrdinal(string name)
    {
        CheckDisposed();
        var headers = _headers;
        if (headers == null)
            throw new InvalidOperationException();

        for (var i = headers.Length - 1; i >= 0; i--)
            if (headers[i] == name)
                return i;

        throw new IndexOutOfRangeException();
    }

    public override string GetString(int ordinal)
    {
        var cellText = GetCellText(ordinal);
        if (cellText == null || cellText == _nullString)
            throw new InvalidCastException($"Cannot cast {nameof(DBNull)} to {nameof(String)}.");

        return cellText;
    }

    public override object GetValue(int ordinal)
    {
        var cell = GetCell(ordinal);
        switch (cell.DataType?.Value)
        {
            case CellValues.Boolean:
            {
                var txt = GetCellText(cell);
                return txt == null 
                    ? DBNull.Value
                    : txt == "1";
            }
            case CellValues.Error:
            {
                if (cell.CellValue?.Text == "#N/A")
                    return DBNull.Value;

                var txt = GetCellText(cell);
                return txt == null
                    ? DBNull.Value
                    : new Exception(txt);
            }
            case CellValues.SharedString:
            case CellValues.String:
            case CellValues.InlineString:
            {
                var txt = GetCellText(cell);
                return txt == null || txt == _nullString
                    ? DBNull.Value
                    : txt;
            }
            case CellValues.Date:
            {
                var txt = GetCellText(cell);
                return txt == null
                    ? DBNull.Value
                    : DateTime.Parse(txt);
            }
            case CellValues.Number:
            case null:
            default:
            {
                var txt = GetCellText(cell);
                return txt == null
                    ? DBNull.Value
                    : double.Parse(txt);
            }
        }
    }

    public T? GetIParsable<T>(int ordinal, IFormatProvider? provider)
        where T : IParsable<T?>
    {
        CheckDisposed();
        var s = "";
        return T.Parse(s, provider);
    }

    public override int GetValues(object[] values)
    {
        throw new NotImplementedException();
    }

    public override bool IsDBNull(int ordinal)
    {
        var cell = GetCell(ordinal);
        
        if (cell.DataType?.Value == CellValues.Error
            && cell.CellValue?.Text == "#N/A")
            return true;

        var cellText = GetCellText(cell);
        return cellText == null || cellText == _nullString;
    }

    public override int FieldCount
    {
        get
        {
            CheckDisposed();
            return _cells?.Length ?? 0;
        }
    }

    public override object this[int ordinal] => GetValue(ordinal);

    public override object this[string name] => GetValue(GetOrdinal(name));

    public override int RecordsAffected => 0;

    public override bool HasRows
    {
        get
        {
            CheckDisposed();
            return _worksheetReader != null;
        }
    }

    public override bool IsClosed => _disposed;

    public string? WorkSheetName
    {
        get
        {
            if (_workSheetName != null)
                return _workSheetName;

            var currentPart = _worksheetPartEnumerator?.Current;
            return currentPart != null
                ? _workSheetNamesById[_document.WorkbookPart!.GetIdOfPart(currentPart)]
                : null;
        }
    }

    public override bool NextResult()
    {
        CheckDisposed();

        var enumerator = _worksheetPartEnumerator;
        if (enumerator == null)
            return false;

        if (!enumerator.MoveNext())
            return false;

        _worksheetReader = InitializeWorksheetReader(enumerator.Current);
        return true;
    }

    public override bool Read()
    {
        CheckDisposed();

        var reader = _worksheetReader;
        if (reader == null || !reader.ReadFirstChild())
            return false;

        var cells = new List<Cell>();
        if (_headers == null)
        {
            var newHeaders = new List<string>();
            do
            {
                if (reader.ElementType != typeof(Cell))
                    continue;

                var cell = (Cell) reader.LoadCurrentElement()!;
                cells.Add(cell);
                newHeaders.Add(cell.CellReference?.Value?.TrimEnd('1')
                               ?? throw new FormatException("Failed to read column name."));
            } while (reader.ReadNextSibling());
            _headers = newHeaders.ToArray();
        }
        else
            do
            {
                if (reader.ElementType != typeof(Cell))
                    continue;

                cells.Add((Cell) reader.LoadCurrentElement()!);
            } while (reader.ReadNextSibling());

        _cells = cells.ToArray();

        while (reader.Read())
        {
            if (reader.ElementType != typeof(Row))
                continue;

            return true;
        }

        _worksheetReader = null;
        reader.Dispose();
        return true;
    }

    OpenXmlReader? InitializeWorksheetReader(WorksheetPart part)
    {
        var reader = OpenXmlReader.Create(part);
        _headers = null;
        try
        {
            while (reader.Read())
            {
                if (reader.ElementType != typeof(Row))
                    continue;

                if (!_header || _headers != null)
                    return reader;

                if (!reader.ReadFirstChild())
                {
                    reader.Dispose();
                    return null;
                }
                var newHeaders = new List<string>();
                do
                {
                    if (reader.ElementType != typeof(Cell))
                        continue;

                    var cell = (Cell) reader.LoadCurrentElement()!;
                    var cellText = GetCellText(cell);
                    if (cellText != null && cellText != _nullString)
                        newHeaders.Add(cellText);
                    else
                        newHeaders.Add(cell.CellReference?.Value?.TrimEnd('1')
                                       ?? throw new FormatException("Failed to read column name."));
                } while (reader.ReadNextSibling());
                _headers = newHeaders.ToArray();
            }
            reader.Dispose();
            return null;
        }
        catch
        {
            reader.Dispose();
            throw;
        }
    }

    string? GetCellText(int ordinal)
    {
        CheckDisposed();
        var cells = _cells;
        if (cells == null)
            throw new InvalidOperationException("Call read first");

        var cell =  cells[ordinal];
        switch (cell.DataType?.Value ?? CellValues.Number)
        {
            case CellValues.Boolean:
            case CellValues.Number:
            case CellValues.Error:
            case CellValues.InlineString:
            case CellValues.Date:
                return cell.CellValue?.Text;
            case CellValues.SharedString:
                return _sharedStringTable?[
                           int.Parse(cell.CellValue?.Text ??
                                     throw new FormatException("The cell has shared string format but no value."))] ??
                       throw new InvalidOperationException($"Found shared string but no {nameof(SharedStringTable)}");
            case CellValues.String:
                throw new NotSupportedException("Cannot read cells containing a formula string.");
            default:
                throw new NotImplementedException();
        }
    }

    string? GetCellText(Cell cell)
    {
        switch (cell.DataType?.Value ?? CellValues.Number)
        {
            case CellValues.Boolean:
            case CellValues.Number:
            case CellValues.Error:
            case CellValues.InlineString:
            case CellValues.Date:
                return cell.CellValue?.Text;
            case CellValues.SharedString:
                return _sharedStringTable?[
                           int.Parse(cell.CellValue?.Text ??
                                     throw new FormatException("The cell has shared string format but no value."))] ??
                       throw new InvalidOperationException($"Found shared string but no {nameof(SharedStringTable)}");
            case CellValues.String:
                throw new NotSupportedException("Cannot read cells containing a formula string.");
            default:
                throw new NotImplementedException();
        }
    }

    Cell GetCell(int ordinal)
    {
        CheckDisposed();
        var cells = _cells;
        if (cells == null)
            throw new InvalidOperationException("Call read first");

        return cells[ordinal];
    }

    public override int Depth { get; }

    public override IEnumerator GetEnumerator()
    {
        CheckDisposed();
        throw new NotImplementedException();
    }

    void CheckDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(XlsxReader));
    }

    protected override void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing)
        {
            _worksheetReader?.Dispose();
            _worksheetPartEnumerator?.Dispose();
            _document.Dispose();
            _connection?.Close();
        }

        _disposed = true;
    }
}