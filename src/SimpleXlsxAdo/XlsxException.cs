using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.Serialization;

namespace SimpleXlsxAdo;

public sealed class XlsxException : DbException
{
    internal XlsxException() { }
    internal XlsxException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    internal XlsxException(string? message) : base(message) { }
    internal XlsxException(string? message, Exception? innerException) : base(message, innerException) { }
    internal XlsxException(string? message, int errorCode) : base(message, errorCode) { }
}