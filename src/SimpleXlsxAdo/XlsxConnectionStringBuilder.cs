using System.Data.Common;

namespace SimpleXlsxAdo;

public sealed class XlsxConnectionStringBuilder : DbConnectionStringBuilder
{
    HashSet<string> _validKeys = new HashSet<string>
    {
        nameof(Path).ToLowerInvariant(),
        nameof(Header).ToLowerInvariant(),
        NullStringName.ToLowerInvariant(),
    };
    public XlsxConnectionStringBuilder(string? connectionString = null)
    {
        ConnectionString = connectionString;
        foreach (string key in Keys)
            if (!_validKeys.Contains(key))
                throw new ArgumentException($"Invalid keyword '{key}' in connection string.", nameof(connectionString));
    }

    public string Path
    {
        get => ContainsKey(nameof(Path)) ? (string)this[nameof(Path)] : string.Empty;
        set => this[nameof(Path)] = value;
    }

    public bool Header
    {
        get => ContainsKey(nameof(Header)) && (bool.TryParse((string)this[nameof(Header)], out var header)
            ? header
            : throw new FormatException($"Cannot convert value {this[nameof(Header)]} to {nameof(Boolean)}"));
        set => this[nameof(Header)] = value.ToString().ToLower();
    }

    const string NullStringName = "Null String";
    public string? NullString
    {
        get => ContainsKey(NullStringName) ? (string)base[NullStringName] : null;
        set
        {
            if (value == null)
                Remove(NullStringName);
            else
                this[NullStringName] = value;        }
    }
}