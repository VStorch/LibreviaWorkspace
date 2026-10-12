using System.Globalization;

namespace Librevia.Format.Pdf;

/// <summary>The words of the PDF language this module needs to read and write metadata.</summary>
public sealed record PdfName(string Value);

/// <summary>Raw bytes: what they mean depends on the encoding, decided in <see cref="PdfText"/>.</summary>
public sealed class PdfString(byte[] bytes)
{
    public byte[] Bytes { get; } = bytes;
}

/// <summary>Kept as written, so a number read is written back the same.</summary>
public sealed record PdfNumber(string Raw)
{
    public long? Integer =>
        long.TryParse(Raw, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var value) ? value : null;

    public static PdfNumber Of(long value) => new(value.ToString(CultureInfo.InvariantCulture));
}

public sealed record PdfRef(int Number, int Generation);

public sealed class PdfNull
{
    public static readonly PdfNull Instance = new();

    private PdfNull()
    {
    }
}

public sealed class PdfArray : List<object>;

/// <summary>In the file's order, so a rewritten dictionary keeps the keys where they were.</summary>
public sealed class PdfDictionary
{
    private readonly List<KeyValuePair<string, object>> _entries = [];

    public IEnumerable<KeyValuePair<string, object>> Entries => _entries;

    public object? this[string key]
    {
        get => _entries.FirstOrDefault(entry => entry.Key == key).Value;
        set
        {
            var index = _entries.FindIndex(entry => entry.Key == key);
            if (value is null)
            {
                if (index >= 0) _entries.RemoveAt(index);
            }
            else if (index >= 0)
            {
                _entries[index] = new(key, value);
            }
            else
            {
                _entries.Add(new(key, value));
            }
        }
    }

    public PdfDictionary Copy()
    {
        var copy = new PdfDictionary();
        foreach (var (key, value) in _entries) copy[key] = value;
        return copy;
    }
}

/// <summary>The data as it is in the file, still encoded by its filters.</summary>
public sealed record PdfStream(PdfDictionary Dictionary, byte[] Data);

public sealed class PdfException(string message) : Exception(message);
