using System.Text;

namespace Librevia.Format.Pdf;

/// <summary>
/// The file's structure: the newest trailer and where each object lives, merging every
/// cross-reference section from the last one back through <c>/Prev</c>.
/// </summary>
internal sealed class PdfFile
{
    /// <summary>Against a <c>/Prev</c> chain that loops or never ends.</summary>
    private const int MaxSections = 1000;

    private readonly byte[] _data;
    private readonly Dictionary<int, Location> _locations = [];
    private readonly Dictionary<int, (byte[] Data, Dictionary<int, int> Offsets, int First)> _objectStreams = [];

    private abstract record Location;

    private sealed record InFile(long Offset, int Generation) : Location;

    private sealed record Compressed(int Stream, int Index) : Location;

    public PdfDictionary Trailer { get; }

    /// <summary>Where the newest cross-reference section starts: the update's <c>/Prev</c>.</summary>
    public long LastXref { get; }

    /// <summary>The newest section is a stream: the update writes one too.</summary>
    public bool UsesXrefStream { get; }

    public int Size => (int)((Trailer["Size"] as PdfNumber)?.Integer ?? 0);

    /// <summary>
    /// How many incremental updates were appended. Each one leaves the previous revision in the
    /// file, readable by anyone who looks; a linearized PDF starts with two sections of its own.
    /// </summary>
    public int Updates { get; }

    private const int LinearizationWindow = 1024;

    public PdfFile(byte[] data)
    {
        _data = data;
        if (PdfParser.IndexOf(data, "%PDF-"u8, 0) is < 0 or > 1024)
        {
            throw new PdfException("Este arquivo não é um PDF.");
        }

        LastXref = FindStartXref();
        var trailer = ReadSection(LastXref, out var isStream);
        Trailer = trailer;
        UsesXrefStream = isStream;

        var visited = new HashSet<long> { LastXref };
        var previous = trailer["Prev"];
        while (previous is PdfNumber { Integer: { } offset } && visited.Add(offset) && visited.Count < MaxSections)
        {
            previous = ReadSection(offset, out _)["Prev"];
        }

        var linearized = data.AsSpan(0, Math.Min(data.Length, LinearizationWindow)).IndexOf("/Linearized"u8) >= 0;
        var original = linearized ? 2 : 1;
        Updates = Math.Max(0, visited.Count - original);
    }

    public string Version()
    {
        var start = PdfParser.IndexOf(_data, "%PDF-"u8, 0) + 5;
        var end = start;
        while (end < _data.Length && end - start < 8 && (char.IsAsciiDigit((char)_data[end]) || _data[end] == '.')) end++;
        var header = Encoding.ASCII.GetString(_data, start, end - start);

        // The catalog may announce a newer version than the header: an update cannot rewrite the header.
        var declared = (Resolve(Resolve(Trailer["Root"]) is PdfDictionary root ? root["Version"] : null) as PdfName)?.Value;
        return declared is not null && string.CompareOrdinal(declared, header) > 0 ? declared : header;
    }

    /// <summary>Follows references; anything else comes back as it is. A missing object is null.</summary>
    public object? Resolve(object? value)
    {
        for (var hops = 0; value is PdfRef reference && hops < 32; hops++)
        {
            value = Load(reference.Number);
        }

        return value is PdfNull ? null : value;
    }

    public int GenerationOf(int number) => _locations.GetValueOrDefault(number) is InFile file ? file.Generation : 0;

    private object? Load(int number) => _locations.GetValueOrDefault(number) switch
    {
        InFile file => ReadIndirect(file.Offset, number),
        Compressed compressed => ReadCompressed(compressed),
        _ => null,
    };

    private object ReadIndirect(long offset, int expected)
    {
        if (offset < 0 || offset >= _data.Length) throw PdfParser.Broken();
        var parser = new PdfParser(_data, (int)offset);
        if (parser.ReadInteger() != expected) throw PdfParser.Broken();
        parser.ReadInteger();
        parser.Expect("obj");
        return parser.ReadIndirectBody(length => (Resolve(length) as PdfNumber)?.Integer);
    }

    private object? ReadCompressed(Compressed location)
    {
        if (!_objectStreams.TryGetValue(location.Stream, out var stream))
        {
            stream = LoadObjectStream(location.Stream);
            _objectStreams[location.Stream] = stream;
        }

        if (!stream.Offsets.TryGetValue(location.Index, out var offset)) return null;
        return new PdfParser(stream.Data, stream.First + offset).ReadObject();
    }

    /// <summary>The header lists, for each object, its number and where it starts after <c>/First</c>.</summary>
    private (byte[] Data, Dictionary<int, int> Offsets, int First) LoadObjectStream(int number)
    {
        if (_locations.GetValueOrDefault(number) is not InFile file
            || ReadIndirect(file.Offset, number) is not PdfStream stream)
        {
            throw PdfParser.Broken();
        }

        var data = PdfFilters.Decode(stream, Resolve);
        var count = (Resolve(stream.Dictionary["N"]) as PdfNumber)?.Integer ?? 0;
        var first = (int)((Resolve(stream.Dictionary["First"]) as PdfNumber)?.Integer ?? 0);
        var header = new PdfParser(data, 0);
        var offsets = new Dictionary<int, int>();
        for (var index = 0; index < count; index++)
        {
            header.ReadInteger();
            offsets[index] = (int)header.ReadInteger();
        }

        return (data, offsets, first);
    }

    private long FindStartXref()
    {
        var from = Math.Max(0, _data.Length - 2048);
        var found = -1;
        for (var at = PdfParser.IndexOf(_data, "startxref"u8, from); at >= 0; at = PdfParser.IndexOf(_data, "startxref"u8, at + 1))
        {
            found = at;
        }

        if (found < 0) throw PdfParser.Broken();
        var parser = new PdfParser(_data, found + "startxref".Length);
        return parser.ReadInteger();
    }

    /// <summary>Records the section's objects not seen in a newer one, and returns its trailer.</summary>
    private PdfDictionary ReadSection(long offset, out bool isStream)
    {
        if (offset < 0 || offset >= _data.Length) throw PdfParser.Broken();
        var parser = new PdfParser(_data, (int)offset);
        isStream = parser.PeekKeyword() != "xref";
        if (isStream) return ReadXrefStream(parser);

        parser.Expect("xref");
        ReadXrefTable(parser);
        parser.Expect("trailer");
        if (parser.ReadObject() is not PdfDictionary trailer) throw PdfParser.Broken();

        // A hybrid file lists its compressed objects in a stream the classic table points to.
        if (trailer["XRefStm"] is PdfNumber { Integer: { } hidden } && hidden != offset)
        {
            ReadXrefStream(new PdfParser(_data, (int)hidden));
        }

        return trailer;
    }

    private void ReadXrefTable(PdfParser parser)
    {
        while (parser.PeekKeyword() is { } word && word != "trailer")
        {
            var first = (int)parser.ReadInteger();
            var count = parser.ReadInteger();
            for (var index = 0; index < count; index++)
            {
                var offset = parser.ReadInteger();
                var generation = (int)parser.ReadInteger();
                var kind = parser.PeekKeyword();
                if (kind is not ("n" or "f")) throw PdfParser.Broken();
                parser.Expect(kind);
                if (kind == "n") _locations.TryAdd(first + index, new InFile(offset, generation));
            }
        }
    }

    private PdfDictionary ReadXrefStream(PdfParser parser)
    {
        parser.ReadInteger();
        parser.ReadInteger();
        parser.Expect("obj");
        if (parser.ReadIndirectBody(length => (length as PdfNumber)?.Integer) is not PdfStream stream)
        {
            throw PdfParser.Broken();
        }

        var widths = (stream.Dictionary["W"] as PdfArray)?.Select(width => (int)((width as PdfNumber)?.Integer ?? -1)).ToArray();
        if (widths is not { Length: 3 } || widths.Any(width => width is < 0 or > 8)) throw PdfParser.Broken();

        var size = (stream.Dictionary["Size"] as PdfNumber)?.Integer ?? 0;
        var index = stream.Dictionary["Index"] as PdfArray ?? [PdfNumber.Of(0), PdfNumber.Of(size)];
        ReadXrefEntries(PdfFilters.Decode(stream, value => value), widths, index);
        return stream.Dictionary;
    }

    private void ReadXrefEntries(byte[] data, int[] widths, PdfArray index)
    {
        var entry = widths.Sum();
        var at = 0;
        for (var pair = 0; pair + 1 < index.Count; pair += 2)
        {
            var first = (index[pair] as PdfNumber)?.Integer ?? 0;
            var count = (index[pair + 1] as PdfNumber)?.Integer ?? 0;
            for (var item = 0; item < count && at + entry <= data.Length; item++, at += entry)
            {
                // Without the first field every entry is an object in the file.
                var type = widths[0] == 0 ? 1 : Field(data, at, widths[0]);
                var second = Field(data, at + widths[0], widths[1]);
                var third = Field(data, at + widths[0] + widths[1], widths[2]);
                var number = (int)(first + item);
                if (type == 1) _locations.TryAdd(number, new InFile(second, (int)third));
                else if (type == 2) _locations.TryAdd(number, new Compressed((int)second, (int)third));
            }
        }
    }

    private static long Field(byte[] data, int at, int width)
    {
        long value = 0;
        for (var index = 0; index < width; index++) value = (value << 8) | data[at + index];
        return value;
    }
}
