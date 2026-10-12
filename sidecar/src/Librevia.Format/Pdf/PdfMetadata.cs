using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Serialization;
using Librevia.Format.Docx;

namespace Librevia.Format.Pdf;

/// <summary>What the properties dialog shows. <c>Application</c> and <c>Producer</c> are read only.</summary>
public sealed record PdfInspection(
    [property: JsonPropertyName("version")] string Version,
    [property: JsonPropertyName("encrypted")] bool Encrypted,
    [property: JsonPropertyName("signed")] bool Signed,
    [property: JsonPropertyName("properties")] PropertiesDto? Properties,
    [property: JsonPropertyName("application")] string? Application,
    [property: JsonPropertyName("producer")] string? Producer);

/// <summary>
/// The information dictionary and the XMP packet, changed by an <b>incremental update</b>: the
/// original bytes stay as they are, and the new objects, a cross-reference section and a trailer
/// pointing back to the old one go after them. That is also what keeps an earlier signature valid
/// for the revision it signed.
/// </summary>
public static class PdfMetadata
{
    public const string EncryptedRefusal =
        "Este PDF é criptografado. O Librevia não altera as propriedades de PDFs criptografados.";

    private static readonly string[] DateKeys = ["CreationDate", "ModDate"];

    public static PdfInspection Inspect(byte[] bytes)
    {
        var file = new PdfFile(bytes);
        if (file.Trailer["Encrypt"] is not null)
        {
            return new PdfInspection(file.Version(), Encrypted: true, Signed: false, null, null, null);
        }

        var values = Values(file);
        string? Value(string key) => values.GetValueOrDefault(key);
        var properties = new PropertiesDto(
            Title: Value("Title"),
            Subject: Value("Subject"),
            Creator: Value("Author"),
            Keywords: Value("Keywords"),
            Created: Value("CreationDate"),
            Modified: Value("ModDate"));

        return new PdfInspection(
            file.Version(),
            Encrypted: false,
            Signed: IsSigned(file),
            properties == new PropertiesDto() ? null : properties,
            Value("Creator"),
            Value("Producer"));
    }

    /// <summary>The bytes to append; empty when nothing changes.</summary>
    public static byte[] Update(byte[] bytes, PropertiesDto patch)
    {
        var file = new PdfFile(bytes);
        if (file.Trailer["Encrypt"] is not null) throw new PdfException(EncryptedRefusal);

        var changes = ChangesOf(patch);
        var update = new IncrementalUpdate(file, bytes.Length, NeedsLineBreak(bytes));
        var infoChanged = UpdateInfo(file, update, changes);
        var xmpChanged = UpdateXmp(file, update, changes);
        return infoChanged || xmpChanged ? update.Finish() : [];
    }

    private static bool NeedsLineBreak(byte[] bytes) => bytes.Length > 0 && bytes[^1] is not (byte)'\n' and not (byte)'\r';

    /// <summary>By the information dictionary's key; dates as W3CDTF. Empty clears the field.</summary>
    private static Dictionary<string, string> ChangesOf(PropertiesDto patch)
    {
        var changes = new Dictionary<string, string>();
        void Add(string key, string? value)
        {
            if (value is not null) changes[key] = value;
        }

        Add("Title", patch.Title);
        Add("Author", patch.Creator);
        Add("Subject", patch.Subject);
        Add("Keywords", patch.Keywords);
        Add("ModDate", patch.Modified);
        return changes;
    }

    /// <summary>The dictionary first, the XMP for what it lacks (PDF 2.0 may have only XMP).</summary>
    private static Dictionary<string, string> Values(PdfFile file)
    {
        var values = new Dictionary<string, string>();
        if (file.Resolve(file.Trailer["Info"]) is PdfDictionary info)
        {
            foreach (var (key, value) in info.Entries)
            {
                if (file.Resolve(value) is not PdfString text) continue;
                var decoded = PdfText.Decode(text).Trim();
                if (DateKeys.Contains(key)) decoded = PdfText.ToW3cdtf(decoded) ?? "";
                if (decoded.Length > 0) values[key] = decoded;
            }
        }

        if (XmpOf(file) is { } packet)
        {
            foreach (var (key, value) in Xmp.Read(packet.Document))
            {
                if (!DateKeys.Contains(key) || PdfText.IsW3cdtf(value)) values.TryAdd(key, value);
            }
        }

        return values;
    }

    private static bool UpdateInfo(PdfFile file, IncrementalUpdate update, Dictionary<string, string> changes)
    {
        var reference = file.Trailer["Info"] as PdfRef;
        var current = file.Resolve(file.Trailer["Info"]) as PdfDictionary;
        var info = current?.Copy() ?? new PdfDictionary();
        var changed = false;
        foreach (var (key, value) in changes)
        {
            var wanted = DateKeys.Contains(key) ? PdfText.ToPdfDate(value) ?? "" : value;
            var existing = file.Resolve(info[key]) is PdfString text ? PdfText.Decode(text) : null;
            if (wanted.Length == 0 ? existing is null : existing == wanted) continue;

            info[key] = wanted.Length == 0 ? null : PdfText.Encode(wanted);
            changed = true;
        }

        if (!changed) return false;
        var target = reference ?? update.NewReference();
        update.Add(target, Encoding.Latin1.GetBytes(PdfWriter.Write(info)));
        update.Trailer["Info"] = target;
        return true;
    }

    /// <summary>
    /// Only when the file has a packet: one is not created, since the dictionary alone is what
    /// readers fall back to.
    /// </summary>
    private static bool UpdateXmp(PdfFile file, IncrementalUpdate update, Dictionary<string, string> changes)
    {
        if (XmpOf(file) is not { } packet) return false;

        var before = Xmp.Write(packet.Document);
        Xmp.Update(packet.Document, changes);
        var data = Xmp.Write(packet.Document);
        if (before.AsSpan().SequenceEqual(data)) return false;

        var dictionary = packet.Stream.Dictionary.Copy();
        foreach (var key in new[] { "Filter", "DecodeParms", "DL" }) dictionary[key] = null;
        dictionary["Length"] = PdfNumber.Of(data.Length);

        var body = new MemoryStream();
        body.Write(Encoding.Latin1.GetBytes(PdfWriter.Write(dictionary) + "\nstream\n"));
        body.Write(data);
        body.Write("\nendstream"u8);
        update.Add(packet.Reference, body.ToArray());
        return true;
    }

    private static (PdfRef Reference, PdfStream Stream, System.Xml.Linq.XDocument Document)? XmpOf(PdfFile file)
    {
        if (file.Resolve(file.Trailer["Root"]) is not PdfDictionary catalog) return null;
        if (catalog["Metadata"] is not PdfRef reference || file.Resolve(reference) is not PdfStream stream) return null;

        var document = Xmp.Parse(PdfFilters.Decode(stream, file.Resolve));
        if (document is null) throw new PdfException("Os metadados XMP deste PDF estão danificados.");
        return (reference, stream, document);
    }

    /// <summary>A signature field marked in the form, or a certification (<c>/Perms</c>).</summary>
    private static bool IsSigned(PdfFile file)
    {
        if (file.Resolve(file.Trailer["Root"]) is not PdfDictionary catalog) return false;
        if (catalog["Perms"] is not null) return true;
        var form = file.Resolve(catalog["AcroForm"]) as PdfDictionary;
        var flags = (file.Resolve(form?["SigFlags"]) as PdfNumber)?.Integer ?? 0;
        return (flags & 1) != 0;
    }

    /// <summary>The objects after the original bytes, then their cross-reference section.</summary>
    private sealed class IncrementalUpdate(PdfFile file, long originalLength, bool lineBreak)
    {
        private readonly MemoryStream _output = new();
        private readonly SortedDictionary<int, (long Offset, int Generation)> _offsets = [];
        private int _nextNumber = file.Size;

        public PdfDictionary Trailer { get; } = TrailerOf(file);

        private long Offset => originalLength + _output.Length;

        public PdfRef NewReference() => new(_nextNumber++, 0);

        public void Add(PdfRef reference, byte[] body)
        {
            if (lineBreak && _output.Length == 0) _output.WriteByte((byte)'\n');
            _offsets[reference.Number] = (Offset, reference.Generation);
            _output.Write(Encoding.Latin1.GetBytes($"{reference.Number} {reference.Generation} obj\n"));
            _output.Write(body);
            _output.Write("\nendobj\n"u8);
        }

        public byte[] Finish()
        {
            Trailer["Prev"] = PdfNumber.Of(file.LastXref);
            Trailer["ID"] = NewId(Trailer["ID"]);
            if (file.UsesXrefStream) WriteXrefStream();
            else WriteXrefTable();
            return _output.ToArray();
        }

        private void WriteXrefTable()
        {
            Trailer["Size"] = PdfNumber.Of(Math.Max(_nextNumber, _offsets.Keys.Max() + 1));
            var start = Offset;
            var table = new StringBuilder("xref\n");
            foreach (var (number, (offset, generation)) in _offsets)
            {
                // Each entry is exactly 20 bytes, its line break included.
                table.Append(CultureInfo.InvariantCulture, $"{number} 1\n{offset:D10} {generation:D5} n\r\n");
            }

            table.Append("trailer\n").Append(PdfWriter.Write(Trailer));
            table.Append(CultureInfo.InvariantCulture, $"\nstartxref\n{start}\n%%EOF\n");
            _output.Write(Encoding.Latin1.GetBytes(table.ToString()));
        }

        /// <summary>Uncompressed: three fields, 1, 4 and 2 bytes wide.</summary>
        private void WriteXrefStream()
        {
            var self = NewReference();
            var start = Offset;
            _offsets[self.Number] = (start, 0);
            Trailer["Size"] = PdfNumber.Of(Math.Max(_nextNumber, _offsets.Keys.Max() + 1));

            var entries = new MemoryStream();
            var index = new PdfArray();
            foreach (var (number, (offset, generation)) in _offsets)
            {
                index.Add(PdfNumber.Of(number));
                index.Add(PdfNumber.Of(1));
                entries.WriteByte(1);
                entries.Write([(byte)(offset >> 24), (byte)(offset >> 16), (byte)(offset >> 8), (byte)offset]);
                entries.Write([(byte)(generation >> 8), (byte)generation]);
            }

            var dictionary = Trailer.Copy();
            dictionary["Type"] = new PdfName("XRef");
            dictionary["W"] = new PdfArray { PdfNumber.Of(1), PdfNumber.Of(4), PdfNumber.Of(2) };
            dictionary["Index"] = index;
            dictionary["Length"] = PdfNumber.Of(entries.Length);

            _output.Write(Encoding.Latin1.GetBytes($"{self.Number} 0 obj\n{PdfWriter.Write(dictionary)}\nstream\n"));
            entries.WriteTo(_output);
            _output.Write(Encoding.Latin1.GetBytes($"\nendstream\nendobj\nstartxref\n{start}\n%%EOF\n"));
        }

        /// <summary>The newest trailer, without what only described its own section.</summary>
        private static PdfDictionary TrailerOf(PdfFile file)
        {
            var trailer = file.Trailer.Copy();
            foreach (var key in new[] { "Prev", "XRefStm", "Type", "W", "Index", "Length", "Filter", "DecodeParms", "DL" })
            {
                trailer[key] = null;
            }

            return trailer;
        }

        /// <summary>The first half names the document and stays; the second names this version.</summary>
        private static PdfArray? NewId(object? current) =>
            current is PdfArray { Count: 2 } id && id[0] is PdfString permanent
                ? [permanent, new PdfString(RandomNumberGenerator.GetBytes(16))]
                : current as PdfArray;
    }
}
