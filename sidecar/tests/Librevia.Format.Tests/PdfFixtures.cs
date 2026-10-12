using System.IO.Compression;
using System.Text;

namespace Librevia.Format.Tests;

/// <summary>Small PDFs built in code, with the structures metadata editing must read.</summary>
public static class PdfFixtures
{
    public const string Xmp =
        """<?xpacket begin="﻿" id="W5M0MpCehiHzreSzNTczkc9d"?>""" +
        """<x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">""" +
        """<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:pdf="http://ns.adobe.com/pdf/1.3/" """ +
        """xmlns:xmp="http://ns.adobe.com/xap/1.0/" xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/" """ +
        """xmp:ModifyDate="2025-02-03T04:05:06Z" pdfaid:part="2" pdfaid:conformance="B">""" +
        """<dc:title><rdf:Alt><rdf:li xml:lang="x-default">Relatório anual</rdf:li></rdf:Alt></dc:title>""" +
        """<dc:creator><rdf:Seq><rdf:li>Ana Lima</rdf:li></rdf:Seq></dc:creator>""" +
        """<pdf:Producer>Gerador 1.0</pdf:Producer>""" +
        """</rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end="w"?>""";

    /// <summary>The information dictionary, as a writer of the 2000s would leave it.</summary>
    public const string Info =
        "<< /Title (Relat\\363rio anual) /Author <FEFF0041006E00610020004C0069006D0061> /Subject (Contas) " +
        "/Producer (Gerador 1.0) /Creator (Editor) /CreationDate (D:20250102030405-03'00') " +
        "/ModDate (D:20250203040506Z) /Custom (fica) >>";

    /// <summary>A classic cross-reference table. Each body is one object, numbered from 1.</summary>
    public static byte[] Classic(string? info = Info, string? xmp = null, string catalogExtras = "", string trailerExtras = "")
    {
        var objects = new List<byte[]>
        {
            Ascii($"<< /Type /Catalog /Pages 2 0 R{(xmp is null ? "" : " /Metadata 5 0 R")}{catalogExtras} >>"),
            Ascii("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"),
            Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>"),
            Ascii(info ?? "null"),
        };
        if (xmp is not null) objects.Add(Stream("<< /Type /Metadata /Subtype /XML", Encoding.UTF8.GetBytes(xmp)));

        var output = new MemoryStream();
        Write(output, "%PDF-1.4\n%âãÏÓ\n");
        var offsets = objects.Select((body, index) => WriteObject(output, index + 1, body)).ToList();

        var xref = output.Length;
        Write(output, $"xref\n0 {objects.Count + 1}\n0000000000 65535 f\r\n");
        foreach (var offset in offsets) Write(output, $"{offset:D10} 00000 n\r\n");
        var infoEntry = info is null ? "" : " /Info 4 0 R";
        Write(output, $"trailer\n<< /Size {objects.Count + 1} /Root 1 0 R{infoEntry} /ID [<0123456789ABCDEF0123456789ABCDEF> <0123456789ABCDEF0123456789ABCDEF>]{trailerExtras} >>\n");
        Write(output, $"startxref\n{xref}\n%%EOF\n");
        return output.ToArray();
    }

    /// <summary>
    /// PDF 1.5 style: the catalog and the information dictionary inside an object stream, and the
    /// cross-reference stream compressed with the PNG "up" predictor.
    /// </summary>
    public static byte[] WithXrefStream()
    {
        var output = new MemoryStream();
        Write(output, "%PDF-1.7\n");
        var pages = WriteObject(output, 2, Ascii("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"));
        var page = WriteObject(output, 3, Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>"));

        var catalog = "<< /Type /Catalog /Pages 2 0 R >>";
        var header = $"1 0 4 {catalog.Length + 1} ";
        var packed = Deflate(Ascii(header + catalog + " " + Info));
        var objectStream = WriteObject(
            output, 5, Stream($"<< /Type /ObjStm /N 2 /First {header.Length} /Filter /FlateDecode", packed));

        var xref = output.Length;
        var rows = new (byte Type, long Second, int Third)[]
        {
            (0, 0, 65535), (2, 5, 0), (1, pages, 0), (1, page, 0), (2, 5, 1), (1, objectStream, 0), (1, xref, 0),
        };
        var data = Deflate(PredictUp(rows.SelectMany(row => Row(row.Type, row.Second, row.Third)).ToArray(), 6));
        WriteObject(output, 6, Stream(
            "<< /Type /XRef /Size 7 /W [1 4 1] /Root 1 0 R /Info 4 0 R /Filter /FlateDecode " +
            "/DecodeParms << /Predictor 12 /Columns 6 >> /ID [<AA> <AA>]", data));
        Write(output, $"startxref\n{xref}\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] Row(byte type, long second, int third) =>
        [type, (byte)(second >> 24), (byte)(second >> 16), (byte)(second >> 8), (byte)second, (byte)third];

    private static byte[] PredictUp(byte[] data, int columns)
    {
        var output = new MemoryStream();
        var previous = new byte[columns];
        for (var start = 0; start < data.Length; start += columns)
        {
            output.WriteByte(2);
            for (var index = 0; index < columns; index++) output.WriteByte((byte)(data[start + index] - previous[index]));
            previous = data[start..(start + columns)];
        }

        return output.ToArray();
    }

    private static byte[] Deflate(byte[] data)
    {
        var output = new MemoryStream();
        using (var zlib = new ZLibStream(output, CompressionLevel.Optimal, leaveOpen: true)) zlib.Write(data);
        return output.ToArray();
    }

    private static byte[] Stream(string openDictionary, byte[] data) =>
        [.. Ascii($"{openDictionary} /Length {data.Length} >>\nstream\n"), .. data, .. Ascii("\nendstream")];

    private static long WriteObject(MemoryStream output, int number, byte[] body)
    {
        var offset = output.Length;
        Write(output, $"{number} 0 obj\n");
        output.Write(body);
        Write(output, "\nendobj\n");
        return offset;
    }

    private static void Write(MemoryStream output, string text) => output.Write(Encoding.Latin1.GetBytes(text));

    private static byte[] Ascii(string text) => Encoding.Latin1.GetBytes(text);
}
