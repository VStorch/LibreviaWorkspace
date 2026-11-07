using System.IO.Compression;
using System.Text;

namespace Librevia.Format.Tests;

/// <summary>
/// ODT packages built in code from small, readable XML parts. <see cref="LibreOffice"/> mirrors
/// what LibreOffice writes; the rest isolates one construct each.
/// </summary>
public static class OdtFixtures
{
    public const string MimeType = "application/vnd.oasis.opendocument.text";

    private const string Namespaces =
        """xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" """ +
        """xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" """ +
        """xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" """ +
        """xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" """ +
        """xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" """ +
        """xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" """ +
        """xmlns:xlink="http://www.w3.org/1999/xlink" """ +
        """xmlns:dc="http://purl.org/dc/elements/1.1/" """ +
        """xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" """ +
        """xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" """ +
        """xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" """ +
        """xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" """ +
        """office:version="1.3" """;

    /// <summary>A4 portrait, 2 cm margins, no header: LibreOffice's default page.</summary>
    public const string DefaultPage =
        """<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/><style:header-style/><style:footer-style/></style:page-layout>""";

    public const string DefaultMaster = """<style:master-page style:name="Standard" style:page-layout-name="pm1"/>""";

    /// <summary>The named styles LibreOffice writes for a new text document, trimmed.</summary>
    public const string LibreOfficeStyles =
        """<style:default-style style:family="paragraph"><style:paragraph-properties style:tab-stop-distance="1.251cm"/><style:text-properties style:font-name="Liberation Serif" fo:font-size="12pt"/></style:default-style>""" +
        """<style:style style:name="Standard" style:family="paragraph" style:class="text"/>""" +
        """<style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:class="text"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.247cm" fo:line-height="115%"/></style:style>""" +
        """<style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text"><style:paragraph-properties fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:keep-with-next="always"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="14pt"/></style:style>""" +
        """<style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="1" style:class="text"><style:text-properties fo:font-size="130%" fo:font-weight="bold"/></style:style>""" +
        """<style:style style:name="Strong_20_Emphasis" style:display-name="Strong Emphasis" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>""";

    public const string LibreOfficeFonts =
        """<office:font-face-decls><style:font-face style:name="Liberation Serif" svg:font-family="&apos;Liberation Serif&apos;" style:font-family-generic="roman"/><style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" style:font-family-generic="swiss"/></office:font-face-decls>""";

    /// <summary>
    /// A heading, a body paragraph through the automatic style P1 (over "Text body"), T1 in bold, a
    /// character style, collapsed and real spaces, a tab and a line break, and a soft page break.
    /// </summary>
    public static byte[] LibreOffice() => Build(
        automatic:
            """<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Text_20_body"><style:paragraph-properties fo:text-align="center"/></style:style>""" +
            """<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>""",
        body:
            """<text:h text:style-name="Heading_20_1" text:outline-level="1">Título</text:h>""" +
            """<text:p text:style-name="P1">Olá  <text:span text:style-name="T1">mundo</text:span> e<text:s text:c="2"/>mais<text:tab/>fim<text:line-break/>linha <text:span text:style-name="Strong_20_Emphasis">forte</text:span></text:p>""" +
            """<text:soft-page-break/>""" +
            """<text:p text:style-name="Text_20_body">Corpo</text:p>""",
        styles: LibreOfficeStyles,
        fonts: LibreOfficeFonts);

    /// <summary>
    /// What LibreOffice puts around the text: declarations before the first block, a list, a table
    /// and a paragraph with a footnote, and the entries a save must not touch.
    /// </summary>
    public static byte[] Rich() => Build(
        automatic:
            """<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:text-align="center"/></style:style>""" +
            """<text:list-style style:name="L1"><text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.27cm" fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-bullet></text:list-style>""",
        body:
            """<text:sequence-decls><text:sequence-decl text:display-outline-level="0" text:name="Illustration"/></text:sequence-decls>""" +
            """<text:p text:style-name="P1">Primeiro</text:p>""" +
            """<text:list text:style-name="L1"><text:list-item><text:p>um</text:p></text:list-item><text:list-item><text:p>dois</text:p></text:list-item></text:list>""" +
            """<table:table table:name="Tabela1"><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>b</text:p></table:table-cell></table:table-row></table:table>""" +
            """<text:p>Com nota<text:note text:id="ftn1" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>nota</text:p></text:note-body></text:note> aqui</text:p>""" +
            """<text:p>Último</text:p>""" +
            """<text:soft-page-break/>""",
        meta: """<dc:title>Relatório</dc:title><meta:initial-creator>Ana</meta:initial-creator>""",
        files: new Dictionary<string, byte[]>
        {
            ["settings.xml"] = Utf8("""<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"""),
            ["Thumbnails/thumbnail.png"] = Png(),
            ["Configurations2/accelerator/current.xml"] = Utf8("<x/>"),
            ["Object 1/content.xml"] = Utf8("""<?xml version="1.0" encoding="UTF-8"?><math xmlns="http://www.w3.org/1998/Math/MathML"/>"""),
        });

    /// <summary>The same document with prefixes of its own choosing: <c>o:</c>, <c>s:</c>, <c>t:</c>.</summary>
    public static byte[] OwnPrefixes()
    {
        const string spaces =
            """xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" """ +
            """xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" o:version="1.3" """;
        return Zip(
        [
            ("mimetype", Encoding.ASCII.GetBytes(MimeType)),
            ("content.xml", Utf8($"""<?xml version="1.0" encoding="UTF-8"?><o:document-content {spaces}><o:body><o:text><t:p>Um</t:p><t:p>Dois</t:p></o:text></o:body></o:document-content>""")),
            ("styles.xml", Utf8($"""<?xml version="1.0" encoding="UTF-8"?><o:document-styles {spaces}><o:styles><s:style s:name="Standard" s:family="paragraph"/></o:styles></o:document-styles>""")),
            ("META-INF/manifest.xml", Utf8(
                """<?xml version="1.0" encoding="UTF-8"?><m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" m:version="1.3">""" +
                $"""<m:file-entry m:full-path="/" m:media-type="{MimeType}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/>""" +
                """<m:file-entry m:full-path="styles.xml" m:media-type="text/xml"/></m:manifest>""")),
        ]);
    }

    /// <summary>A whole package: the parts that are not given are LibreOffice's defaults.</summary>
    public static byte[] Build(
        string body,
        string automatic = "",
        string styles = LibreOfficeStyles,
        string fonts = LibreOfficeFonts,
        string pageLayouts = DefaultPage,
        string masters = DefaultMaster,
        string? meta = null,
        IReadOnlyDictionary<string, byte[]>? files = null,
        string manifestExtra = "",
        string? mimeType = MimeType)
    {
        var parts = new List<(string Name, byte[] Data)>();
        if (mimeType is not null) parts.Add(("mimetype", Encoding.ASCII.GetBytes(mimeType)));
        parts.Add(("content.xml", Utf8(
            $"""<?xml version="1.0" encoding="UTF-8"?><office:document-content {Namespaces}>{fonts}<office:automatic-styles>{automatic}</office:automatic-styles><office:body><office:text>{body}</office:text></office:body></office:document-content>""")));
        parts.Add(("styles.xml", Utf8(
            $"""<?xml version="1.0" encoding="UTF-8"?><office:document-styles {Namespaces}>{fonts}<office:styles>{styles}</office:styles><office:automatic-styles>{pageLayouts}</office:automatic-styles><office:master-styles>{masters}</office:master-styles></office:document-styles>""")));
        if (meta is not null)
        {
            parts.Add(("meta.xml", Utf8(
                $"""<?xml version="1.0" encoding="UTF-8"?><office:document-meta {Namespaces}><office:meta>{meta}</office:meta></office:document-meta>""")));
        }

        foreach (var (name, data) in files ?? new Dictionary<string, byte[]>()) parts.Add((name, data));
        parts.Add(("META-INF/manifest.xml", Utf8(
            """<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">""" +
            $"""<manifest:file-entry manifest:full-path="/" manifest:media-type="{mimeType ?? MimeType}"/>""" +
            manifestExtra + ManifestEntries(parts, manifestExtra) + "</manifest:manifest>")));
        return Zip(parts);
    }

    /// <summary>Every part the extra entries do not declare, as LibreOffice asks of a valid package.</summary>
    private static string ManifestEntries(List<(string Name, byte[] Data)> parts, string declared)
    {
        var paths = parts.Select(part => part.Name).Where(name => name != "mimetype").ToList();
        var objects = paths.Where(path => path.StartsWith("Object ", StringComparison.Ordinal) && path.Contains('/'))
            .Select(path => path[..(path.IndexOf('/') + 1)]);
        var entries = objects.Distinct().Select(folder => (folder, "application/vnd.oasis.opendocument.formula"))
            .Concat(paths.Select(path => (path, MediaTypeOf(path))));
        return string.Concat(entries
            .Where(entry => !declared.Contains($"full-path=\"{entry.Item1}\"", StringComparison.Ordinal))
            .Select(entry => $"""<manifest:file-entry manifest:full-path="{entry.Item1}" manifest:media-type="{entry.Item2}"/>"""));
    }

    private static string MediaTypeOf(string path) => Path.GetExtension(path) switch
    {
        ".xml" => "text/xml",
        ".png" => "image/png",
        ".svg" => "image/svg+xml",
        _ => "application/octet-stream",
    };

    public static byte[] Zip(IEnumerable<(string Name, byte[] Data)> parts)
    {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, data) in parts)
            {
                var level = name == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Optimal;
                using var stream = archive.CreateEntry(name, level).Open();
                stream.Write(data);
            }
        }

        return buffer.ToArray();
    }

    /// <summary>A 2 × 3 pixel PNG, header only: enough to measure.</summary>
    public static byte[] Png()
    {
        var bytes = new byte[33];
        byte[] signature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52];
        signature.CopyTo(bytes, 0);
        bytes[19] = 2;
        bytes[23] = 3;
        return bytes;
    }

    public static byte[] Utf8(string text) => Encoding.UTF8.GetBytes(text);
}
