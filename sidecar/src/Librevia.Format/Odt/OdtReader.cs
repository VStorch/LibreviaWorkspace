using System.IO.Compression;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>odt.open</c>: the same <see cref="OpenResult"/> as <c>docx.open</c>, with an <c>oid</c> on
/// each body block (<see cref="OdtSourceBlock"/>). What the model does not carry stays in the file
/// while its block is not edited: it is <b>invisible</b>, and <see cref="OdtWriter"/> declares it
/// lost only when it rewrites the block.
/// </summary>
public static class OdtReader
{
    /// <summary>The ceiling of one decompressed part; see <see cref="OdtPackage"/>.</summary>
    public const long MaxEntryBytes = OdtPackage.MaxEntryBytes;

    public static OpenResult Read(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = OpenArchive(stream);
        var source = OdtSource.Read(OdtPackage.Open(archive));

        var inventory = new Inventory();
        foreach (var notice in source.Notices.Lost) inventory.NoteInvisible(notice.Key, notice.Value);
        return new OpenResult(source.Model(), inventory);
    }

    /// <summary>An untrusted file: whatever the zip reader throws becomes a sentence.</summary>
    internal static ZipArchive OpenArchive(Stream stream)
    {
        try
        {
            return new ZipArchive(stream, ZipArchiveMode.Read);
        }
        catch (Exception problem) when (problem is InvalidDataException or IOException or NotSupportedException or ArgumentException)
        {
            throw new OdtException("sidecar.error.odtUnreadable", problem);
        }
    }
}

/// <summary>An ODT read: the package, its blocks and the model they give.</summary>
internal sealed class OdtSource
{
    public required OdtPackage Package { get; init; }
    public required XElement Text { get; init; }
    public required OdtStyles Styles { get; init; }
    public required List<OdtSourceBlock> Blocks { get; init; }
    public required PageSetupDto Page { get; init; }
    public required StyleSheetDto Sheet { get; init; }
    public PropertiesDto? Properties { get; init; }
    public required Inventory Notices { get; init; }

    /// <summary>The master page of the first block, whose page setup the model carries.</summary>
    public string? MasterPage { get; init; }

    /// <summary>The elements of <c>office:text</c>, in order: what each block's index points to.</summary>
    public required List<XElement> Elements { get; init; }

    public static OdtSource Read(OdtPackage package)
    {
        var text = package.Content.Root?.Element(Ns.Office + "body")?.Element(Ns.Office + "text")
                   ?? throw new OdtException("sidecar.error.odtWithoutContent");

        var notices = new Inventory();
        var styles = OdtStyles.Read(package.Content, package.Styles);
        var context = new OdtReadContext(package, styles, notices);
        var blocks = new OdtBodyReader(context).ReadBlocks(text);
        foreach (var block in blocks) block.Stamp();

        return new OdtSource
        {
            Package = package,
            Text = text,
            Styles = styles,
            Blocks = blocks,
            Page = OdtPageReader.Read(package.Styles, context.MasterPage, notices),
            Sheet = OdtStyleSheet.Read(styles),
            Properties = OdtProperties.Read(package.Meta, OdtPrivacyMark.IsSet(package.ReadXml(OdtPrivacyMark.SettingsPath))),
            Notices = notices,
            MasterPage = context.MasterPage,
            Elements = [.. text.Elements()],
        };
    }

    public DocumentModelDto Model()
    {
        var doc = Node.Of("doc");
        doc.Content = [.. Blocks.SelectMany(block => block.Nodes)];
        if (doc.Content.Count == 0) doc.Content.Add(OdtParagraphReader.Empty());
        return new DocumentModelDto(Page, doc, Sheet, Properties: Properties);
    }
}
