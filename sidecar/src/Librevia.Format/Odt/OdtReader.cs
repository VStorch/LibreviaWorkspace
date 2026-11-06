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
            throw new OdtException(OdtException.Unreadable, problem);
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
    public List<CommentDto>? Comments { get; init; }
    public NotesDto? Notes { get; init; }
    public required OdtTrackedChanges Changes { get; init; }

    /// <summary>The <c>office:name</c> of each named comment, by its id in <see cref="Comments"/>.</summary>
    public required IReadOnlyDictionary<string, string> CommentNames { get; init; }

    /// <summary>In order, the last the page's: what <see cref="OdtSectionPlan"/> maps the model's back to.</summary>
    public required List<OdtReadSection> Sections { get; init; }

    /// <summary>The read section whose first block this is, or <c>null</c>.</summary>
    public int? SectionStartingAt(int blockNumber)
    {
        var index = Sections.FindIndex(section => section.StartBlock == blockNumber);
        return index < 0 ? null : index;
    }

    /// <summary>The master page of the page setup, the last section's.</summary>
    public string? MasterPage => Sections[^1].Master;

    /// <summary>The elements of <c>office:text</c>, in order: what each block's index points to.</summary>
    public required List<XElement> Elements { get; init; }

    public static OdtSource Read(OdtPackage package)
    {
        var text = package.Content.Root?.Element(Ns.Office + "body")?.Element(Ns.Office + "text")
                   ?? throw new OdtException(OdtException.WithoutContent);

        var notices = new Inventory();
        var styles = OdtStyles.Read(package.Content, package.Styles);
        var context = new OdtReadContext(package, styles, notices);
        var sectionReader = new OdtSectionReader(context, notices);
        var blocks = new OdtBodyReader(context).ReadBlocks(text, sectionReader);
        foreach (var block in blocks) block.Stamp();
        var sections = sectionReader.Sections();

        return new OdtSource
        {
            Package = package,
            Text = text,
            Styles = styles,
            Blocks = blocks,
            Page = sections[^1].Setup,
            Sections = sections,
            Sheet = OdtStyleSheet.Read(styles),
            Properties = OdtProperties.Read(package.Meta, OdtPrivacyMark.IsSet(package.ReadXml(OdtPrivacyMark.SettingsPath))),
            Notices = notices,
            Comments = context.Comments.Library,
            CommentNames = context.Comments.Names,
            Notes = OdtNoteCounter.Read(package.Styles),
            Changes = context.Changes.Source,
            Elements = [.. text.Elements()],
        };
    }

    public DocumentModelDto Model()
    {
        var doc = Node.Of("doc");
        doc.Content = [.. Blocks.SelectMany(block => block.Nodes)];
        if (doc.Content.Count == 0) doc.Content.Add(OdtParagraphReader.Empty());
        var earlier = Sections.Count > 1 ? Sections[..^1].Select(section => section.Setup).ToList() : null;
        return new DocumentModelDto(
            Page, doc, Sheet, Sections: earlier, Comments: Comments, TrackChanges: Changes.Tracking ? true : null, Notes: Notes, Properties: Properties);
    }
}
