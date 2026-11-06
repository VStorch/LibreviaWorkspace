using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>What the readers of one <c>content.xml</c> share.</summary>
internal sealed class OdtReadContext(OdtPackage package, OdtStyles styles, Inventory inventory)
{
    public OdtPackage Package => package;
    public OdtStyles Styles => styles;

    private OdtCommentReader? _comments;

    public OdtCommentReader Comments => _comments ??= new(package.Content);

    private OdtChangeReader? _changes;

    public OdtChangeReader Changes => _changes ??= new(OdtTrackedChanges.Read(package.Content));

    /// <summary>Inside a note's body, where ODF allows no other note.</summary>
    public bool InNote { get; set; }

    /// <summary>The document's, or the block's being read: see <see cref="BeginBlock"/>.</summary>
    private readonly Inventory _document = inventory;

    public Inventory Inventory { get; private set; } = inventory;

    private bool _sawBlock;

    /// <summary>The <c>text:section</c> the model carries as a section with columns, see <see cref="OdtSectionReader"/>.</summary>
    public XElement? ModeledSection { get; set; }

    /// <summary>The first block in the text: a page break before it would only add a blank sheet.</summary>
    public bool IsFirstBlock => !_sawBlock;

    public void NoteLoss(string key, string? value = null) => Inventory.NoteLoss(key, value);

    /// <summary>
    /// What one body block holds and the model does not: lost only if that block is edited, so the
    /// writer needs it block by block.
    /// </summary>
    public void BeginBlock() => Inventory = new Inventory();

    public Inventory EndBlock()
    {
        var block = Inventory;
        Inventory = _document;
        foreach (var notice in block.Lost) _document.NoteLoss(notice.Key, notice.Value);
        return block;
    }

    public void SeeBlock() => _sawBlock = true;
}
