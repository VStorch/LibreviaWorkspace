using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// A child of <c>office:text</c> and the nodes it gave. Each node carries an <c>oid</c>: the first
/// <c>o7</c>, the next ones <c>o7.2</c>, <c>o7.3</c>… (a page break before a paragraph, a section's
/// blocks). The block is kept only when all of them come back in order and unchanged.
/// </summary>
/// <param name="Number">From 1, in document order.</param>
/// <param name="ElementIndex">Its place among the elements of <c>office:text</c>.</param>
/// <param name="Notices">What it holds and the model does not: lost if it is rewritten.</param>
internal sealed record OdtSourceBlock(int Number, int ElementIndex, List<Node> Nodes, Inventory Notices)
{
    private const string Prefix = "o";

    public string OidAt(int part) => part == 0 ? $"{Prefix}{Number}" : $"{Prefix}{Number}.{part + 1}";

    public void Stamp()
    {
        for (var part = 0; part < Nodes.Count; part++) Nodes[part].With("oid", OidAt(part));
    }
}
