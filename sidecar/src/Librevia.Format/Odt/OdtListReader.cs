using System.Text.Json.Nodes;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>text:list</c> → <c>bulletList</c>/<c>orderedList</c>, one per nesting level as in ODF. The
/// list style's levels travel as <c>numbering</c> on the outer list; a separate list restarts the
/// count unless it continues another (<c>text:continue-numbering</c>, <c>text:continue-list</c>),
/// and then it shares its key, as Word's lists sharing a definition.
/// </summary>
internal sealed class OdtListReader(OdtReadContext context, OdtBodyReader body)
{
    private readonly Dictionary<string, string> _keysById = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _lastKeyByStyle = new(StringComparer.Ordinal);
    private string? _lastKey;
    private int _keys;

    private sealed record Scope(string? StyleName, List<string> Kinds, JsonArray Levels, string Key)
    {
        public string KindAt(int depth) => depth < Kinds.Count ? Kinds[depth] : OdtListLevels.Bullet;
    }

    public Node Read(XElement list) => ReadList(list, parent: null, depth: 0);

    private Node ReadList(XElement list, Scope? parent, int depth)
    {
        var scope = ScopeOf(list, parent, depth);
        var kind = scope.KindAt(depth);
        var node = Node.Of(kind);
        if (kind == OdtListLevels.Ordered) node.With("start", 1);
        if (parent is null || !ReferenceEquals(parent, scope) || parent.KindAt(depth - 1) != kind)
        {
            node.With("numbering", new JsonObject { ["key"] = scope.Key, ["levels"] = scope.Levels.DeepClone() });
        }

        var items = new List<Node>();
        foreach (var child in list.Elements())
        {
            if (child.Name != Ns.Text + "list-item" && child.Name != Ns.Text + "list-header")
            {
                if (child.Name != Ns.Text + "soft-page-break") OdtNotices.Element(context.Inventory, child);
                continue;
            }

            AddItem(items, ItemContent(child, scope, depth));
        }

        if (items.Count == 0) items.Add(Node.Of("listItem", OdtParagraphReader.Empty()));
        node.Content = items;
        return node;
    }

    private List<Node> ItemContent(XElement item, Scope scope, int depth)
    {
        var content = new List<Node>();
        foreach (var child in item.Elements())
        {
            if (child.Name == Ns.Text + "list") content.Add(ReadList(child, scope, depth + 1));
            else content.AddRange(body.ReadBlock(child, flat: true));
        }

        return content;
    }

    /// <summary>
    /// An item holding only a deeper list is how ODF skips a level: the list goes under the
    /// previous item instead of drawing an empty one.
    /// </summary>
    private static void AddItem(List<Node> items, List<Node> content)
    {
        var opensWithText = content.Count > 0 && content[0].Type is "paragraph" or "heading";
        if (!opensWithText && content.Count > 0 && items.Count > 0 && content.All(IsList))
        {
            (items[^1].Content ??= []).AddRange(content);
            return;
        }

        if (!opensWithText) content.Insert(0, OdtParagraphReader.Empty());
        var listItem = Node.Of("listItem");
        listItem.Content = content;
        items.Add(listItem);
    }

    private static bool IsList(Node node) => node.Type is OdtListLevels.Ordered or OdtListLevels.Bullet;

    /// <summary>A nested list without a style of its own belongs to the outer one.</summary>
    private Scope ScopeOf(XElement list, Scope? parent, int depth)
    {
        var declared = list.Attr(Ns.Text + "style-name");
        if (parent is not null && (declared is null || declared == parent.StyleName)) return parent;

        var styleName = declared ?? FirstParagraphListStyle(list);
        var (kinds, levels) = OdtListLevels.Read(context.Styles.ListStyle(styleName), context.Inventory);
        var restart = RestartOf(list, depth, kinds);
        if (restart is { } start && levels.Count > depth && levels[depth] is JsonObject level) level["start"] = start;

        var key = parent is null && restart is null ? ContinuedKey(list, styleName) : null;
        key ??= $"odt{++_keys}";
        Remember(list, styleName, key, topLevel: parent is null);
        return new Scope(styleName, kinds, levels, key);
    }

    /// <summary>The list style a paragraph names: LibreOffice often puts it there, not on the list.</summary>
    private string? FirstParagraphListStyle(XElement list)
    {
        var paragraph = list.Descendants().FirstOrDefault(element =>
            element.Name == Ns.Text + "p" || element.Name == Ns.Text + "h");
        return context.Styles.ListStyleOfParagraph(paragraph?.Attr(Ns.Text + "style-name"));
    }

    /// <summary>The first item's <c>text:start-value</c> on a numbered level restarts the count.</summary>
    private static int? RestartOf(XElement list, int depth, List<string> kinds)
    {
        var first = list.Elements(Ns.Text + "list-item").FirstOrDefault();
        var ordered = depth < kinds.Count && kinds[depth] == OdtListLevels.Ordered;
        return ordered ? OdtLength.Integer(first?.Attr(Ns.Text + "start-value")) : null;
    }

    private string? ContinuedKey(XElement list, string? styleName)
    {
        if (list.Attr(Ns.Text + "continue-list") is { } id) return _keysById.GetValueOrDefault(id);
        if (list.Attr(Ns.Text + "continue-numbering") != "true") return null;
        return styleName is not null && _lastKeyByStyle.TryGetValue(styleName, out var key) ? key : _lastKey;
    }

    private void Remember(XElement list, string? styleName, string key, bool topLevel)
    {
        if (list.Attr(Ns.Xml + "id") is { } id) _keysById[id] = key;
        if (!topLevel) return;
        if (styleName is not null) _lastKeyByStyle[styleName] = key;
        _lastKey = key;
    }
}
