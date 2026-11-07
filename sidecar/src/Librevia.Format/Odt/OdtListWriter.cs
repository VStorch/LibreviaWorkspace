using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The inverse of <see cref="OdtListReader"/>: a list → <c>text:list</c> with an automatic list
/// style made from its <c>numbering</c> levels. A list whose key an earlier list already has
/// continues its count (<c>text:continue-numbering</c>), under the same style, as Word counts a
/// definition across lists. ODF allows no table in a list item: it goes right after the list.
/// </summary>
internal sealed partial class OdtListWriter(OdtWriteContext context, OdtBlockWriter blocks)
{
    private static readonly Dictionary<string, string> NumberFormats = new(StringComparer.Ordinal)
    {
        ["decimal"] = "1",
        ["decimalZero"] = "1",
        ["lowerRoman"] = "i",
        ["upperRoman"] = "I",
        ["lowerLetter"] = "a",
        ["upperLetter"] = "A",
        ["none"] = string.Empty,
    };

    private const string DefaultBullet = "•";

    private List<XElement> _hoisted = [];
    private OdtPageStart? _page;

    /// <param name="page">The master page the list's first paragraph opens.</param>
    public List<XElement> Write(Node list, OdtPageStart? page = null)
    {
        var outer = _hoisted;
        _hoisted = [];
        _page = page;
        var written = new List<XElement> { List(list) };
        written.AddRange(_hoisted);
        _hoisted = outer;
        return written;
    }

    private XElement List(Node list)
    {
        var numbering = Attr.Node(list, "numbering") as JsonObject;
        var key = numbering?["key"] is JsonValue value && value.TryGetValue<string>(out var text) ? text : null;

        var element = new XElement(Ns.Text + "list");
        if (key is not null && context.ListStyles.TryGetValue(key, out var known))
        {
            element.Add(new XAttribute(Ns.Text + "style-name", known), new XAttribute(Ns.Text + "continue-numbering", "true"));
        }
        else
        {
            var style = context.Automatic.Add(OdtAutomaticStyles.List, ListStyle(list, numbering?["levels"] as JsonArray, numbering?["overrides"] as JsonObject));
            element.Add(new XAttribute(Ns.Text + "style-name", style));
            if (key is not null) context.ListStyles[key] = style;
        }

        Items(element, list);
        return element;
    }

    private void Items(XElement element, Node list)
    {
        foreach (var item in list.Content ?? [])
        {
            var listItem = new XElement(Ns.Text + "list-item");
            foreach (var child in item.Content ?? [])
            {
                if (child.Type is OdtListLevels.Bullet or OdtListLevels.Ordered)
                {
                    var nested = new XElement(Ns.Text + "list");
                    Items(nested, child);
                    listItem.Add(nested);
                }
                else if (child.Type == "table")
                {
                    _hoisted.AddRange(blocks.Write(child, flat: true));
                }
                else
                {
                    listItem.Add(blocks.Write(child, flat: true, page: child.Type is "paragraph" or "heading" ? TakePage() : null));
                }
            }

            if (!listItem.HasElements) listItem.Add(new XElement(Ns.Text + "p"));
            element.Add(listItem);
        }
    }

    private OdtPageStart? TakePage()
    {
        var page = _page;
        _page = null;
        return page;
    }

    /// <summary>
    /// One level per depth the list reaches, each of the kind of the list at that depth. A restart
    /// (<c>w:startOverride</c>) is the level's start: the list that opens the count states it.
    /// </summary>
    private static XElement ListStyle(Node list, JsonArray? levels, JsonObject? overrides)
    {
        var kinds = new List<string>();
        KindsOf(list, 0, kinds);
        var defined = levels ?? ListLevels.Defaults(list.Type);
        var style = new XElement(Ns.Text + "list-style");
        for (var depth = 0; depth < Math.Max(kinds.Count, defined.Count) && depth < ListLevels.Count; depth++)
        {
            var level = defined.Count > depth ? defined[depth] as JsonObject : null;
            level ??= ListLevels.Defaults(list.Type)[depth] as JsonObject;
            var kind = depth < kinds.Count ? kinds[depth] : KindOfLevel(level);
            var restart = Number(overrides, depth.ToString(CultureInfo.InvariantCulture));
            style.Add(kind == OdtListLevels.Bullet ? Bullet(depth, level) : Numbered(depth, level, restart));
        }

        return style;
    }

    private static void KindsOf(Node list, int depth, List<string> kinds)
    {
        if (depth >= kinds.Count) kinds.Add(list.Type);
        foreach (var child in (list.Content ?? []).SelectMany(item => item.Content ?? []))
        {
            if (child.Type is OdtListLevels.Bullet or OdtListLevels.Ordered) KindsOf(child, depth + 1, kinds);
        }
    }

    private static string KindOfLevel(JsonObject? level) =>
        Text(level, "fmt") == "bullet" ? OdtListLevels.Bullet : OdtListLevels.Ordered;

    private static XElement Bullet(int depth, JsonObject? level)
    {
        var shown = Text(level, "text") is { Length: > 0 } text ? text : DefaultBullet;
        return new XElement(
            Ns.Text + "list-level-style-bullet",
            new XAttribute(Ns.Text + "level", depth + 1),
            new XAttribute(Ns.Text + "bullet-char", shown.EnumerateRunes().First().ToString()),
            Placement(depth, level));
    }

    /// <summary>
    /// LibreOffice's own form of the number text (<c>%1%.%2%.</c>), and the prefix, suffix and levels
    /// shown for any other reader.
    /// </summary>
    private static XElement Numbered(int depth, JsonObject? level, double? restart)
    {
        var text = Text(level, "text") ?? $"%{depth + 1}.";
        var placeholders = Placeholder().Matches(text);
        var prefix = placeholders.Count == 0 ? text : text[..placeholders[0].Index];
        var suffix = placeholders.Count == 0 ? string.Empty : text[(placeholders[^1].Index + placeholders[^1].Length)..];
        var format = placeholders.Count == 0 ? string.Empty : NumberFormats.GetValueOrDefault(Text(level, "fmt") ?? "decimal", "1");
        var start = (int)(restart ?? Number(level, "start") ?? 1);

        var element = new XElement(
            Ns.Text + "list-level-style-number",
            new XAttribute(Ns.Text + "level", depth + 1),
            new XAttribute(Ns.Loext + "num-list-format", Placeholder().Replace(text, "%$1%")),
            new XAttribute(Ns.Style + "num-prefix", prefix),
            new XAttribute(Ns.Style + "num-suffix", suffix),
            new XAttribute(Ns.Style + "num-format", format));
        if (placeholders.Count > 1) element.Add(new XAttribute(Ns.Text + "display-levels", placeholders.Count));
        element.Add(new XAttribute(Ns.Text + "start-value", start), Placement(depth, level));
        return element;
    }

    /// <summary>The label-alignment mode, the one the reader prefers.</summary>
    private static XElement Placement(int depth, JsonObject? level)
    {
        var indent = Millimetres(level, "indentMm") ?? ListLevels.Mm((depth + 1) * ListLevels.IndentStepTwips);
        var hanging = Millimetres(level, "hangingMm") ?? 0;
        return new XElement(
            Ns.Style + "list-level-properties",
            new XAttribute(Ns.Text + "list-level-position-and-space-mode", "label-alignment"),
            new XElement(
                Ns.Style + "list-level-label-alignment",
                new XAttribute(Ns.Text + "label-followed-by", "listtab"),
                new XAttribute(Ns.Text + "list-tab-stop-position", OdtFormatWriter.Mm(indent)!),
                new XAttribute(Ns.Fo + "text-indent", OdtFormatWriter.Mm(-hanging)!),
                new XAttribute(Ns.Fo + "margin-left", OdtFormatWriter.Mm(indent)!)));
    }

    private static string? Text(JsonObject? level, string name) =>
        level?[name] is JsonValue value && value.TryGetValue<string>(out var text) ? text : null;

    private static double? Millimetres(JsonObject? level, string name) => Number(level, name);

    private static double? Number(JsonObject? level, string name) =>
        level?[name] is JsonValue value &&
        double.TryParse(value.ToJsonString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? number
            : null;

    [GeneratedRegex("%([0-9])")]
    private static partial Regex Placeholder();
}
