using System.Text;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The inverse of <see cref="OdtInlineReader"/>: text with marks → spans, links, real spaces, tabs,
/// line breaks and images. ODF collapses white space, so a space after another one, or opening the
/// block, is a <c>text:s</c>.
/// </summary>
internal sealed class OdtInlineWriter(OdtWriteContext context)
{
    private const string Mono = "Liberation Mono";
    private const string DataPrefix = "data:";
    private const string Base64Marker = ";base64,";

    /// <summary>Inline nodes the model has and ODF gets only as text, with the notice they cost.</summary>
    private static readonly Dictionary<string, string> Dropped = new(StringComparer.Ordinal)
    {
        ["commentStart"] = Inventory.Comments,
        ["commentEnd"] = Inventory.Comments,
        ["bookmarkStart"] = OdtNotices.Bookmarks,
        ["bookmarkEnd"] = OdtNotices.Bookmarks,
        ["math"] = Inventory.Equations,
    };

    private static readonly HashSet<string> RevisionMarks = new(StringComparer.Ordinal)
    {
        "insertion", "deletion", "markInsertion", "markDeletion",
    };

    private bool _afterSpace;

    public List<XNode> Write(IEnumerable<Node>? content)
    {
        _afterSpace = true;
        var nodes = new List<XNode>();
        foreach (var node in content ?? []) nodes.AddRange(Inline(node));
        return nodes;
    }

    private IEnumerable<XNode> Inline(Node node)
    {
        switch (node.Type)
        {
            case "text":
                return Wrapped(node.Marks, TextNodes(node.Text ?? string.Empty));
            case "hardBreak":
                _afterSpace = true;
                return [new XElement(Ns.Text + "line-break")];
            case "image":
                return Image(node) is { } frame ? Wrapped(node.Marks, [frame], styled: false) : [];
            case "field":
                return Field(node);
            case "noteRef":
                context.NoteLoss(Attr.String(node, "kind") == "endnote" ? Inventory.Endnotes : Inventory.Footnotes);
                return [];
            default:
                context.NoteLoss(Dropped.GetValueOrDefault(node.Type) ?? Inventory.Element, Dropped.ContainsKey(node.Type) ? null : node.Type);
                return [.. (node.Content ?? []).SelectMany(Inline)];
        }
    }

    /// <summary>
    /// The ODF field of the instruction, with the attributes of the one the block had (its date
    /// style, <c>text:fixed</c>); without an equivalent, only its text.
    /// </summary>
    private IEnumerable<XNode> Field(Node node)
    {
        var result = TextNodes(Attr.String(node, "result") ?? string.Empty);
        if (OdtFields.ElementOf(Attr.String(node, "instr")) is not { } name)
        {
            context.NoteLoss(Inventory.Fields);
            return Wrapped(node.Marks, result);
        }

        var field = context.TakeSourceField(name) is { } original
            ? new XElement(original.Name, original.Attributes())
            : OdtFields.Create(name);
        field.Add(result);
        _afterSpace = false;
        return Wrapped(node.Marks, [field]);
    }

    /// <summary>The span of the marks' style, inside the link when there is one.</summary>
    private IEnumerable<XNode> Wrapped(List<Mark>? marks, List<XNode> nodes, bool styled = true)
    {
        if (nodes.Count == 0) return nodes;
        IEnumerable<XNode> result = nodes;
        if (styled && SpanStyle(marks) is { } style)
        {
            result = [new XElement(Ns.Text + "span", new XAttribute(Ns.Text + "style-name", style), result)];
        }

        var link = marks?.FirstOrDefault(mark => mark.Type == "link");
        if (link is not null && Attr.MarkString(link, "href") is { Length: > 0 } href)
        {
            result = [new XElement(
                Ns.Text + "a",
                new XAttribute(Ns.XLink + "type", "simple"),
                new XAttribute(Ns.XLink + "href", href),
                result)];
        }

        return result;
    }

    /// <summary>A named character style, under an automatic style when the marks add to it.</summary>
    private string? SpanStyle(List<Mark>? marks)
    {
        if (marks is null || marks.Count == 0) return null;
        var named = marks.FirstOrDefault(mark => mark.Type == "charStyle") is { } charStyle
            ? context.NamedStyle(OdtStyles.TextFamily, Attr.MarkString(charStyle, "styleId"))
            : null;
        var properties = OdtFormatWriter.Text(FormatOf(marks), context.Styles.Fonts);
        return properties.HasAttributes
            ? context.Automatic.Style(OdtAutomaticStyles.Text, OdtStyles.TextFamily, named, properties)
            : named;
    }

    private CharacterFormat FormatOf(List<Mark> marks)
    {
        var format = CharacterFormat.Empty;
        foreach (var mark in marks)
        {
            var on = !(mark.Attrs?.TryGetValue("off", out var off) == true && off?.GetValue<bool>() == true);
            format = mark.Type switch
            {
                "bold" => format with { Bold = on },
                "italic" => format with { Italic = on },
                "underline" => format with { Underline = on },
                "strike" => format with { Strike = on },
                "caps" => format with { Caps = on },
                "smallCaps" => format with { SmallCaps = on },
                "superscript" => format with { Position = on ? CharacterFormat.Superscript : CharacterFormat.Baseline },
                "subscript" => format with { Position = on ? CharacterFormat.Subscript : CharacterFormat.Baseline },
                "highlight" => format with { Background = Color(Attr.MarkString(mark, "color")) },
                "code" => format with { FontFamily = Mono },
                "textStyle" => TextStyle(format, mark),
                _ => Other(format, mark),
            };
        }

        return format;
    }

    private static CharacterFormat TextStyle(CharacterFormat format, Mark mark) => format with
    {
        Color = Color(Attr.MarkString(mark, "color")) ?? format.Color,
        FontSizePt = Attr.Points(Attr.MarkString(mark, "fontSize")) ?? format.FontSizePt,
        FontFamily = Attr.MarkString(mark, "fontFamily") is { Length: > 0 } font ? font : format.FontFamily,
        Background = Color(Attr.MarkString(mark, "backgroundColor")) ?? format.Background,
    };

    /// <summary>A comment or a revision has no place in a rewritten ODT block: its text stays.</summary>
    private CharacterFormat Other(CharacterFormat format, Mark mark)
    {
        if (mark.Type is "link" or "charStyle") return format;
        if (RevisionMarks.Contains(mark.Type)) context.NoteLoss(OdtNotices.Revisions);
        else if (mark.Type == "comment") context.NoteLoss(Inventory.Comments);
        else context.NoteLoss(Inventory.Element, mark.Type);
        return format;
    }

    private static string? Color(string? value) => CharacterFormat.ColorOf(value);

    private List<XNode> TextNodes(string text)
    {
        var nodes = new List<XNode>();
        var run = new StringBuilder();
        var spaces = 0;

        void FlushRun()
        {
            if (run.Length == 0) return;
            nodes.Add(new XText(run.ToString()));
            run.Clear();
        }

        void FlushSpaces()
        {
            if (spaces == 0) return;
            nodes.Add(spaces == 1
                ? new XElement(Ns.Text + "s")
                : new XElement(Ns.Text + "s", new XAttribute(Ns.Text + "c", spaces)));
            spaces = 0;
        }

        foreach (var letter in text)
        {
            if (letter == ' ' && !_afterSpace)
            {
                run.Append(' ');
                _afterSpace = true;
            }
            else if (letter == ' ')
            {
                FlushRun();
                spaces++;
            }
            else if (letter is '\t' or '\n')
            {
                FlushRun();
                FlushSpaces();
                nodes.Add(new XElement(Ns.Text + (letter == '\t' ? "tab" : "line-break")));
                _afterSpace = true;
            }
            else if (letter != '\r')
            {
                FlushSpaces();
                run.Append(letter);
                _afterSpace = false;
            }
        }

        FlushRun();
        FlushSpaces();
        return nodes;
    }

    /// <summary>A frame as a character, as the reader keeps it; the bytes go to <see cref="OdtPictures"/>.</summary>
    private XElement? Image(Node node)
    {
        var (bytes, mediaType) = Decode(Attr.String(node, "src"));
        if (bytes is null || mediaType is null)
        {
            context.NoteLoss(OdtNotices.ImageWithoutContent);
            return null;
        }

        if (!OdtPictures.IsSupported(mediaType))
        {
            context.NoteLoss(OdtNotices.ImageType, mediaType);
            return null;
        }

        var natural = ImageSize.Of(bytes);
        var width = Attr.Double(node, "width") ?? natural?.Width ?? 1;
        var height = Attr.Double(node, "height") ?? natural?.Height ?? 1;
        _afterSpace = false;
        var frame = new XElement(
            Ns.Draw + "frame",
            new XAttribute(Ns.Draw + "name", context.NewImageName()),
            new XAttribute(Ns.Text + "anchor-type", "as-char"),
            new XAttribute(Ns.Svg + "width", OdtFormatWriter.PixelsToMm(width)),
            new XAttribute(Ns.Svg + "height", OdtFormatWriter.PixelsToMm(height)),
            new XAttribute(Ns.Draw + "z-index", 0),
            new XElement(
                Ns.Draw + "image",
                new XAttribute(Ns.XLink + "href", context.Pictures.PathOf(bytes, mediaType)),
                new XAttribute(Ns.XLink + "type", "simple"),
                new XAttribute(Ns.XLink + "show", "embed"),
                new XAttribute(Ns.XLink + "actuate", "onLoad")));
        if (Attr.String(node, "alt") is { Length: > 0 } alt) frame.Add(new XElement(Ns.Svg + "desc", alt));
        return frame;
    }

    private static (byte[]? Bytes, string? MediaType) Decode(string? source)
    {
        if (source is null || !source.StartsWith(DataPrefix, StringComparison.Ordinal)) return (null, null);
        var marker = source.IndexOf(Base64Marker, StringComparison.Ordinal);
        if (marker < 0) return (null, null);
        try
        {
            return (Convert.FromBase64String(source[(marker + Base64Marker.Length)..]), source[DataPrefix.Length..marker]);
        }
        catch (FormatException)
        {
            return (null, null);
        }
    }
}
