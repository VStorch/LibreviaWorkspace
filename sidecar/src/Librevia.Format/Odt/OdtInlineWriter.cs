using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The inverse of <see cref="OdtInlineReader"/>: text with marks → spans, links, real spaces, tabs,
/// line breaks and images. ODF collapses white space, so a space after another one, or opening the
/// block, is a <c>text:s</c>.
/// </summary>
internal sealed partial class OdtInlineWriter(OdtWriteContext context, OdtBlockWriter blocks)
{
    private const string DefaultHighlight = "#ffff00";
    private const string PageToken = "{n}";

    /// <summary>The same as <c>ALLOWED_EXTERNAL_PROTOCOLS</c>.</summary>
    private static readonly HashSet<string> SafeSchemes = new(StringComparer.OrdinalIgnoreCase) { "http", "https", "mailto" };
    private const string TotalToken = "{total}";

    public const string Mono = "Liberation Mono";
    private const string DataPrefix = "data:";
    private const string Base64Marker = ";base64,";

    private static readonly HashSet<string> RevisionMarks = new(StringComparer.Ordinal)
    {
        "insertion", "deletion", "markInsertion", "markDeletion",
    };

    private readonly OdtInlineObjects _objects = new(context, blocks);
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
            case "text" when context.Changes is null && Deleted(node.Marks):
                context.NoteLoss(OdtNotices.Revisions);
                return [];
            case "text":
                return Wrapped(node.Marks, context.InBand ? BandText(node.Text ?? string.Empty) : TextNodes(node.Text ?? string.Empty));
            case "hardBreak":
                _afterSpace = true;
                return [new XElement(Ns.Text + "line-break")];
            case "image":
                return Image(node) is { } frame ? Wrapped(node.Marks, [frame], styled: false) : [];
            case "field":
                return Field(node);
            case var type when OdtInlineObjects.Types.Contains(type):
                return Objects(node);
            default:
                context.NoteLoss(Inventory.Element, node.Type);
                return [.. (node.Content ?? []).SelectMany(Inline)];
        }
    }

    /// <summary>A formula is in the run's span and link, as a character; a note or anchor is not.</summary>
    private IEnumerable<XNode> Objects(Node node)
    {
        var written = _objects.Write(node).ToList();
        if (node.Type != "math" || written.Count == 0) return written;
        _afterSpace = false;
        return Wrapped(node.Marks, written);
    }

    /// <summary>What "Accept all" removes: the deleted text goes, the revision is declared.</summary>
    private static bool Deleted(List<Mark>? marks) => marks?.Any(mark => mark.Type == "deletion") == true;

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
        if (link is not null && SafeHref(Attr.MarkString(link, "href")) is { } href)
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
                "highlight" => format with { Background = Color(Attr.MarkString(mark, "color")) ?? DefaultHighlight },
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

    /// <summary>
    /// A comment mark has no place in a rewritten ODT block, nor a revision outside the body: its
    /// text stays. In the body, <see cref="Write(IReadOnlyList{Node}, Func{Node, XElement})"/> writes the revision.
    /// </summary>
    private CharacterFormat Other(CharacterFormat format, Mark mark)
    {
        if (mark.Type is "link" or "charStyle") return format;
        if (context.Changes is not null && mark.Type is Revisions.Insertion or Revisions.Deletion) return format;
        if (RevisionMarks.Contains(mark.Type)) context.NoteLoss(OdtNotices.Revisions);
        else if (mark.Type == "comment") context.NoteLoss(Inventory.Comments);
        else context.NoteLoss(Inventory.Element, mark.Type);
        return format;
    }

    /// <summary>Hex as the editor writes it, or a CSS color by name or <c>rgb()</c>.</summary>
    private static string? Color(string? value) =>
        CharacterFormat.ColorOf(value) ?? (ColorValue.Hex(value) is { } hex ? "#" + hex.ToLowerInvariant() : null);

    /// <summary>An internal anchor or a protocol the editor opens; <c>javascript:</c>, <c>file:</c> and <c>data:</c> become text.</summary>
    public static string? SafeHref(string? href)
    {
        var trimmed = href?.Trim();
        if (string.IsNullOrEmpty(trimmed)) return null;
        if (trimmed.StartsWith('#')) return trimmed.Length > 1 ? trimmed : null;
        return Uri.TryCreate(trimmed, UriKind.Absolute, out var uri) && SafeSchemes.Contains(uri.Scheme) ? trimmed : null;
    }

    private List<XNode> TextNodes(string text)
    {
        var builder = new TextBuilder(_afterSpace);
        builder.Append(text);
        _afterSpace = builder.AfterSpace;
        return builder.Nodes();
    }

    /// <summary>A text on its own, as a block's first text: a comment's paragraph, a band's piece.</summary>
    /// <param name="pageFields">Whether <c>{n}</c> and <c>{total}</c> are the page fields.</param>
    public static List<XNode> PlainText(string text, bool pageFields = false)
    {
        var builder = new TextBuilder(afterSpace: true);
        if (pageFields) builder.AppendWithFields(text);
        else builder.Append(text);
        return builder.Nodes();
    }

    private List<XNode> BandText(string text)
    {
        var builder = new TextBuilder(_afterSpace);
        builder.AppendWithFields(text);
        _afterSpace = builder.AfterSpace;
        return builder.Nodes();
    }

    /// <summary>ODF collapses white space: a space after another, or opening the block, is <c>text:s</c>.</summary>
    private sealed class TextBuilder(bool afterSpace)
    {
        private readonly List<XNode> _nodes = [];
        private readonly StringBuilder _run = new();
        private int _spaces;

        public bool AfterSpace { get; private set; } = afterSpace;

        public void Append(string text)
        {
            foreach (var letter in text) Append(letter);
        }

        public void AppendWithFields(string text)
        {
            foreach (var part in BandFields().Split(text))
            {
                if (part is not (PageToken or TotalToken))
                {
                    Append(part);
                    continue;
                }

                FlushRun();
                FlushSpaces();
                _nodes.Add(OdtFields.Create(part == PageToken ? "page-number" : "page-count", "1"));
                AfterSpace = false;
            }
        }

        public List<XNode> Nodes()
        {
            FlushRun();
            FlushSpaces();
            return _nodes;
        }

        private void Append(char letter)
        {
            if (letter == ' ' && !AfterSpace)
            {
                _run.Append(' ');
                AfterSpace = true;
            }
            else if (letter == ' ')
            {
                FlushRun();
                _spaces++;
            }
            else if (letter is '\t' or '\n')
            {
                FlushRun();
                FlushSpaces();
                _nodes.Add(new XElement(Ns.Text + (letter == '\t' ? "tab" : "line-break")));
                AfterSpace = true;
            }
            else if (letter != '\r')
            {
                FlushSpaces();
                _run.Append(letter);
                AfterSpace = false;
            }
        }

        private void FlushRun()
        {
            if (_run.Length == 0) return;
            _nodes.Add(new XText(_run.ToString()));
            _run.Clear();
        }

        private void FlushSpaces()
        {
            if (_spaces == 0) return;
            _nodes.Add(_spaces == 1
                ? new XElement(Ns.Text + "s")
                : new XElement(Ns.Text + "s", new XAttribute(Ns.Text + "c", _spaces)));
            _spaces = 0;
        }
    }

    /// <summary>A frame as a character, as the reader keeps it; the bytes go to <see cref="OdtPictures"/>.</summary>
    private XElement? Image(Node node)
    {
        var source = Attr.String(node, "src");
        if (Picture(context, source) is not { } path) return null;

        var natural = NaturalSize(source);
        _afterSpace = false;
        var frame = CharacterFrame(context, path, Attr.Double(node, "width") ?? natural?.Width ?? 1, Attr.Double(node, "height") ?? natural?.Height ?? 1);
        if (Attr.String(node, "alt") is { Length: > 0 } alt) frame.Add(new XElement(Ns.Svg + "desc", alt));
        return frame;
    }

    /// <summary>A picture as a character, in CSS pixels.</summary>
    public static XElement CharacterFrame(OdtWriteContext context, string path, double width, double height) => new(
        Ns.Draw + "frame",
        new XAttribute(Ns.Draw + "name", context.NewImageName()),
        new XAttribute(Ns.Text + "anchor-type", "as-char"),
        new XAttribute(Ns.Svg + "width", OdtFormatWriter.PixelsToMm(width)),
        new XAttribute(Ns.Svg + "height", OdtFormatWriter.PixelsToMm(height)),
        new XAttribute(Ns.Draw + "z-index", 0),
        ImageElement(path));

    /// <summary>The package path of a <c>data:</c> image, or <c>null</c>, declared, when it has no place there.</summary>
    public static string? Picture(OdtWriteContext context, string? source)
    {
        var (bytes, mediaType) = Decode(source);
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

        return context.Pictures.PathOf(bytes, mediaType);
    }

    /// <summary>The natural size in pixels of a <c>data:</c> image, when its header says it.</summary>
    public static (int Width, int Height)? NaturalSize(string? source) =>
        Decode(source).Bytes is { } bytes && ImageSize.Of(bytes) is { } size ? (size.Width, size.Height) : null;

    public static XElement ImageElement(string path) => new(
        Ns.Draw + "image",
        new XAttribute(Ns.XLink + "href", path),
        new XAttribute(Ns.XLink + "type", "simple"),
        new XAttribute(Ns.XLink + "show", "embed"),
        new XAttribute(Ns.XLink + "actuate", "onLoad"));

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

    [GeneratedRegex(@"(\{n\}|\{total\})")]
    private static partial Regex BandFields();
}
