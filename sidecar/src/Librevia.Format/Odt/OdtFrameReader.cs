using System.Text.Json.Nodes;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>draw:frame</c> → the editor's <c>image</c>, or <c>math</c> for a formula object. A frame
/// carries one rendition the editor uses: any other embedded object or a text box goes to the
/// inventory, even when it brings a preview image.
/// </summary>
internal sealed class OdtFrameReader(OdtReadContext context)
{
    private const string AsCharacter = "as-char";

    private static readonly Dictionary<string, string> MimeByExtension = new(StringComparer.OrdinalIgnoreCase)
    {
        [".png"] = "image/png",
        [".jpg"] = "image/jpeg",
        [".jpeg"] = "image/jpeg",
        [".gif"] = "image/gif",
        [".bmp"] = "image/bmp",
    };

    /// <summary>The ones the DOCX writer embeds; see <c>ImageWriter</c>.</summary>
    private static readonly HashSet<string> Supported = new(StringComparer.Ordinal)
    {
        "image/png", "image/jpeg", "image/gif", "image/bmp",
    };

    public Node? Read(XElement frame)
    {
        if ((frame.Element(Ns.Draw + "object") ?? frame.Element(Ns.Draw + "object-ole")) is { } embedded)
        {
            return ObjectOf(frame, embedded);
        }

        if (frame.Element(Ns.Draw + "text-box") is not null)
        {
            context.NoteLoss(OdtNotices.TextBoxes);
            return null;
        }

        if (frame.Element(Ns.Draw + "image") is not { } image)
        {
            if (frame.Elements().FirstOrDefault(child => child.Name.Namespace == Ns.Draw) is { } other)
            {
                OdtNotices.Element(context.Inventory, other);
            }

            return null;
        }

        return ImageOf(frame, image);
    }

    private Node? ImageOf(XElement frame, XElement image)
    {
        if (SourceOf(image) is not { } src) return null;

        var node = Node.Of("image").With("src", src);
        var alt = frame.Element(Ns.Svg + "desc")?.Value ?? frame.Element(Ns.Svg + "title")?.Value;
        if (!string.IsNullOrEmpty(alt)) node.With("alt", alt);

        if (OdtLength.Pixels(frame.Attr(Ns.Svg + "width")) is { } width and > 0 &&
            OdtLength.Pixels(frame.Attr(Ns.Svg + "height")) is { } height and > 0)
        {
            node.With("width", width);
            node.With("height", height);
        }

        if ((frame.Attr(Ns.Text + "anchor-type") ?? AsCharacter) != AsCharacter)
        {
            context.NoteLoss(OdtNotices.ImagePosition);
        }

        return node;
    }

    /// <summary>A formula is an equation; any other object is named by its tag.</summary>
    private Node? ObjectOf(XElement frame, XElement embedded)
    {
        var path = PathOf(embedded.Attr(Ns.XLink + "href"));
        if (path is null || !context.Package.IsFormula(path))
        {
            OdtNotices.Element(context.Inventory, embedded);
            return null;
        }

        if (embedded.Name.LocalName != "object" || FormulaOf(path) is not { } math)
        {
            context.NoteLoss(Inventory.Equations);
            return null;
        }

        return new Node { Type = "math" }
            .With("mathml", math.ToString(SaveOptions.DisableFormatting))
            .With("latex", frame.Element(Ns.Svg + "desc")?.Value ?? string.Empty)
            .With("display", math.Attr("display") == "block" && IsAlone(frame))
            .With("lossy", new JsonArray())
            .With("editable", true);
    }

    /// <summary>
    /// The MathML of the formula's sub-document, through <see cref="MathMlFilter"/> as an edited one
    /// would be. What the filter drops stays in the object until its block is rewritten.
    /// </summary>
    private XElement? FormulaOf(string folder)
    {
        var bytes = context.Package.ReadBytes(folder.TrimEnd('/') + "/content.xml");
        var root = bytes is null ? null : XmlParts.Read(bytes)?.Root;
        if (root is null || root.Name != MathMlFilter.MathMl + "math") return null;

        var plain = Unprefixed(root);
        var math = MathMlFilter.Sanitize(plain.ToString(SaveOptions.DisableFormatting));
        if (math is not null && math.DescendantsAndSelf().Count() < plain.DescendantsAndSelf().Count()) context.NoteLoss(Inventory.Equations);
        return math;
    }

    /// <summary>
    /// Older LibreOffice writes <c>math:math</c>: the filter takes MathML only without a prefix, and
    /// foreign elements, attributes and comments go.
    /// </summary>
    private static XElement Unprefixed(XElement element) => new(
        element.Name,
        element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration && attribute.Name.Namespace == XNamespace.None),
        element.Nodes().Select(MathNodeOf).OfType<XNode>());

    private static XNode? MathNodeOf(XNode node) => node switch
    {
        XElement inner when inner.Name.Namespace == MathMlFilter.MathMl => Unprefixed(inner),
        XText text => new XText(text.Value),
        _ => null,
    };

    /// <summary>LibreOffice marks every formula <c>block</c>: only one alone in its paragraph is displayed.</summary>
    private static bool IsAlone(XElement frame)
    {
        var paragraph = frame.Ancestors().FirstOrDefault(ancestor => ancestor.Name == Ns.Text + "p" || ancestor.Name == Ns.Text + "h");
        if (paragraph is null || paragraph.Descendants(Ns.Draw + "frame").Count() != 1) return false;
        var outside = paragraph.DescendantNodes().OfType<XText>().Where(text => !text.Ancestors(Ns.Draw + "frame").Any());
        return outside.All(text => string.IsNullOrWhiteSpace(text.Value));
    }

    /// <summary>The bytes inside the package, or in <c>office:binary-data</c>; never a link outside.</summary>
    private string? SourceOf(XElement image)
    {
        byte[]? bytes;
        string? mime;
        if (image.Element(Ns.Office + "binary-data") is { } inline)
        {
            bytes = Base64(inline.Value);
            mime = bytes is null ? null : MimeOfBytes(bytes);
        }
        else
        {
            var path = PathOf(image.Attr(Ns.XLink + "href"));
            bytes = path is null ? null : context.Package.ReadBytes(path);
            mime = path is null ? null : context.Package.MediaTypeOf(path) is { Length: > 0 } declared
                ? declared
                : MimeByExtension.GetValueOrDefault(Path.GetExtension(path));
        }

        if (bytes is null)
        {
            context.NoteLoss(OdtNotices.ImageWithoutContent);
            return null;
        }

        if (mime is null || !Supported.Contains(mime))
        {
            context.NoteLoss(OdtNotices.ImageType, mime ?? "?");
            return null;
        }

        return $"data:{mime};base64,{Convert.ToBase64String(bytes)}";
    }

    /// <summary><c>Pictures/a.png</c> or <c>./Object 1</c>; an absolute or external address is not.</summary>
    private static string? PathOf(string? href)
    {
        if (string.IsNullOrWhiteSpace(href)) return null;
        var path = href.StartsWith("./", StringComparison.Ordinal) ? href[2..] : href;
        return path.Contains(':') || path.StartsWith('/') || path.StartsWith("../", StringComparison.Ordinal) ? null : path;
    }

    private static byte[]? Base64(string text)
    {
        try
        {
            return Convert.FromBase64String(string.Concat(text.Where(letter => !char.IsWhiteSpace(letter))));
        }
        catch (FormatException)
        {
            return null;
        }
    }

    private static string? MimeOfBytes(byte[] bytes) =>
        ImageSize.Of(bytes) is null ? null
        : bytes[0] == 0x89 ? "image/png"
        : bytes[0] == 0xFF ? "image/jpeg"
        : bytes[0] == (byte)'G' ? "image/gif"
        : "image/bmp";
}
