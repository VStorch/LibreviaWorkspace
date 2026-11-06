using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>draw:frame</c> → the editor's <c>image</c>. A frame carries one rendition the editor uses: an
/// embedded object or a text box goes to the inventory, even when it brings a preview image.
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
            NoteObject(embedded);
            return null;
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
    private void NoteObject(XElement embedded)
    {
        var path = PathOf(embedded.Attr(Ns.XLink + "href"));
        if (path is not null && context.Package.IsFormula(path)) context.NoteLoss(Inventory.Equations);
        else OdtNotices.Element(context.Inventory, embedded);
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
