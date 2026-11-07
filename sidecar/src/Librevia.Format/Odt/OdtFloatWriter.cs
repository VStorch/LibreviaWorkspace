using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// An anchored image or text box (<c>floating.ts</c>) → a frame anchored to the paragraph, at the
/// position the model gives, as close as ODF allows. A text box grows with its text, as in Word: the
/// model's height is its minimum. The line under a corporate header has no ODF frame and goes.
/// </summary>
internal sealed class OdtFloatWriter(OdtWriteContext context, OdtBlockWriter blocks)
{
    private const string Rule = "rule";
    private const string Image = "image";
    private const double MinSizeMm = 1;
    private const double MinLineWidthPt = 0.25;
    private const double DefaultLineWidthPt = 0.75;

    private static readonly Dictionary<string, string> Relative = new(StringComparer.Ordinal)
    {
        ["page"] = "page",
        ["margin"] = "page-content",
        ["column"] = "paragraph",
        ["paragraph"] = "paragraph",
        ["character"] = "char",
        ["line"] = "paragraph",
    };

    private static readonly Dictionary<string, string> Wraps = new(StringComparer.Ordinal)
    {
        ["none"] = "run-through",
        ["square"] = "parallel",
        ["tight"] = "parallel",
        ["through"] = "parallel",
        ["topAndBottom"] = "none",
    };

    private static readonly HashSet<string> HorizontalAligns = new(StringComparer.Ordinal) { "left", "center", "right", "inside", "outside" };

    public static List<FloatDto> FloatsOf(Node node) =>
        Attr.Node(node, "floats") is JsonArray floats ? Parse(floats) : [];

    public static List<FloatDto> Parse(JsonArray floats)
    {
        try
        {
            return JsonSerializer.Deserialize<List<FloatDto>>(floats, DocxJson.Options) ?? [];
        }
        catch (JsonException)
        {
            return [];
        }
    }

    public IEnumerable<XElement> Write(IEnumerable<FloatDto> floats) => floats.Select(Frame).OfType<XElement>();

    private XElement? Frame(FloatDto anchored)
    {
        if (anchored.Kind == Rule) return null;
        var body = anchored.Kind == Image ? ImageOf(anchored) : TextBox(anchored);
        if (body is null) return null;

        var horizontal = HorizontalAligns.Contains(anchored.HorizontalAlign ?? string.Empty) ? anchored.HorizontalAlign : null;
        var vertical = anchored.VerticalAlign switch
        {
            "center" => "middle",
            "top" or "bottom" => anchored.VerticalAlign,
            _ => null,
        };
        var style = context.Automatic.Style(OdtAutomaticStyles.Graphic, OdtAutomaticStyles.GraphicFamily, null, Properties(anchored, horizontal, vertical));
        var frame = new XElement(
            Ns.Draw + "frame",
            new XAttribute(Ns.Draw + "style-name", style),
            new XAttribute(Ns.Draw + "name", context.NewFrameName()),
            new XAttribute(Ns.Text + "anchor-type", "paragraph"));
        if (horizontal is null) frame.Add(new XAttribute(Ns.Svg + "x", OdtFormatWriter.Mm((anchored.HorizontalOffsetMm ?? 0) + anchored.DxMm)!));
        if (vertical is null) frame.Add(new XAttribute(Ns.Svg + "y", OdtFormatWriter.Mm((anchored.VerticalOffsetMm ?? 0) + anchored.DyMm)!));
        frame.Add(new XAttribute(Ns.Svg + "width", OdtFormatWriter.Mm(Math.Max(MinSizeMm, anchored.WidthMm))!));
        if (anchored.Kind == Image) frame.Add(new XAttribute(Ns.Svg + "height", OdtFormatWriter.Mm(Math.Max(MinSizeMm, anchored.HeightMm))!));
        frame.Add(new XAttribute(Ns.Draw + "z-index", 0), body);
        return frame;
    }

    private XElement? ImageOf(FloatDto anchored)
    {
        if (OdtInlineWriter.Picture(context, anchored.Src) is not { } path) return null;
        return OdtInlineWriter.ImageElement(path);
    }

    private XElement TextBox(FloatDto anchored)
    {
        context.NotelessDepth++;
        var content = (anchored.Content ?? []).SelectMany(child => blocks.Write(child, flat: true)).ToList();
        context.NotelessDepth--;
        if (content.Count == 0) content.Add(new XElement(Ns.Text + "p"));
        return new XElement(
            Ns.Draw + "text-box",
            new XAttribute(Ns.Fo + "min-height", OdtFormatWriter.Mm(Math.Max(MinSizeMm, anchored.HeightMm))!),
            content);
    }

    /// <summary>Without a declared fill it is transparent: LibreOffice's default paints it blue.</summary>
    private static XElement Properties(FloatDto anchored, string? horizontal, string? vertical)
    {
        var properties = new XElement(
            Ns.Style + "graphic-properties",
            new XAttribute(Ns.Style + "wrap", anchored.Behind ? "run-through" : Wraps.GetValueOrDefault(anchored.Wrap, "parallel")),
            new XAttribute(Ns.Style + "run-through", anchored.Behind ? "background" : "foreground"),
            new XAttribute(Ns.Style + "horizontal-pos", horizontal ?? "from-left"),
            new XAttribute(Ns.Style + "horizontal-rel", Relative.GetValueOrDefault(anchored.HorizontalFrom, "paragraph")),
            new XAttribute(Ns.Style + "vertical-pos", vertical ?? "from-top"),
            new XAttribute(Ns.Style + "vertical-rel", Relative.GetValueOrDefault(anchored.VerticalFrom, "paragraph")));
        if (CharacterFormat.ColorOf(anchored.Fill) is { } fill)
        {
            properties.Add(new XAttribute(Ns.Fo + "background-color", fill), new XAttribute(Ns.Draw + "fill", "solid"), new XAttribute(Ns.Draw + "fill-color", fill));
        }
        else
        {
            properties.Add(new XAttribute(Ns.Fo + "background-color", "transparent"), new XAttribute(Ns.Draw + "fill", "none"));
        }

        var line = CharacterFormat.ColorOf(anchored.Line);
        var width = OdtFormatWriter.Pt(Math.Max(MinLineWidthPt, anchored.LineWidthPt == 0 ? DefaultLineWidthPt : anchored.LineWidthPt));
        properties.Add(new XAttribute(Ns.Fo + "border", line is null ? "none" : $"{width} solid {line}"), new XAttribute(Ns.Fo + "padding", "0mm"));
        return properties;
    }
}
