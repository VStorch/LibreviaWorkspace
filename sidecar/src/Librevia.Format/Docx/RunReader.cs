using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>OOXML character formatting → editor marks.</summary>
public static class RunReader
{
    /// <c>&lt;w:b/&gt;</c> turns on, <c>&lt;w:b w:val="false"/&gt;</c> turns off: presence is not
    /// enough.
    public static bool IsOn(OnOffType? toggle) =>
        toggle is not null && (toggle.Val is null || toggle.Val.Value);

    /// <param name="inherited">With it, what the style turns on and the run turns off comes out as
    /// <see cref="Off"/>.</param>
    /// <param name="directOnly">
    /// On a paragraph drawn by styles, only marks that **differ** from the inherited ones come out,
    /// as in the writer (<c>ParagraphWriter.DropWhatRepeatsTheStyle</c>).
    /// </param>
    public static List<Mark>? MarksOf(
        RunProperties? properties,
        string? hyperlink,
        FontTable? fonts = null,
        RunProperties? inherited = null,
        bool directOnly = false)
    {
        var from = directOnly ? inherited : null;

        var marks = new List<Mark>();

        if (hyperlink is not null)
        {
            marks.Add(Mark.Of("link", "href", hyperlink));
        }

        if (properties is not null)
        {
            if (IsOn(properties.Bold)) { if (!IsOn(from?.Bold)) marks.Add(Mark.Of("bold")); }
            else if (IsOn(inherited?.Bold)) marks.Add(Off("bold"));
            if (IsOn(properties.Italic)) { if (!IsOn(from?.Italic)) marks.Add(Mark.Of("italic")); }
            else if (IsOn(inherited?.Italic)) marks.Add(Off("italic"));
            if (IsOn(properties.Strike)) { if (!IsOn(from?.Strike)) marks.Add(Mark.Of("strike")); }
            else if (IsOn(inherited?.Strike)) marks.Add(Off("strike"));
            if (IsOn(properties.Caps) && !IsOn(from?.Caps)) marks.Add(Mark.Of("caps"));
            if (IsOn(properties.SmallCaps) && !IsOn(from?.SmallCaps)) marks.Add(Mark.Of("smallCaps"));

            // `baseline` is normal. Emitted even when it repeats the style, like highlight: the
            // style CSS does not draw them.
            var vertical = properties.VerticalTextAlignment?.Val;
            if (vertical is not null)
            {
                if (vertical.Value == VerticalPositionValues.Superscript) marks.Add(Mark.Of("superscript"));
                else if (vertical.Value == VerticalPositionValues.Subscript) marks.Add(Mark.Of("subscript"));
            }

            // `w:u` carries the underline style, and "none" turns it off.
            if (IsUnderlined(properties)) { if (from is null || !IsUnderlined(from)) marks.Add(Mark.Of("underline")); }
            else if (inherited is not null && IsUnderlined(inherited)) marks.Add(Off("underline"));

            var highlight = HighlightOf(properties);
            if (highlight is not null)
            {
                marks.Add(Mark.Of("highlight", "color", highlight));
            }

            var style = TextStyleOf(properties, fonts);

            // Field by field: a color repeating the style goes, a size that differs stays.
            if (style?.Attrs is { } attributes && from is not null && TextStyleOf(from, fonts)?.Attrs is { } baseline)
            {
                foreach (var (name, value) in baseline)
                {
                    if (attributes.TryGetValue(name, out var own) && JsonNode.DeepEquals(own, value))
                    {
                        attributes.Remove(name);
                    }
                }

                if (attributes.Count == 0) style = null;
            }

            if (style is not null)
            {
                marks.Add(style);
            }
        }

        return marks.Count == 0 ? null : marks;
    }

    private static bool IsUnderlined(RunProperties properties) =>
        properties.Underline?.Val is not null && properties.Underline.Val.Value != UnderlineValues.None;

    /// <summary>
    /// A range the author un-bolded in a heading: <c>font-weight: 400</c> on screen, <c>w:b
    /// w:val="0"</c> in the file.
    /// </summary>
    public static Mark Off(string type) => Mark.Of(type, "off", true);

    private static Mark? TextStyleOf(RunProperties properties, FontTable? fonts)
    {
        var attributes = new Dictionary<string, System.Text.Json.Nodes.JsonNode?>();

        var color = ColorOf(properties.Color?.Val);
        if (color is not null) attributes["color"] = color;

        if (properties.FontSize?.Val is not null &&
            double.TryParse(properties.FontSize.Val.Value, out var halfPoints))
        {
            attributes["fontSize"] = FormatPoints(Unit.HalfPointsToPoints(halfPoints));
        }

        // With the generic substitute behind it (FontTable).
        var font = properties.RunFonts?.Ascii?.Value ?? properties.RunFonts?.HighAnsi?.Value;
        if (!string.IsNullOrWhiteSpace(font)) attributes["fontFamily"] = fonts?.Stack(font) ?? font;

        return attributes.Count == 0 ? null : new Mark { Type = "textStyle", Attrs = attributes };
    }

    /// <summary>
    /// Public for <see cref="StyleReader"/>: the same measure in the style and in the run.
    /// </summary>
    public static string FormatPoints(double points) =>
        points == Math.Floor(points)
            ? $"{(int)points}pt"
            : points.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "pt";

    /// <c>w:highlight</c> carries a name, <c>w:shd</c> a hex value. By element, to serve OOXML's
    /// three <c>w:rPr</c>s (run, style, default).
    public static string? HighlightOf(OpenXmlElement properties)
    {
        var highlight = properties.GetFirstChild<Highlight>()?.Val;
        if (highlight is not null && highlight.Value != HighlightColorValues.None)
        {
            return NamedHighlight(highlight.Value.ToString());
        }

        return ColorOf(properties.GetFirstChild<Shading>()?.Fill);
    }

    /// <summary>Or <c>null</c> when it is not a color.</summary>
    public static string? ColorOf(StringValue? value) =>
        IsRealColor(value?.Value) ? "#" + value!.Value!.TrimStart('#').ToLowerInvariant() : null;

    /// <summary>
    /// "auto" is not a color: ignoring it avoids writing black where none was asked for.
    /// </summary>
    private static bool IsRealColor(string? value) =>
        !string.IsNullOrWhiteSpace(value) &&
        !value.Equals("auto", StringComparison.OrdinalIgnoreCase) &&
        !value.Equals("none", StringComparison.OrdinalIgnoreCase);

    /// <summary>Word's values.</summary>
    private static string NamedHighlight(string name) => name.ToLowerInvariant() switch
    {
        "yellow" => "#ffff00",
        "green" => "#00ff00",
        "cyan" => "#00ffff",
        "magenta" => "#ff00ff",
        "blue" => "#0000ff",
        "red" => "#ff0000",
        "darkblue" => "#000080",
        "darkcyan" => "#008080",
        "darkgreen" => "#008000",
        "darkmagenta" => "#800080",
        "darkred" => "#800000",
        "darkyellow" => "#808000",
        "darkgray" => "#808080",
        "lightgray" => "#c0c0c0",
        "black" => "#000000",
        _ => "#ffff00",
    };
}
