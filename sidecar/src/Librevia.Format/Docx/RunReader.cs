using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>Formatação de caractere do OOXML → marcas do editor.</summary>
public static class RunReader
{
    /// <summary><c>&lt;w:b/&gt;</c> liga, <c>&lt;w:b w:val="false"/&gt;</c> desliga: a presença não basta.</summary>
    public static bool IsOn(OnOffType? toggle) =>
        toggle is not null && (toggle.Val is null || toggle.Val.Value);

    /// <param name="inherited">Com ele, o que o estilo liga e o run desliga sai como <see cref="Off"/>.</param>
    /// <param name="directOnly">
    /// No parágrafo desenhado pelos estilos, só sai marca do que **difere** do
    /// herdado, como no escritor (<c>ParagraphWriter.DropWhatRepeatsTheStyle</c>).
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

            // `baseline` é o normal. Emitido mesmo quando repete o estilo, como o
            // realce: o CSS dos estilos não os desenha.
            var vertical = properties.VerticalTextAlignment?.Val;
            if (vertical is not null)
            {
                if (vertical.Value == VerticalPositionValues.Superscript) marks.Add(Mark.Of("superscript"));
                else if (vertical.Value == VerticalPositionValues.Subscript) marks.Add(Mark.Of("subscript"));
            }

            // `w:u` carrega o estilo do sublinhado, e "none" desliga.
            if (IsUnderlined(properties)) { if (from is null || !IsUnderlined(from)) marks.Add(Mark.Of("underline")); }
            else if (inherited is not null && IsUnderlined(inherited)) marks.Add(Off("underline"));

            var highlight = HighlightOf(properties);
            if (highlight is not null)
            {
                marks.Add(Mark.Of("highlight", "color", highlight));
            }

            var style = TextStyleOf(properties, fonts);

            // Campo a campo: a cor que repete o estilo sai, o tamanho que difere fica.
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

    /// <summary>O trecho que o autor tirou do negrito de um título: <c>font-weight: 400</c> na tela, <c>w:b w:val="0"</c> no arquivo.</summary>
    public static Mark Off(string type) => Mark.Of(type, "off", true);

    private static Mark? TextStyleOf(RunProperties properties, FontTable? fonts)
    {
        var attributes = new Dictionary<string, System.Text.Json.Nodes.JsonNode?>();

        var color = ColorOf(properties.Color?.Val);
        if (color is not null) attributes["color"] = color;

        // `w:sz` vem em meios-pontos.
        if (properties.FontSize?.Val is not null &&
            double.TryParse(properties.FontSize.Val.Value, out var halfPoints))
        {
            attributes["fontSize"] = FormatPoints(halfPoints / 2);
        }

        // Com a substituta genérica atrás (FontTable).
        var font = properties.RunFonts?.Ascii?.Value ?? properties.RunFonts?.HighAnsi?.Value;
        if (!string.IsNullOrWhiteSpace(font)) attributes["fontFamily"] = fonts?.Stack(font) ?? font;

        return attributes.Count == 0 ? null : new Mark { Type = "textStyle", Attrs = attributes };
    }

    /// <summary>Pública para <see cref="StyleReader"/>: a mesma medida no estilo e no trecho.</summary>
    public static string FormatPoints(double points) =>
        points == Math.Floor(points)
            ? $"{(int)points}pt"
            : points.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "pt";

    /// <summary>
    /// <c>w:highlight</c> traz nome, <c>w:shd</c> traz hexadecimal. Pelo elemento, para
    /// servir aos três <c>w:rPr</c> do OOXML (trecho, estilo, padrão).
    /// </summary>
    public static string? HighlightOf(OpenXmlElement properties)
    {
        var highlight = properties.GetFirstChild<Highlight>()?.Val;
        if (highlight is not null && highlight.Value != HighlightColorValues.None)
        {
            return NamedHighlight(highlight.Value.ToString());
        }

        return ColorOf(properties.GetFirstChild<Shading>()?.Fill);
    }

    /// <summary>Ou <c>null</c> quando não é cor.</summary>
    public static string? ColorOf(StringValue? value) =>
        IsRealColor(value?.Value) ? "#" + value!.Value!.TrimStart('#').ToLowerInvariant() : null;

    /// <summary>"auto" não é cor: ignorá-lo evita gravar preto onde não se pediu.</summary>
    private static bool IsRealColor(string? value) =>
        !string.IsNullOrWhiteSpace(value) &&
        !value.Equals("auto", StringComparison.OrdinalIgnoreCase) &&
        !value.Equals("none", StringComparison.OrdinalIgnoreCase);

    /// <summary>Os valores do Word.</summary>
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
