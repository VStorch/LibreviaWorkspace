using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Traduz a formatação de caractere do OOXML para as marcas do editor.
/// </summary>
public static class RunReader
{
    /// <summary>
    /// Propriedade de alternância do OOXML: <c>&lt;w:b/&gt;</c> liga,
    /// <c>&lt;w:b w:val="false"/&gt;</c> desliga.
    /// </summary>
    /// <remarks>
    /// Ler a simples presença do elemento é o erro clássico aqui: um estilo que
    /// desliga negrito explicitamente viraria negrito ligado.
    /// </remarks>
    public static bool IsOn(OnOffType? toggle) =>
        toggle is not null && (toggle.Val is null || toggle.Val.Value);

    public static List<Mark>? MarksOf(RunProperties? properties, string? hyperlink)
    {
        var marks = new List<Mark>();

        if (hyperlink is not null)
        {
            marks.Add(Mark.Of("link", "href", hyperlink));
        }

        if (properties is not null)
        {
            if (IsOn(properties.Bold)) marks.Add(Mark.Of("bold"));
            if (IsOn(properties.Italic)) marks.Add(Mark.Of("italic"));
            if (IsOn(properties.Strike)) marks.Add(Mark.Of("strike"));
            if (IsOn(properties.Caps)) marks.Add(Mark.Of("caps"));
            if (IsOn(properties.SmallCaps)) marks.Add(Mark.Of("smallCaps"));

            // `w:u` não é alternância: carrega o estilo do sublinhado, e "none"
            // é a forma de desligar.
            if (properties.Underline?.Val is not null &&
                properties.Underline.Val.Value != UnderlineValues.None)
            {
                marks.Add(Mark.Of("underline"));
            }

            var highlight = HighlightOf(properties);
            if (highlight is not null)
            {
                marks.Add(Mark.Of("highlight", "color", highlight));
            }

            var style = TextStyleOf(properties);
            if (style is not null)
            {
                marks.Add(style);
            }
        }

        return marks.Count == 0 ? null : marks;
    }

    private static Mark? TextStyleOf(RunProperties properties)
    {
        var attributes = new Dictionary<string, System.Text.Json.Nodes.JsonNode?>();

        var color = ColorOf(properties.Color?.Val);
        if (color is not null) attributes["color"] = color;

        // `w:sz` vem em meios-pontos: 20 significa 10 pt.
        if (properties.FontSize?.Val is not null &&
            double.TryParse(properties.FontSize.Val.Value, out var halfPoints))
        {
            attributes["fontSize"] = FormatPoints(halfPoints / 2);
        }

        var font = properties.RunFonts?.Ascii?.Value ?? properties.RunFonts?.HighAnsi?.Value;
        if (!string.IsNullOrWhiteSpace(font)) attributes["fontFamily"] = font;

        return attributes.Count == 0 ? null : new Mark { Type = "textStyle", Attrs = attributes };
    }

    private static string FormatPoints(double points) =>
        points == Math.Floor(points)
            ? $"{(int)points}pt"
            : points.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "pt";

    /// <summary>
    /// Fundo do texto: `w:highlight` traz nome de cor, `w:shd` traz hexadecimal.
    /// </summary>
    private static string? HighlightOf(RunProperties properties)
    {
        if (properties.Highlight?.Val is not null &&
            properties.Highlight.Val.Value != HighlightColorValues.None)
        {
            return NamedHighlight(properties.Highlight.Val.Value.ToString());
        }

        var fill = properties.Shading?.Fill?.Value;
        return IsRealColor(fill) ? "#" + fill!.TrimStart('#').ToLowerInvariant() : null;
    }

    private static string? ColorOf(StringValue? value) =>
        IsRealColor(value?.Value) ? "#" + value!.Value!.TrimStart('#').ToLowerInvariant() : null;

    /// <summary>
    /// "auto" quer dizer "decida você" e não é uma cor; ignorá-lo evita gravar
    /// preto explícito onde o documento não pedia cor nenhuma.
    /// </summary>
    private static bool IsRealColor(string? value) =>
        !string.IsNullOrWhiteSpace(value) &&
        !value.Equals("auto", StringComparison.OrdinalIgnoreCase) &&
        !value.Equals("none", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// O OOXML nomeia as cores de destaque; o editor guarda hexadecimal.
    /// Os valores são os do Word.
    /// </summary>
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
