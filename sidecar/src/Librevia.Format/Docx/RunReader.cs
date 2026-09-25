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

    /// <param name="inherited">
    /// O que o estilo do parágrafo dá aos runs. Com ele, o que o estilo liga e o
    /// run desliga sai como a marca com <c>off</c> — ver <see cref="Off"/>.
    /// </param>
    public static List<Mark>? MarksOf(
        RunProperties? properties,
        string? hyperlink,
        FontTable? fonts = null,
        RunProperties? inherited = null)
    {
        var marks = new List<Mark>();

        if (hyperlink is not null)
        {
            marks.Add(Mark.Of("link", "href", hyperlink));
        }

        if (properties is not null)
        {
            if (IsOn(properties.Bold)) marks.Add(Mark.Of("bold"));
            else if (IsOn(inherited?.Bold)) marks.Add(Off("bold"));
            if (IsOn(properties.Italic)) marks.Add(Mark.Of("italic"));
            else if (IsOn(inherited?.Italic)) marks.Add(Off("italic"));
            if (IsOn(properties.Strike)) marks.Add(Mark.Of("strike"));
            else if (IsOn(inherited?.Strike)) marks.Add(Off("strike"));
            if (IsOn(properties.Caps)) marks.Add(Mark.Of("caps"));
            if (IsOn(properties.SmallCaps)) marks.Add(Mark.Of("smallCaps"));

            // `w:vertAlign` também não é alternância: traz o valor, e
            // `baseline` é o normal — que não é marca nenhuma. Enquanto isto
            // ficava de fora, a fórmula e a nota de referência do documento
            // abriam na linha do texto, e voltavam assim para o arquivo.
            var vertical = properties.VerticalTextAlignment?.Val;
            if (vertical is not null)
            {
                if (vertical.Value == VerticalPositionValues.Superscript) marks.Add(Mark.Of("superscript"));
                else if (vertical.Value == VerticalPositionValues.Subscript) marks.Add(Mark.Of("subscript"));
            }

            // `w:u` não é alternância: carrega o estilo do sublinhado, e "none"
            // é a forma de desligar.
            if (IsUnderlined(properties)) marks.Add(Mark.Of("underline"));
            else if (inherited is not null && IsUnderlined(inherited)) marks.Add(Off("underline"));

            var highlight = HighlightOf(properties);
            if (highlight is not null)
            {
                marks.Add(Mark.Of("highlight", "color", highlight));
            }

            var style = TextStyleOf(properties, fonts);
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
    /// A marca "desligado": o estilo do parágrafo liga, o run desliga.
    /// </summary>
    /// <remarks>
    /// Com o bloco desenhado pelo estilo, a ausência da marca quer dizer "o que o
    /// estilo disser". O trecho que o autor tirou do negrito de um título precisa
    /// dizer outra coisa, e é esta marca que o diz — na tela (`font-weight: 400`)
    /// e de volta no arquivo (`w:b w:val="0"`).
    /// </remarks>
    public static Mark Off(string type) => Mark.Of(type, "off", true);

    private static Mark? TextStyleOf(RunProperties properties, FontTable? fonts)
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

        // Com a substituta genérica atrás: a fonte que o documento pede pode não
        // existir na máquina de quem abre, e sem dizer de que tipo ela é o
        // navegador cai na serifa do fim da pilha. Ver FontTable.
        var font = properties.RunFonts?.Ascii?.Value ?? properties.RunFonts?.HighAnsi?.Value;
        if (!string.IsNullOrWhiteSpace(font)) attributes["fontFamily"] = fonts?.Stack(font) ?? font;

        return attributes.Count == 0 ? null : new Mark { Type = "textStyle", Attrs = attributes };
    }

    /// <summary>
    /// Uma medida em pontos, do jeito que o CSS a escreve.
    /// </summary>
    /// <remarks>
    /// Pública para <see cref="StyleReader"/>: o tamanho da fonte de um estilo é
    /// a mesma medida do tamanho da fonte de um trecho, e duas formatações da
    /// mesma coisa divergiriam no primeiro meio-ponto.
    /// </remarks>
    public static string FormatPoints(double points) =>
        points == Math.Floor(points)
            ? $"{(int)points}pt"
            : points.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "pt";

    /// <summary>
    /// Fundo do texto: `w:highlight` traz nome de cor, `w:shd` traz hexadecimal.
    /// </summary>
    /// <remarks>
    /// Recebe o elemento, e não o `w:rPr` tipado, para servir também ao
    /// <see cref="StyleReader"/>: o OOXML tem uma classe diferente para o `w:rPr`
    /// do trecho, o do estilo e o do padrão do documento, e os três têm este
    /// mesmo filho.
    /// </remarks>
    public static string? HighlightOf(OpenXmlElement properties)
    {
        var highlight = properties.GetFirstChild<Highlight>()?.Val;
        if (highlight is not null && highlight.Value != HighlightColorValues.None)
        {
            return NamedHighlight(highlight.Value.ToString());
        }

        return ColorOf(properties.GetFirstChild<Shading>()?.Fill);
    }

    /// <summary>A cor como o editor a guarda, ou <c>null</c> quando não é cor.</summary>
    public static string? ColorOf(StringValue? value) =>
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
