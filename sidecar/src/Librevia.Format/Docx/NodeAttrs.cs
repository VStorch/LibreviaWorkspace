using System.Globalization;
using System.Text.Json;

namespace Librevia.Format.Docx;

/// <summary>
/// Leitura dos atributos de um nó do editor, com a unidade já convertida.
/// </summary>
/// <remarks>
/// Ficava dentro de <see cref="ParagraphWriter"/>, e passou a viver sozinho
/// quando a montagem do `w:pPr` saiu de lá: os dois escritores leem os mesmos
/// atributos, e duas cópias das mesmas conversões é como as unidades divergem.
/// </remarks>
internal static class Attr
{
    public static bool Bool(Node node, string name) =>
        node.Attrs is not null
        && node.Attrs.TryGetValue(name, out var value)
        && value is not null
        && value.GetValueKind() == JsonValueKind.True;

    public static string? String(Node node, string name) =>
        node.Attrs is not null && node.Attrs.TryGetValue(name, out var value) && value is not null
            ? value.GetValueKind() == JsonValueKind.String ? value.GetValue<string>() : null
            : null;

    public static double? Double(Node node, string name)
    {
        if (node.Attrs is null || !node.Attrs.TryGetValue(name, out var value) || value is null) return null;
        return value.GetValueKind() == JsonValueKind.Number ? value.GetValue<double>() : null;
    }

    public static int? Int(Node node, string name)
    {
        if (node.Attrs is null || !node.Attrs.TryGetValue(name, out var value) || value is null) return null;
        return value.GetValueKind() == JsonValueKind.Number ? value.GetValue<int>() : null;
    }

    public static string? MarkString(Mark mark, string name) =>
        mark.Attrs is not null && mark.Attrs.TryGetValue(name, out var value) && value is not null
            ? value.GetValueKind() == JsonValueKind.String ? value.GetValue<string>() : null
            : null;

    /// <summary>
    /// 1 twip = 1/1440 de polegada.
    /// </summary>
    /// <remarks>
    /// Arredonda para longe do zero, e não para o par, que é o padrão do .NET. A
    /// conversão morava em dois lugares com os dois arredondamentos — aqui e na
    /// margem da página —, e duas respostas para o mesmo milímetro é como a
    /// medida que o painel mostra deixa de ser a que o arquivo tem.
    /// </remarks>
    public static int MmToTwips(double mm) =>
        (int)Math.Round(mm * 1440 / 25.4, MidpointRounding.AwayFromZero);

    /// <inheritdoc cref="MmToTwips(double)"/>
    public static int? MmToTwips(double? mm) => mm is null ? null : MmToTwips(mm.Value);

    /// <summary>
    /// Uma medida do CSS em pontos, que é a unidade do OOXML.
    /// </summary>
    /// <remarks>
    /// O pixel do CSS vale três quartos de ponto — 96 px por polegada contra 72
    /// pt. Enquanto a conversão não existia, `font-size: 16px` era lido como
    /// "16" e voltava para o arquivo como 16 pt: um terço maior do que a pessoa
    /// escolheu. Unidade que não sabemos converter devolve <c>null</c>, para
    /// quem chamou registrar a perda em vez de inventar um número.
    /// </remarks>
    public static double? Points(string? css)
    {
        if (string.IsNullOrWhiteSpace(css)) return null;

        var text = css.Trim();
        var digits = text.TrimEnd('%', 'a', 'c', 'e', 'i', 'm', 'n', 'p', 'r', 't', 'x', ' ');
        if (!double.TryParse(digits, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)) return null;

        var unit = text[digits.Length..].Trim().ToLowerInvariant();
        return unit switch
        {
            "" or "pt" => value,
            "px" => value * 0.75,
            "in" => value * 72,
            "cm" => value * 72 / 2.54,
            "mm" => value * 72 / 25.4,
            "pc" => value * 12,
            _ => null,
        };
    }
}
