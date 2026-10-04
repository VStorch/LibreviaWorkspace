using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Librevia.Format.Docx;

/// <summary>
/// Os atributos de um nó do editor, já na unidade do OOXML. Uma cópia só das
/// conversões para os dois escritores.
/// </summary>
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

    /// <summary>O atributo cru, para o que não é escalar, como a lista <c>colwidth</c>.</summary>
    public static JsonNode? Node(Node node, string name) =>
        node.Attrs is not null && node.Attrs.TryGetValue(name, out var value) ? value : null;

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
    /// 1 twip = 1/1440 de polegada. Arredonda para longe do zero, e não para o par
    /// do .NET: o painel e o arquivo mostram a mesma medida.
    /// </summary>
    public static int MmToTwips(double mm) =>
        (int)Math.Round(mm * 1440 / 25.4, MidpointRounding.AwayFromZero);

    /// <inheritdoc cref="MmToTwips(double)"/>
    public static int? MmToTwips(double? mm) => mm is null ? null : MmToTwips(mm.Value);

    /// <summary>
    /// Medida do CSS em pontos: o pixel vale três quartos de ponto (96 px contra
    /// 72 pt por polegada). Unidade desconhecida volta <c>null</c>, para quem chamou
    /// registrar a perda.
    /// </summary>
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
