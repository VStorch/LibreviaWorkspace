using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Librevia.Format.Docx;

/// <summary>
/// An editor node's attributes, already in OOXML units. A single copy of the conversions for both
/// writers.
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

    /// <summary>The raw attribute, for what is not scalar, like the <c>colwidth</c> list.</summary>
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
    /// Rounds away from zero, not to even as .NET does: the panel and the file show the same
    /// measure.
    /// </summary>
    public static int MmToTwips(double mm) =>
        (int)Math.Round(mm * Unit.TwipsPerInch / Unit.MillimetersPerInch, MidpointRounding.AwayFromZero);

    /// <inheritdoc cref="MmToTwips(double)"/>
    public static int? MmToTwips(double? mm) => mm is null ? null : MmToTwips(mm.Value);

    /// <summary>
    /// A CSS measure in points. An unknown unit returns <c>null</c>, so the caller records the
    /// loss.
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
            "px" => value * ((double)Unit.PointsPerInch / Unit.PixelsPerInch),
            "in" => value * Unit.PointsPerInch,
            "cm" => value * Unit.PointsPerInch / Unit.CentimetersPerInch,
            "mm" => value * Unit.PointsPerInch / Unit.MillimetersPerInch,
            "pc" => value * 12,
            _ => null,
        };
    }
}
