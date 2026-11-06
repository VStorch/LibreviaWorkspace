using System.Globalization;

namespace Librevia.Format.Odt;

/// <summary>ODF lengths (<c>2cm</c>, <c>0.5in</c>, <c>12pt</c>) in the editor's units.</summary>
internal static class OdtLength
{
    private const int PicasPerInch = 6;
    public const int MillimetresDecimals = 2;
    public const int PercentPerUnit = 100;
    private const int PointsDecimals = 1;
    private const int ExactPointsDecimals = 2;

    /// <summary>Or <c>null</c> for a percentage, an unknown unit or garbage.</summary>
    public static double? Millimetres(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var text = value.Trim();
        var digits = text.TrimEnd('a', 'c', 'i', 'm', 'n', 'p', 't', 'x');
        if (!double.TryParse(digits, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) return null;

        double? inches = text[digits.Length..] switch
        {
            "mm" => number / Unit.MillimetersPerInch,
            "cm" => number / Unit.CentimetersPerInch,
            "in" => number,
            "pt" => number / Unit.PointsPerInch,
            "pc" => number / PicasPerInch,
            "px" => number / Unit.PixelsPerInch,
            _ => null,
        };
        return inches * Unit.MillimetersPerInch;
    }

    public static double? RoundedMillimetres(string? value) =>
        Millimetres(value) is { } mm ? Math.Round(mm, MillimetresDecimals) : null;

    public static double? Points(string? value) =>
        Millimetres(value) is { } mm ? Math.Round(mm * Unit.PointsPerInch / Unit.MillimetersPerInch, PointsDecimals) : null;

    /// <summary>For a border, where a tenth of a point is too coarse.</summary>
    public static double? ExactPoints(string? value) =>
        Millimetres(value) is { } mm ? Math.Round(mm * Unit.PointsPerInch / Unit.MillimetersPerInch, ExactPointsDecimals) : null;

    /// <summary>The editor measures images and columns in whole CSS pixels.</summary>
    public static int? Pixels(string? value) =>
        Millimetres(value) is { } mm ? (int)Math.Round(mm * Unit.PixelsPerInch / Unit.MillimetersPerInch) : null;

    public static int? PixelsOfMillimetres(double? mm) =>
        mm is { } value ? (int)Math.Round(value * Unit.PixelsPerInch / Unit.MillimetersPerInch) : null;

    /// <summary><c>150%</c> → 1.5.</summary>
    public static double? Percent(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || !value.Trim().EndsWith('%')) return null;
        return double.TryParse(value.Trim().TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? number / PercentPerUnit
            : null;
    }

    public static int? Integer(string? value) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) ? number : null;
}
