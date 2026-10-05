using System.Globalization;
using ClosedXML.Excel;

namespace Librevia.Format.Xlsx;

/// <summary>
/// XLSX masks ↔ the app's six formats, by **intent**, not by mask: an unknown currency is worth
/// more as currency than as general.
/// </summary>
public static class NumberFormats
{
    /// <summary>
    /// By mask: there are half a dozen in a sheet of hundreds of thousands of cells. The cap holds
    /// a file with one mask per cell. Concurrent because tests run in parallel.
    /// </summary>
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<
        (string Code, int Id), (string? Format, int? Decimals)> Known = new();

    private const int MaxKnownFormats = 4096;

    public static (string? Format, int? Decimals) Read(IXLNumberFormat format)
    {
        var key = (format.Format ?? string.Empty, format.NumberFormatId);
        if (Known.TryGetValue(key, out var known)) return known;

        var computed = Interpret(key.Item1, key.Item2);
        if (Known.Count < MaxKnownFormats) Known[key] = computed;
        return computed;
    }

    private static (string? Format, int? Decimals) Interpret(string code, int numberFormatId)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            return FromBuiltin(numberFormatId);
        }

        // "General" before any heuristic: the "a" in General is the one in `dd/mm/aaaa`.
        if (IsGeneral(code))
        {
            return (null, null);
        }

        // Literals go once.
        var mask = WithoutLiterals(code);
        return (KindOf(code, mask), DecimalsOf(mask));
    }

    private static bool IsGeneral(string code) =>
        code.Split(';').All(section =>
            section.Trim().Equals("General", StringComparison.OrdinalIgnoreCase)
            || section.Trim().Equals("Geral", StringComparison.OrdinalIgnoreCase));

    /// <summary>OOXML builtins that appear in practice; the rest is general.</summary>
    private static (string? Format, int? Decimals) FromBuiltin(int id) => id switch
    {
        0 => (null, null),
        1 => ("number", 0),
        2 => ("number", 2),
        3 => ("number", 0),
        4 => ("number", 2),
        9 => ("percent", 0),
        10 => ("percent", 2),
        14 or 15 or 16 or 17 or 22 => ("date", null),
        5 or 6 or 37 or 38 or 41 or 42 => ("currency", 0),
        7 or 8 or 39 or 40 or 43 or 44 => ("currency", 2),
        // Hours: the app has no time format, and "date" would show the wrong date.
        18 or 19 or 20 or 21 or 45 or 46 or 47 => (null, null),
        49 => ("text", null),
        _ => (null, null),
    };

    private static string? KindOf(string code, string mask)
    {
        if (mask.Contains('y') || mask.Contains('a') || mask.Contains('d') || mask.Contains('M'))
        {
            // `m` alone is minutes; with day or year, it is month.
            if (mask.Contains('y') || mask.Contains('a') || mask.Contains('d')) return "date";
        }

        if (mask.Contains('%')) return "percent";
        if (code.Contains("R$") || code.Contains('$') || code.Contains('€') || code.Contains("[$")) return "currency";
        if (mask.Contains('@')) return "text";
        if (mask.Contains('0') || mask.Contains('#')) return "number";

        return null;
    }

    /// <summary>The zeros after the dot, which is always the mask's decimal.</summary>
    private static int? DecimalsOf(string mask)
    {
        // The first section (positive) is the one almost always seen.
        var first = mask.Split(';')[0];
        var dot = first.LastIndexOf('.');
        if (dot < 0) return first.Contains('0') || first.Contains('#') ? 0 : null;

        var digits = 0;
        for (var at = dot + 1; at < first.Length; at++)
        {
            if (first[at] is '0' or '#') digits++;
            else break;
        }

        return digits;
    }

    /// <summary>
    /// Quotes, brackets and backslash: what they enclose is not mask, so <c>[$R$-416]</c> is not a
    /// date.
    /// </summary>
    private static string WithoutLiterals(string code)
    {
        var clean = new System.Text.StringBuilder(code.Length);
        var inQuotes = false;
        var brackets = 0;
        var escaped = false;

        foreach (var character in code)
        {
            if (escaped)
            {
                escaped = false;
                continue;
            }

            if (character == '\\' && !inQuotes)
            {
                escaped = true;
                continue;
            }

            switch (character)
            {
                case '"':
                    inQuotes = !inQuotes;
                    continue;
                case '[' when !inQuotes:
                    brackets++;
                    continue;
                case ']' when !inQuotes && brackets > 0:
                    brackets--;
                    continue;
            }

            if (!inQuotes && brackets == 0) clean.Append(character);
        }

        return clean.ToString();
    }

    /// <summary>Currency comes out in reais: the model's format does not keep which one.</summary>
    public static string? Mask(string? format, int? decimals)
    {
        var places = Math.Clamp(decimals ?? DefaultDecimals(format), 0, 10);
        var fraction = places == 0 ? string.Empty : "." + new string('0', places);

        return format switch
        {
            "number" => "#,##0" + fraction,
            "currency" => "\"R$\" #,##0" + fraction,
            "percent" => "0" + fraction + "%",
            "date" => "dd/mm/yyyy",
            "text" => "@",
            _ => null,
        };
    }

    private static int DefaultDecimals(string? format) => format switch
    {
        "currency" => 2,
        "percent" => 0,
        _ => 0,
    };

    /// <summary>To compare without rewriting.</summary>
    public static bool Matches(IXLNumberFormat existing, string? format, int? decimals)
    {
        var (readFormat, readDecimals) = Read(existing);
        if (readFormat != format) return false;

        // Undefined decimals match the other side's default.
        var a = decimals ?? DefaultDecimals(format);
        var b = readDecimals ?? DefaultDecimals(readFormat);
        return format is null || a == b;
    }

    internal static string Invariant(double value) => value.ToString(CultureInfo.InvariantCulture);
}
