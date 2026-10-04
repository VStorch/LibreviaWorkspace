using System.Globalization;
using ClosedXML.Excel;

namespace Librevia.Format.Xlsx;

/// <summary>
/// Máscaras do XLSX ↔ os seis formatos do aplicativo, pela **intenção**, e não pela
/// máscara: uma moeda desconhecida vale mais como moeda que como geral.
/// </summary>
public static class NumberFormats
{
    /// <summary>
    /// Por máscara: são meia dúzia numa planilha de centenas de milhares de células.
    /// O teto segura o arquivo com uma máscara por célula. Concorrente porque os
    /// testes rodam em paralelo.
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

        // "General" antes de qualquer heurística: o "a" de General é o de `dd/mm/aaaa`.
        if (IsGeneral(code))
        {
            return (null, null);
        }

        // Os literais saem uma vez só.
        var mask = WithoutLiterals(code);
        return (KindOf(code, mask), DecimalsOf(mask));
    }

    private static bool IsGeneral(string code) =>
        code.Split(';').All(section =>
            section.Trim().Equals("General", StringComparison.OrdinalIgnoreCase)
            || section.Trim().Equals("Geral", StringComparison.OrdinalIgnoreCase));

    /// <summary>Os embutidos do OOXML que aparecem na prática; o resto é geral.</summary>
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
        // Horas: o aplicativo não tem formato de hora, e "data" mostraria a data errada.
        18 or 19 or 20 or 21 or 45 or 46 or 47 => (null, null),
        49 => ("text", null),
        _ => (null, null),
    };

    private static string? KindOf(string code, string mask)
    {
        if (mask.Contains('y') || mask.Contains('a') || mask.Contains('d') || mask.Contains('M'))
        {
            // `m` sozinho é minuto; com dia ou ano, é mês.
            if (mask.Contains('y') || mask.Contains('a') || mask.Contains('d')) return "date";
        }

        if (mask.Contains('%')) return "percent";
        if (code.Contains("R$") || code.Contains('$') || code.Contains('€') || code.Contains("[$")) return "currency";
        if (mask.Contains('@')) return "text";
        if (mask.Contains('0') || mask.Contains('#')) return "number";

        return null;
    }

    /// <summary>Os zeros depois do ponto, que é sempre o decimal da máscara.</summary>
    private static int? DecimalsOf(string mask)
    {
        // A primeira seção (positivo) é a que se vê quase sempre.
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

    /// <summary>Aspas, colchetes e barra invertida: sem isso o <c>d</c> de <c>[sidecar/src/Librevia.Format/Docx/RunReader.cshmtBc416]</c> seria data.</summary>
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

    /// <summary>A moeda sai em reais: o formato do modelo não guarda qual é.</summary>
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

    /// <summary>Para comparar sem regravar.</summary>
    public static bool Matches(IXLNumberFormat existing, string? format, int? decimals)
    {
        var (readFormat, readDecimals) = Read(existing);
        if (readFormat != format) return false;

        // Casas indefinidas casam com o padrão do outro lado.
        var a = decimals ?? DefaultDecimals(format);
        var b = readDecimals ?? DefaultDecimals(readFormat);
        return format is null || a == b;
    }

    internal static string Invariant(double value) => value.ToString(CultureInfo.InvariantCulture);
}
