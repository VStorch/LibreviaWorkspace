using System.Globalization;

namespace Librevia.Format.Docx;

/// <summary>
/// Cor do CSS → hexadecimal de seis dígitos, que é a única forma que o OOXML
/// entende.
/// </summary>
/// <remarks>
/// O editor guarda cor como o CSS a escreve, e o CSS tem cinco jeitos de dizer
/// vermelho. O Word tem um: `FF0000`. Enquanto o valor ia direto para o
/// atributo, um `rgb(255, 0, 0)` — que é o que o navegador devolve ao ler uma
/// cor de um `style` — virava `w:color w:val="rgb(255, 0, 0)"`, e o Word abria o
/// documento como danificado. O nome `red`, mais discreto, era aceito e
/// desenhado como preto.
///
/// O canal alfa é descartado de propósito: o OOXML não tem transparência de
/// texto, e aproximá-la sobre um fundo que não conhecemos seria inventar uma cor
/// que ninguém pediu. Cor que não dá para converter volta <c>null</c>, e quem
/// chamou registra a perda — nunca grava um valor inválido no arquivo.
/// </remarks>
internal static class ColorValue
{
    /// <summary>
    /// As cores que o CSS nomeia e aparecem em documento de verdade.
    /// </summary>
    /// <remarks>
    /// As 148 do CSS não cabem aqui sem virar tabela de dados. Estas são as 16
    /// do HTML original — as que as barras de ferramentas oferecem — mais as
    /// cinzas e as que o Word usa nos seus próprios temas. O resto cai no
    /// inventário, que é honesto: melhor avisar que a cor se perdeu do que
    /// gravar um palpite.
    /// </remarks>
    private static readonly Dictionary<string, string> Named = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = "000000",
        ["silver"] = "C0C0C0",
        ["gray"] = "808080",
        ["grey"] = "808080",
        ["white"] = "FFFFFF",
        ["maroon"] = "800000",
        ["red"] = "FF0000",
        ["purple"] = "800080",
        ["fuchsia"] = "FF00FF",
        ["magenta"] = "FF00FF",
        ["green"] = "008000",
        ["lime"] = "00FF00",
        ["olive"] = "808000",
        ["yellow"] = "FFFF00",
        ["navy"] = "000080",
        ["blue"] = "0000FF",
        ["teal"] = "008080",
        ["aqua"] = "00FFFF",
        ["cyan"] = "00FFFF",
        ["orange"] = "FFA500",
        ["pink"] = "FFC0CB",
        ["brown"] = "A52A2A",
        ["gold"] = "FFD700",
        ["darkgray"] = "A9A9A9",
        ["darkgrey"] = "A9A9A9",
        ["lightgray"] = "D3D3D3",
        ["lightgrey"] = "D3D3D3",
        ["whitesmoke"] = "F5F5F5",
    };

    /// <summary>O hexadecimal de seis dígitos, ou <c>null</c> se não der.</summary>
    public static string? Hex(string? css)
    {
        if (string.IsNullOrWhiteSpace(css)) return null;

        var value = css.Trim();

        // "auto" e "transparent" não são cores: são a ausência de uma. Gravá-las
        // como preto poria cor onde o documento não pedia nenhuma.
        if (value.Equals("auto", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("inherit", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("currentcolor", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("transparent", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (Named.TryGetValue(value, out var named)) return named;
        if (value.StartsWith('#')) return FromHex(value[1..]);
        if (value.StartsWith("rgb", StringComparison.OrdinalIgnoreCase)) return FromRgb(value);

        // Já pode ser o hexadecimal do próprio OOXML, sem cerquilha — e é só essa
        // forma que passa: seis dígitos, que é como o Word os grava.
        //
        // A forma curta aqui aceitaria qualquer palavra que por acaso seja
        // hexadecimal: `fade` virava FFAADD, uma cor errada no lugar do aviso de
        // que a cor se perdeu. Nome fora da tabela devolve `null`, e quem chamou
        // registra a perda.
        return value.Length == 6 ? FromHex(value) : null;
    }

    private static string? FromHex(string digits)
    {
        if (!digits.All(Uri.IsHexDigit)) return null;

        // `#rgb` é a forma curta: cada dígito vale por dois. O alfa de `#rrggbbaa`
        // e de `#rgba` é descartado com o resto da transparência.
        return digits.Length switch
        {
            3 or 4 => string.Concat(digits[..3].Select(digit => new string(digit, 2))).ToUpperInvariant(),
            6 or 8 => digits[..6].ToUpperInvariant(),
            _ => null,
        };
    }

    /// <summary>`rgb(255, 0, 0)` e `rgba(255 0 0 / 50%)`, que é o que o navegador devolve.</summary>
    private static string? FromRgb(string value)
    {
        var open = value.IndexOf('(', StringComparison.Ordinal);
        var close = value.LastIndexOf(')');
        if (open < 0 || close < open) return null;

        // Vírgula ou espaço separam os canais, e a barra separa o alfa: as duas
        // gramáticas do CSS moderno, tratadas como uma.
        var parts = value[(open + 1)..close]
            .Split([',', ' ', '/'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length < 3) return null;

        var channels = new int[3];
        for (var index = 0; index < 3; index++)
        {
            if (Channel(parts[index]) is not { } channel) return null;
            channels[index] = channel;
        }

        return $"{channels[0]:X2}{channels[1]:X2}{channels[2]:X2}";
    }

    private static int? Channel(string text)
    {
        var percent = text.EndsWith('%');
        var digits = percent ? text[..^1] : text;
        if (!double.TryParse(digits, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) return null;

        var scaled = percent ? number * 255 / 100 : number;
        return (int)Math.Clamp(Math.Round(scaled), 0, 255);
    }
}
