using System.Globalization;
using ClosedXML.Excel;

namespace Librevia.Format.Xlsx;

/// <summary>
/// Formato numérico do XLSX ↔ os seis formatos do aplicativo.
/// </summary>
/// <remarks>
/// O XLSX descreve formato por um código de máscara — <c>#.##0,00</c>,
/// <c>[$R$-416]</c>, <c>dd/mm/aaaa</c> — com uma variedade infinita. O modelo do
/// aplicativo tem seis. A tradução é necessariamente aproximada, e a escolha
/// aqui é sempre a mesma: **reconhecer a intenção**, não a máscara. Uma célula
/// com máscara de moeda desconhecida vale mais como "moeda" do que como "geral",
/// porque é o que o usuário vê e é o que ele espera continuar vendo.
/// </remarks>
public static class NumberFormats
{
    /// <summary>Máscara → (formato do aplicativo, casas decimais).</summary>
    public static (string? Format, int? Decimals) Read(IXLNumberFormat format)
    {
        var code = format.Format;
        if (string.IsNullOrWhiteSpace(code))
        {
            return FromBuiltin(format.NumberFormatId);
        }

        // "General" escrito por extenso é formato nenhum — e precisa sair antes
        // de qualquer heurística, porque o "a" de Gener**a**l é o mesmo "a" de
        // `dd/mm/aaaa`. Sem esta linha, toda planilha do LibreOffice abre com a
        // coluna de quantidades cheia de datas de 1900. Medido em arquivo real.
        if (IsGeneral(code))
        {
            return (null, null);
        }

        return (KindOf(code), DecimalsOf(code));
    }

    /// <summary>O nome do formato geral, como o Excel e o LibreOffice o escrevem.</summary>
    private static bool IsGeneral(string code) =>
        code.Split(';').All(section =>
            section.Trim().Equals("General", StringComparison.OrdinalIgnoreCase)
            || section.Trim().Equals("Geral", StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Códigos embutidos do formato, que vêm sem máscara escrita.
    /// </summary>
    /// <remarks>
    /// A tabela é do padrão OOXML. Só entram os que aparecem na prática — o
    /// resto cai em "geral", que é o comportamento correto para o desconhecido.
    /// </remarks>
    private static (string? Format, int? Decimals) FromBuiltin(int id) => id switch
    {
        0 => (null, null),
        1 => ("number", 0),
        2 => ("number", 2),
        3 => ("number", 0),
        4 => ("number", 2),
        9 => ("percent", 0),
        10 => ("percent", 2),
        // 14 a 17 e 22 são as máscaras de data e data-hora.
        14 or 15 or 16 or 17 or 22 => ("date", null),
        // 5 a 8 e 41 a 44 são as de moeda e contábil.
        5 or 6 or 37 or 38 or 41 or 42 => ("currency", 0),
        7 or 8 or 39 or 40 or 43 or 44 => ("currency", 2),
        // 18 a 21 e 45 a 47 são horas: o aplicativo não tem formato de hora, e
        // "data" mostraria a data errada de um valor que é só hora.
        18 or 19 or 20 or 21 or 45 or 46 or 47 => (null, null),
        49 => ("text", null),
        _ => (null, null),
    };

    private static string? KindOf(string code)
    {
        var mask = WithoutLiterals(code);

        // Data antes de número: `dd/mm/aaaa` também tem dígitos na máscara.
        if (mask.Contains('y') || mask.Contains('a') || mask.Contains('d') || mask.Contains('M'))
        {
            // `m` sozinho é minuto; com dia ou ano por perto, é mês.
            if (mask.Contains('y') || mask.Contains('a') || mask.Contains('d')) return "date";
        }

        if (mask.Contains('%')) return "percent";
        if (code.Contains("R$") || code.Contains('$') || code.Contains('€') || code.Contains("[$")) return "currency";
        if (mask.Contains('@')) return "text";
        if (mask.Contains('0') || mask.Contains('#')) return "number";

        return null;
    }

    /// <summary>
    /// Quantas casas decimais a máscara pede.
    /// </summary>
    /// <remarks>
    /// Conta os zeros depois do separador decimal. A máscara do formato usa
    /// sempre ponto para decimal, independente do idioma do arquivo.
    /// </remarks>
    private static int? DecimalsOf(string code)
    {
        var mask = WithoutLiterals(code);

        // Uma máscara pode ter seção para positivo, negativo e zero. A primeira
        // é a que manda no que se vê quase sempre.
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
    /// Tira o que está entre aspas, entre colchetes e escapado por barra.
    /// </summary>
    /// <remarks>
    /// Sem isso, o `d` de `[$R$-416]` faria uma máscara de moeda ser lida como
    /// data — e a coluna de preços apareceria como um punhado de datas de 1900.
    /// A barra invertida é a terceira forma de escrever literal (`\d` é a letra
    /// d, não o dia), e cai na mesma armadilha.
    /// </remarks>
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

    /// <summary>
    /// O caminho de volta: formato do aplicativo → máscara para gravar.
    /// </summary>
    /// <remarks>
    /// A máscara de moeda sai em reais porque é a moeda do usuário deste
    /// aplicativo — e porque o formato "moeda" do modelo não guarda qual é.
    /// </remarks>
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

    /// <summary>Máscara equivalente à do modelo, para comparar sem regravar.</summary>
    public static bool Matches(IXLNumberFormat existing, string? format, int? decimals)
    {
        var (readFormat, readDecimals) = Read(existing);
        if (readFormat != format) return false;

        // Casas indefinidas de um lado casam com o padrão do outro: o modelo
        // omite o que nunca foi escolhido à mão.
        var a = decimals ?? DefaultDecimals(format);
        var b = readDecimals ?? DefaultDecimals(readFormat);
        return format is null || a == b;
    }

    internal static string Invariant(double value) => value.ToString(CultureInfo.InvariantCulture);
}
