using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A lista de um parágrafo: que tipo, em que nível, com que marca e recuo.
/// </summary>
public sealed record ListStyle(string Kind, int Level, string? Marker, double? IndentMm, double? HangingMm);

/// <summary>
/// Decide se um parágrafo numerado é lista com marcador ou lista ordenada, e
/// com que marca ele é desenhado.
/// </summary>
/// <remarks>
/// O caminho é indireto de propósito no OOXML: o parágrafo aponta um `numId`,
/// que aponta um `abstractNumId`, que guarda um formato por nível. Ler o
/// formato direto do parágrafo não é possível — ele simplesmente não está lá.
/// </remarks>
public sealed class NumberingReader(MainDocumentPart part)
{
    private readonly Dictionary<int, int> _abstractByNum = Build(part);

    public ListStyle? ListKindOf(ParagraphProperties? properties)
    {
        var numbering = properties?.NumberingProperties;
        var numId = numbering?.NumberingId?.Val?.Value;
        if (numId is null or 0) return null;

        var level = numbering?.NumberingLevelReference?.Val?.Value ?? 0;
        var definition = LevelOf(numId.Value, level);

        return new ListStyle(
            KindOf(definition),
            level,
            MarkerOf(definition),
            TwipsOf(definition?.PreviousParagraphProperties?.Indentation?.Left?.Value),
            TwipsOf(definition?.PreviousParagraphProperties?.Indentation?.Hanging?.Value));
    }

    private Level? LevelOf(int numId, int level)
    {
        var numbering = part.NumberingDefinitionsPart?.Numbering;
        if (numbering is null || !_abstractByNum.TryGetValue(numId, out var abstractId)) return null;

        return numbering.Elements<AbstractNum>()
            .FirstOrDefault(candidate => candidate.AbstractNumberId?.Value == abstractId)
            ?.Elements<Level>()
            .FirstOrDefault(candidate => (candidate.LevelIndex?.Value ?? 0) == level);
    }

    private static string KindOf(Level? definition)
    {
        var format = definition?.NumberingFormat?.Val;
        if (format is null) return "bulletList";

        // "none" também não é ordenada: é numeração desligada naquele nível.
        return format.Value == NumberFormatValues.Bullet || format.Value == NumberFormatValues.None
            ? "bulletList"
            : "orderedList";
    }

    /// <summary>
    /// A marca do nível, como um caractere que qualquer fonte desenha.
    /// </summary>
    /// <remarks>
    /// O `w:lvlText` de uma lista com marcador guarda o caractere, e ele
    /// costuma vir da área de uso privado do Unicode: é assim que o Word grava
    /// os glifos das fontes Symbol e Wingdings. Servido como está, aparece a
    /// caixinha de caractere ausente; trocado pelo equivalente de verdade,
    /// aparece a marca que o documento pede — e não a bolinha que o CSS escolhe
    /// sozinho.
    ///
    /// Só a lista com marcador: a ordenada tem `%1.` e afins, que é uma
    /// gramática de contagem, e quem a resolve é o contador do CSS.
    /// </remarks>
    private static string? MarkerOf(Level? definition)
    {
        if (definition?.NumberingFormat?.Val?.Value != NumberFormatValues.Bullet) return null;

        var text = definition.LevelText?.Val?.Value;
        if (string.IsNullOrEmpty(text)) return null;

        return string.Concat(text.Select(letter => letter switch
        {
            '\uF0A7' or '\uF06E' => '▪',
            '\uF0B7' or '\uF0FC' => '•',
            '\uF0D8' or '\uF0E0' => '➢',
            '\uF06C' => '●',
            '\uF0A8' => '□',
            >= '\uF000' and <= '\uF0FF' => '•',
            _ => letter,
        }));
    }

    /// <summary>
    /// Uma medida de recuo do nível, em milímetros.
    /// </summary>
    /// <remarks>
    /// `w:ind/@left` é onde o **texto** do item começa; `@hanging` é quanto o
    /// marcador fica antes dele. Sem as duas o item sai colado na margem e a
    /// marca encostada na primeira letra.
    /// </remarks>
    private static double? TwipsOf(string? value)
    {
        if (value is null || !int.TryParse(value, out var twips) || twips <= 0) return null;
        return Math.Round(twips / 1440.0 * 25.4, 2);
    }

    private static Dictionary<int, int> Build(MainDocumentPart part)
    {
        var map = new Dictionary<int, int>();
        var numbering = part.NumberingDefinitionsPart?.Numbering;
        if (numbering is null) return map;

        foreach (var instance in numbering.Elements<NumberingInstance>())
        {
            var id = instance.NumberID?.Value;
            var abstractId = instance.AbstractNumId?.Val?.Value;
            if (id is not null && abstractId is not null) map[id.Value] = abstractId.Value;
        }

        return map;
    }
}
