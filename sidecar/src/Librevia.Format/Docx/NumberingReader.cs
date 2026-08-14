using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Decide se um parágrafo numerado é lista com marcador ou lista ordenada.
/// </summary>
/// <remarks>
/// O caminho é indireto de propósito no OOXML: o parágrafo aponta um `numId`,
/// que aponta um `abstractNumId`, que guarda um formato por nível. Ler o
/// formato direto do parágrafo não é possível — ele simplesmente não está lá.
/// </remarks>
public sealed class NumberingReader(MainDocumentPart part)
{
    private readonly Dictionary<int, int> _abstractByNum = Build(part);

    public (string Kind, int Level)? ListKindOf(ParagraphProperties? properties)
    {
        var numbering = properties?.NumberingProperties;
        var numId = numbering?.NumberingId?.Val?.Value;
        if (numId is null or 0) return null;

        var level = numbering?.NumberingLevelReference?.Val?.Value ?? 0;
        return (KindOf(numId.Value, level), level);
    }

    private string KindOf(int numId, int level)
    {
        var numbering = part.NumberingDefinitionsPart?.Numbering;
        if (numbering is null || !_abstractByNum.TryGetValue(numId, out var abstractId)) return "bulletList";

        var definition = numbering.Elements<AbstractNum>()
            .FirstOrDefault(candidate => candidate.AbstractNumberId?.Value == abstractId);

        var format = definition?.Elements<Level>()
            .FirstOrDefault(candidate => (candidate.LevelIndex?.Value ?? 0) == level)
            ?.NumberingFormat?.Val;

        if (format is null) return "bulletList";

        // "none" também não é ordenada: é numeração desligada naquele nível.
        return format.Value == NumberFormatValues.Bullet || format.Value == NumberFormatValues.None
            ? "bulletList"
            : "orderedList";
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
