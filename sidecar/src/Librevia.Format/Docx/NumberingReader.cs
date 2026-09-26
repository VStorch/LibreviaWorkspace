using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A lista de um parágrafo: que tipo, em que nível, com que marca e recuo.
/// </summary>
/// <param name="NumberingId">
/// O `w:numId` do arquivo. Vai para o nó da lista e volta na gravação: é o que
/// mantém a lista editada apontando a mesma numeração, com a mesma marca e a
/// mesma contagem.
/// </param>
/// <param name="Definition">
/// Os nove níveis da numeração, com a chave da contagem e o reinício do
/// `w:num` — ver <see cref="ListLevels"/>.
/// </param>
public sealed record ListStyle(
    string Kind,
    int NumberingId,
    int Level,
    string? Marker,
    double? IndentMm,
    double? HangingMm,
    JsonObject Definition);

/// <summary>
/// Decide se um parágrafo numerado é lista com marcador ou lista ordenada, e
/// com que definição ele é contado e desenhado.
/// </summary>
/// <remarks>
/// O caminho é indireto de propósito no OOXML: o parágrafo aponta um `numId`,
/// que aponta um `abstractNumId`, que guarda um formato por nível. Ler o
/// formato direto do parágrafo não é possível — ele simplesmente não está lá.
///
/// ## A chave da contagem
///
/// No Word quem conta é a definição abstrata, e não o `w:num`: dois `w:num` que
/// apontam o mesmo `w:abstractNum` continuam a mesma contagem — é por isso que
/// o "Reiniciar em 1" do Word cria um `w:num` novo **com** `w:startOverride`, e
/// não um `w:num` novo e só. O `w:num` com reinício (ou com nível próprio em
/// `w:lvlOverride/w:lvl`) é contagem à parte. A chave (`a7`, `n12`) diz qual das
/// duas vale, e a tela conta por ela.
/// </remarks>
public sealed class NumberingReader(MainDocumentPart part)
{
    private readonly Dictionary<int, JsonObject> _definitions = [];

    public ListStyle? ListKindOf(ParagraphProperties? properties)
    {
        var numbering = properties?.NumberingProperties;
        var numId = numbering?.NumberingId?.Val?.Value;
        if (numId is null or 0) return null;

        // `numId` que o `numbering.xml` não define (ou cuja definição abstrata
        // sumiu) não numera nada no Word: o parágrafo aparece sem marca. Tratado
        // como lista, a tela inventava a numeração padrão — e a gravação criava
        // uma definição e reescrevia parágrafos que ninguém tocou.
        var instance = InstanceOf(numId.Value);
        if (instance is null || AbstractOf(instance) is null) return null;

        var level = Math.Clamp(numbering?.NumberingLevelReference?.Val?.Value ?? 0, 0, ListLevels.Count - 1);
        var definition = LevelOf(numId.Value, level);

        return new ListStyle(
            KindOf(definition),
            numId.Value,
            level,
            MarkerOf(definition),
            TwipsOf(definition?.PreviousParagraphProperties?.Indentation?.Left?.Value),
            TwipsOf(definition?.PreviousParagraphProperties?.Indentation?.Hanging?.Value),
            DefinitionOf(numId.Value));
    }

    /// <summary>
    /// A definição de um `numId` do arquivo, como o leitor a entregaria; nula
    /// quando o arquivo não a define.
    /// </summary>
    /// <remarks>É com ela que a gravação confere se o `numId` que o nó traz ainda é o dele.</remarks>
    internal JsonObject? FileDefinitionOf(int numId) =>
        InstanceOf(numId) is { } instance && AbstractOf(instance) is not null ? DefinitionOf(numId) : null;

    /// <summary>
    /// Os nove níveis de uma definição abstrata, com os que faltam preenchidos.
    /// </summary>
    /// <remarks>
    /// O preenchimento segue o primeiro nível da **definição**, e não o parágrafo
    /// que a pediu: lida por um item com marcador e por outro numerado, a mesma
    /// definição dava dois conjuntos de níveis — e a gravação, que compara níveis
    /// para reaproveitar a definição, criava uma duplicada.
    /// </remarks>
    internal static JsonArray LevelsOf(AbstractNum? abstractNum)
    {
        var first = abstractNum?.Elements<Level>().FirstOrDefault(level => (level.LevelIndex?.Value ?? 0) == 0);
        var levels = ListLevels.Defaults(KindOf(first));
        foreach (var level in abstractNum?.Elements<Level>() ?? [])
        {
            var index = level.LevelIndex?.Value ?? 0;
            if (index is >= 0 and < ListLevels.Count) levels[index] = LevelJson(level);
        }

        return levels;
    }

    /// <summary>
    /// A definição inteira de um `w:num`, como o nó da lista a leva.
    /// </summary>
    /// <remarks>
    /// Clonada a cada pedido: o mesmo `numId` aparece em várias listas, e um
    /// `JsonNode` só pode ter um pai.
    /// </remarks>
    private JsonObject DefinitionOf(int numId)
    {
        if (!_definitions.TryGetValue(numId, out var cached))
        {
            cached = Build(numId);
            _definitions[numId] = cached;
        }

        return (JsonObject)cached.DeepClone();
    }

    private JsonObject Build(int numId)
    {
        var instance = InstanceOf(numId);
        var abstractNum = AbstractOf(instance);
        var abstractId = abstractNum?.AbstractNumberId?.Value;

        var levels = LevelsOf(abstractNum);

        var overrides = new JsonObject();
        var ownLevels = false;
        foreach (var levelOverride in instance?.Elements<LevelOverride>() ?? [])
        {
            var index = levelOverride.LevelIndex?.Value ?? 0;
            if (index is < 0 or >= ListLevels.Count) continue;

            if (levelOverride.Level is { } replaced)
            {
                // Nível inteiro trocado no `w:num`: é outra lista, com outra marca.
                levels[index] = LevelJson(replaced);
                ownLevels = true;
            }

            if (levelOverride.StartOverrideNumberingValue?.Val?.Value is { } start)
            {
                overrides[index.ToString(System.Globalization.CultureInfo.InvariantCulture)] = start;
            }
        }

        var separate = ownLevels || overrides.Count > 0 || abstractId is null;
        var definition = new JsonObject
        {
            ["key"] = separate ? $"n{numId}" : $"a{abstractId}",
        };
        if (abstractId is not null) definition["abstractId"] = abstractId.Value;
        definition["levels"] = levels;
        if (overrides.Count > 0) definition["overrides"] = overrides;
        return definition;
    }

    internal static JsonObject LevelJson(Level level)
    {
        var format = ListLevels.FormatName(level);
        var text = level.LevelText?.Val?.Value ?? string.Empty;
        var json = ListLevels.Level(
            format,
            format == "bullet" ? ListLevels.Shown(text) : text,
            level.StartNumberingValue?.Val?.Value ?? 1,
            TwipsOf(level.PreviousParagraphProperties?.Indentation?.Left?.Value),
            TwipsOf(level.PreviousParagraphProperties?.Indentation?.Hanging?.Value),
            level.IsLegalNumberingStyle is { } legal && (legal.Val?.Value ?? true));

        // O que a tela não desenha mas a gravação precisa para recriar o nível
        // noutro arquivo (lista colada): alinhamento, o que vem depois do número
        // e o reinício. Só quando foge do padrão, para não engordar toda lista.
        if (level.LevelJustification?.Val?.InnerText is { } jc && jc is not ("left" or "start")) json["jc"] = jc;
        if (level.LevelSuffix?.Val?.InnerText is { } suff && suff != "tab") json["suff"] = suff;
        if (level.LevelRestart?.Val?.Value is { } restart) json["restart"] = restart;

        // A formatação do número (cor, tamanho, negrito) e o estilo ligado ao
        // nível não cabem na definição: a gravação que precisar recriar o nível
        // sem o original avisa a perda. A fonte do marcador não conta — ela sai
        // do próprio glifo (`ListLevels.GlyphOf`).
        var numberFormatting = level.NumberingSymbolRunProperties?.ChildElements
            .Any(child => child is not RunFonts) ?? false;
        if (numberFormatting || level.ParagraphStyleIdInLevel is not null)
        {
            json["extra"] = true;
        }

        return json;
    }

    private Level? LevelOf(int numId, int level)
    {
        var instance = InstanceOf(numId);
        var replaced = instance?.Elements<LevelOverride>()
            .FirstOrDefault(candidate => (candidate.LevelIndex?.Value ?? 0) == level)?.Level;
        if (replaced is not null) return replaced;

        return AbstractOf(instance)?.Elements<Level>()
            .FirstOrDefault(candidate => (candidate.LevelIndex?.Value ?? 0) == level);
    }

    private NumberingInstance? InstanceOf(int numId) =>
        part.NumberingDefinitionsPart?.Numbering?.Elements<NumberingInstance>()
            .FirstOrDefault(candidate => candidate.NumberID?.Value == numId);

    /// <summary>
    /// A definição abstrata que de fato tem os níveis.
    /// </summary>
    /// <remarks>
    /// Uma lista presa a um estilo de numeração (`w:numStyleLink`) não tem nível
    /// nenhum: os níveis moram na definição que declara o mesmo estilo em
    /// `w:styleLink`. Sem seguir o elo, a lista "Lista numerada" do Word abria
    /// com marcador.
    /// </remarks>
    private AbstractNum? AbstractOf(NumberingInstance? instance)
    {
        var numbering = part.NumberingDefinitionsPart?.Numbering;
        var abstractId = instance?.AbstractNumId?.Val?.Value;
        if (numbering is null || abstractId is null) return null;

        var found = numbering.Elements<AbstractNum>()
            .FirstOrDefault(candidate => candidate.AbstractNumberId?.Value == abstractId);

        if (found?.NumberingStyleLink?.Val?.Value is { } style && !found.Elements<Level>().Any())
        {
            return numbering.Elements<AbstractNum>()
                .FirstOrDefault(candidate => candidate.StyleLink?.Val?.Value == style) ?? found;
        }

        return found;
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

    /// <summary>A marca do nível, como um caractere que qualquer fonte desenha.</summary>
    /// <remarks>
    /// Só a lista com marcador: a ordenada tem `%1.` e afins, que é uma gramática
    /// de contagem, e quem a resolve é `list-numbering.ts`.
    /// </remarks>
    private static string? MarkerOf(Level? definition)
    {
        if (definition?.NumberingFormat?.Val?.Value != NumberFormatValues.Bullet) return null;

        var text = definition.LevelText?.Val?.Value;
        return string.IsNullOrEmpty(text) ? null : ListLevels.Shown(text);
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
        return ListLevels.Mm(twips);
    }
}
