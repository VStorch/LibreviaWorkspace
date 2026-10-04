using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>A lista de um parágrafo: tipo, nível, marca e recuo.</summary>
/// <param name="NumberingId">O <c>w:numId</c>, que volta na gravação para a lista editada manter a numeração.</param>
/// <param name="Definition">Os nove níveis, a chave da contagem e o reinício — ver <see cref="ListLevels"/>.</param>
public sealed record ListStyle(
    string Kind,
    int NumberingId,
    int Level,
    string? Marker,
    double? IndentMm,
    double? HangingMm,
    JsonObject Definition);

/// <summary>
/// O parágrafo aponta um <c>numId</c>, que aponta um <c>abstractNumId</c>, que guarda o
/// formato por nível. No Word conta a definição abstrata: o "Reiniciar em 1" cria um
/// <c>w:num</c> **com** <c>w:startOverride</c>, contagem à parte. A chave (<c>a7</c>,
/// <c>n12</c>) diz qual das duas vale.
/// </summary>
public sealed class NumberingReader(MainDocumentPart part)
{
    private readonly Dictionary<int, JsonObject> _definitions = [];

    public ListStyle? ListKindOf(ParagraphProperties? properties)
    {
        var numbering = properties?.NumberingProperties;
        var numId = numbering?.NumberingId?.Val?.Value;
        if (numId is null or 0) return null;

        // `numId` que o `numbering.xml` não define não numera nada no Word.
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

    /// <summary>Nula quando o arquivo não a define; a gravação confere com ela o <c>numId</c> do nó.</summary>
    internal JsonObject? FileDefinitionOf(int numId) =>
        InstanceOf(numId) is { } instance && AbstractOf(instance) is not null ? DefinitionOf(numId) : null;

    /// <summary>
    /// Os que faltam seguem o primeiro nível da **definição**, e não do parágrafo:
    /// senão a gravação criaria uma duplicada.
    /// </summary>
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

    /// <summary>Clonada a cada pedido: um <c>JsonNode</c> só pode ter um pai.</summary>
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
                // Nível trocado no `w:num` é outra lista.
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

        // O que a gravação precisa para recriar o nível noutro arquivo, só quando foge do padrão.
        if (level.LevelJustification?.Val?.InnerText is { } jc && jc is not ("left" or "start")) json["jc"] = jc;
        if (level.LevelSuffix?.Val?.InnerText is { } suff && suff != "tab") json["suff"] = suff;
        if (level.LevelRestart?.Val?.Value is { } restart) json["restart"] = restart;

        // Formatação do número e estilo do nível não cabem: a gravação que precisar recriar avisa.
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
    /// A lista presa a <c>w:numStyleLink</c> não tem nível: os níveis moram na definição
    /// com o mesmo <c>w:styleLink</c>.
    /// </summary>
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

        // "none" é numeração desligada naquele nível.
        return format.Value == NumberFormatValues.Bullet || format.Value == NumberFormatValues.None
            ? "bulletList"
            : "orderedList";
    }

    /// <summary>Só na lista com marcador; o <c>%1.</c> da ordenada é de <c>list-numbering.ts</c>.</summary>
    private static string? MarkerOf(Level? definition)
    {
        if (definition?.NumberingFormat?.Val?.Value != NumberFormatValues.Bullet) return null;

        var text = definition.LevelText?.Val?.Value;
        return string.IsNullOrEmpty(text) ? null : ListLevels.Shown(text);
    }

    /// <summary><c>@left</c> é onde o texto começa; <c>@hanging</c>, quanto o marcador fica antes.</summary>
    private static double? TwipsOf(string? value)
    {
        if (value is null || !int.TryParse(value, out var twips) || twips <= 0) return null;
        return ListLevels.Mm(twips);
    }
}
