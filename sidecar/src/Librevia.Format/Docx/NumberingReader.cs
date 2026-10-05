using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>A paragraph's list: kind, level, marker and indent.</summary>
/// <param name="NumberingId">The <c>w:numId</c>, which goes back when saving so the edited list
/// keeps its numbering.</param>
/// <param name="Definition">The nine levels, the count key and the restart; see <see
/// cref="ListLevels"/>.</param>
public sealed record ListStyle(
    string Kind,
    int NumberingId,
    int Level,
    string? Marker,
    double? IndentMm,
    double? HangingMm,
    JsonObject Definition);

/// <summary>
/// A paragraph points to a <c>numId</c>, which points to an <c>abstractNumId</c>, which holds the
/// format per level. Word counts by the abstract definition: "Restart at 1" creates a <c>w:num</c>
/// **with** <c>w:startOverride</c>, a separate count. The key (<c>a7</c>, <c>n12</c>) says which
/// applies.
/// </summary>
public sealed class NumberingReader(MainDocumentPart part)
{
    private readonly Dictionary<int, JsonObject> _definitions = [];

    public ListStyle? ListKindOf(ParagraphProperties? properties)
    {
        var numbering = properties?.NumberingProperties;
        var numId = numbering?.NumberingId?.Val?.Value;
        if (numId is null or 0) return null;

        // A `numId` that `numbering.xml` does not define numbers nothing in Word.
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
    /// Null when the file does not define it; saving checks the node's <c>numId</c> against it.
    /// </summary>
    internal JsonObject? FileDefinitionOf(int numId) =>
        InstanceOf(numId) is { } instance && AbstractOf(instance) is not null ? DefinitionOf(numId) : null;

    /// <summary>
    /// Missing levels follow the first level of the **definition**, not of the paragraph: otherwise
    /// saving would create a duplicate.
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

    /// <summary>Cloned per request: a <c>JsonNode</c> can only have one parent.</summary>
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
                // A level replaced in `w:num` is another list.
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

        // What saving needs to recreate the level in another file, only when it departs from the
        // default.
        if (level.LevelJustification?.Val?.InnerText is { } jc && jc is not ("left" or "start")) json["jc"] = jc;
        if (level.LevelSuffix?.Val?.InnerText is { } suff && suff != "tab") json["suff"] = suff;
        if (level.LevelRestart?.Val?.Value is { } restart) json["restart"] = restart;

        // Number formatting and level style do not fit: a save that needs to recreate them warns.
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
    /// A list tied to <c>w:numStyleLink</c> has no levels: they live in the definition with the
    /// same <c>w:styleLink</c>.
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

        // "none" is numbering turned off at that level.
        return format.Value == NumberFormatValues.Bullet || format.Value == NumberFormatValues.None
            ? "bulletList"
            : "orderedList";
    }

    /// <summary>
    /// Only on bullet lists; the ordered <c>%1.</c> belongs to <c>list-numbering.ts</c>.
    /// </summary>
    private static string? MarkerOf(Level? definition)
    {
        if (definition?.NumberingFormat?.Val?.Value != NumberFormatValues.Bullet) return null;

        var text = definition.LevelText?.Val?.Value;
        return string.IsNullOrEmpty(text) ? null : ListLevels.Shown(text);
    }

    /// <c>@left</c> is where the text starts; <c>@hanging</c>, how far before it the marker sits.
    private static double? TwipsOf(string? value)
    {
        if (value is null || !int.TryParse(value, out var twips) || twips <= 0) return null;
        return ListLevels.Mm(twips);
    }
}
