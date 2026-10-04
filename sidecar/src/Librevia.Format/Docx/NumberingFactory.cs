using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A lista feita aqui ganha a **sua** definição: no Word conta a abstrata, e duas
/// listas que a dividissem continuariam a contagem uma da outra. A que traz
/// definição (reiniciada, colada, da galeria) nasce dela; a outra, dos níveis
/// padrão (<c>ListLevels.Defaults</c>). <c>word/numbering.xml</c> só é tocado quando há
/// lista nova.
/// </summary>
internal sealed class NumberingFactory(MainDocumentPart part, HashSet<string> touched, Inventory? inventory = null)
{
    private readonly NumberingReader _reader = new(part);

    /// <summary>Pela chave: duas listas com a mesma chave são a mesma contagem e o mesmo <c>w:num</c>.</summary>
    private readonly Dictionary<string, int> _created = new(StringComparer.Ordinal);
    private HashSet<int>? _declared;

    /// <param name="kind"><c>bulletList</c> ou <c>orderedList</c>.</param>
    /// <param name="declared">O do arquivo, ou herdado da lista de fora.</param>
    /// <param name="definition">A que o nó traz (<c>numbering</c>).</param>
    public int NumberingIdFor(string kind, int? declared, JsonObject? definition = null)
    {
        var key = definition?["key"]?.GetValue<string>();

        // O do arquivo vale se ainda existe e é a **mesma** numeração: a lista colada
        // traz um `numId` que aqui pode ser outro.
        if (declared is > 0 && Defined().Contains(declared.Value) && SameAsFile(declared.Value, definition))
        {
            return declared.Value;
        }

        if (key is not null && _created.TryGetValue(key, out var existing)) return existing;

        var created = Create(kind, definition);
        if (key is not null) _created[key] = created;
        return created;
    }

    private bool SameAsFile(int numId, JsonObject? definition)
    {
        if (definition is null) return true;
        if (_reader.FileDefinitionOf(numId) is not { } file) return false;
        return JsonNode.DeepEquals(file["levels"], definition["levels"])
               && JsonNode.DeepEquals(file["overrides"], definition["overrides"]);
    }

    private HashSet<int> Defined() => _declared ??= [
        .. part.NumberingDefinitionsPart?.Numbering?.Elements<NumberingInstance>()
            .Select(instance => instance.NumberID?.Value ?? 0)
            .Where(id => id > 0) ?? [],
    ];

    private int Create(string kind, JsonObject? definition)
    {
        var definitions = part.NumberingDefinitionsPart;
        if (definitions is null)
        {
            // Documento sem lista não tem a parte.
            definitions = part.AddNewPart<NumberingDefinitionsPart>();
            definitions.Numbering = new Numbering();
        }

        var numbering = definitions.Numbering ??= new Numbering();
        var levels = definition?["levels"] as JsonArray;

        var abstractId = Reusable(numbering, definition, levels) ?? AddAbstract(numbering, kind, definition, levels);

        var numberId = numbering.Elements<NumberingInstance>()
            .Select(existing => (int?)existing.NumberID?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

        var instance = new NumberingInstance(new AbstractNumId { Val = abstractId }) { NumberID = numberId };

        // O reinício é do `w:num`, como o Word grava.
        if (definition?["overrides"] is JsonObject overrides)
        {
            foreach (var (name, value) in overrides.OrderBy(entry => entry.Key, StringComparer.Ordinal))
            {
                if (!int.TryParse(name, out var index) || index is < 0 or >= ListLevels.Count) continue;
                if (value is not JsonValue number || !number.TryGetValue<int>(out var start)) continue;

                instance.AppendChild(new LevelOverride(new StartOverrideNumberingValue { Val = start })
                {
                    LevelIndex = index,
                });
            }
        }

        numbering.AppendChild(instance);

        touched.Add(definitions.Uri.ToString().TrimStart('/'));
        Defined().Add(numberId);

        return numberId;
    }

    /// <summary>Só quando é **a mesma**, pelos níveis: o <c>abstractId</c> colado pode ser outro aqui.</summary>
    private static int? Reusable(Numbering numbering, JsonObject? definition, JsonArray? levels)
    {
        if (levels is null || SourceOf(numbering, definition) is not { } found) return null;
        return JsonNode.DeepEquals(NumberingReader.LevelsOf(found), levels) ? found.AbstractNumberId!.Value : null;
    }

    private static AbstractNum? SourceOf(Numbering numbering, JsonObject? definition)
    {
        if (definition?["abstractId"] is not JsonValue value || !value.TryGetValue<int>(out var abstractId)) return null;
        return numbering.Elements<AbstractNum>()
            .FirstOrDefault(candidate => candidate.AbstractNumberId?.Value == abstractId);
    }

    private int AddAbstract(Numbering numbering, string kind, JsonObject? definition, JsonArray? levels)
    {
        // Os níveis que não mudaram vêm do original, com fonte, cor e estilo do número.
        var source = SourceOf(numbering, definition);
        var original = source?.Elements<Level>()
            .Where(level => level.LevelIndex?.Value is >= 0 and < ListLevels.Count)
            .GroupBy(level => level.LevelIndex!.Value)
            .ToDictionary(group => group.Key, group => group.First());

        var abstractId = numbering.Elements<AbstractNum>()
            .Select(existing => existing.AbstractNumberId?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

        var abstractNum = new AbstractNum(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel })
        {
            AbstractNumberId = abstractId,
        };
        for (var level = 0; level < ListLevels.Count; level++)
        {
            var wanted = levels is not null && level < levels.Count ? levels[level] as JsonObject : null;
            if (wanted is not null && original is not null && original.TryGetValue(level, out var kept) &&
                JsonNode.DeepEquals(NumberingReader.LevelJson(kept), wanted))
            {
                abstractNum.AppendChild(kept.CloneNode(true));
                continue;
            }

            abstractNum.AppendChild(ListLevels.ToOpenXml(wanted, level, kind, inventory));
        }

        // Sequência: os `w:abstractNum` antes dos `w:num`.
        var lastAbstract = numbering.Elements<AbstractNum>().LastOrDefault();
        if (lastAbstract is not null) numbering.InsertAfter(abstractNum, lastAbstract);
        else
        {
            var lastPicture = numbering.Elements<NumberingPictureBullet>().LastOrDefault();
            if (lastPicture is not null) numbering.InsertAfter(abstractNum, lastPicture);
            else numbering.PrependChild(abstractNum);
        }

        return abstractId;
    }
}
