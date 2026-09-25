using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A numeração de uma lista que o documento ainda não tem.
/// </summary>
/// <remarks>
/// No OOXML uma lista é um `w:numId` que aponta uma definição em
/// `word/numbering.xml`. O caminho de volta — lista feita aqui dentro — não
/// existia: o escritor gravava `w:numId w:val="0"`, que no formato quer dizer
/// **sem numeração**. A pessoa criava uma lista, salvava, reabria e encontrava
/// parágrafos comuns, sem marcador, sem recuo e sem aviso nenhum.
///
/// Esta classe fecha a volta. Cada lista nova ganha a **sua** definição: no Word
/// quem conta é a definição abstrata, e duas listas numeradas que dividissem uma
/// continuariam a contagem uma da outra ao reabrir — a tela mostrava 1, 2 e 1, 2,
/// e o Word, 1, 2 e 3, 4. A lista que traz definição (a reiniciada, a colada de
/// outro documento, a escolhida na galeria) nasce dela; a que não traz, dos
/// níveis padrão do Word (`ListLevels.Defaults`).
///
/// `word/numbering.xml` é a única parte fora de `word/document.xml` que a
/// gravação passa a tocar — e só quando há lista nova. Ela entra na lista de
/// graváveis por fora, pelo mesmo mecanismo do cabeçalho em que se digitou: o
/// resto do arquivo continua voltando byte a byte.
/// </remarks>
internal sealed class NumberingFactory(MainDocumentPart part, HashSet<string> touched)
{
    /// <summary>Numeração já criada nesta gravação, pela chave da definição.</summary>
    /// <remarks>
    /// Duas listas com a mesma chave são a mesma contagem — a segunda "continua a
    /// numeração" da primeira —, e têm de sair com o mesmo `w:num`.
    /// </remarks>
    private readonly Dictionary<string, int> _created = new(StringComparer.Ordinal);
    private HashSet<int>? _declared;

    /// <summary>
    /// O `w:numId` a gravar para uma lista do modelo.
    /// </summary>
    /// <param name="kind">`bulletList` ou `orderedList`.</param>
    /// <param name="declared">O `numId` que o nó trouxe do arquivo, ou herdou da lista de fora.</param>
    /// <param name="definition">A definição que o nó traz (`numbering`), se traz.</param>
    public int NumberingIdFor(string kind, int? declared, JsonObject? definition = null)
    {
        // O que veio do arquivo vale, desde que ainda exista: um `numId` que
        // ninguém define desenha lista sem marcador nenhum.
        if (declared is > 0 && Defined().Contains(declared.Value)) return declared.Value;

        var key = definition?["key"]?.GetValue<string>();
        if (key is not null && _created.TryGetValue(key, out var existing)) return existing;

        var created = Create(kind, definition);
        if (key is not null) _created[key] = created;
        return created;
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
            // Documento que nunca teve lista não tem a parte. Criá-la também
            // acrescenta a relação e o tipo de conteúdo, que já são graváveis.
            definitions = part.AddNewPart<NumberingDefinitionsPart>();
            definitions.Numbering = new Numbering();
        }

        var numbering = definitions.Numbering ??= new Numbering();
        var levels = definition?["levels"] as JsonArray;

        var abstractId = Reusable(numbering, kind, definition, levels) ?? AddAbstract(numbering, kind, levels);

        var numberId = numbering.Elements<NumberingInstance>()
            .Select(existing => (int?)existing.NumberID?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

        var instance = new NumberingInstance(new AbstractNumId { Val = abstractId }) { NumberID = numberId };

        // O reinício ("Reiniciar em 1", "Definir valor inicial") é do `w:num`, e
        // não da definição: é assim que o Word o grava, e é o que o faz contar à
        // parte sem mudar a lista de onde a reiniciada saiu.
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

    /// <summary>
    /// A definição abstrata do arquivo que a lista pode reaproveitar.
    /// </summary>
    /// <remarks>
    /// Só quando é **a mesma**: a lista reiniciada aponta a definição de onde
    /// saiu. Um `abstractId` que veio colado de outro documento pode existir aqui
    /// com outra marca, e reaproveitá-lo trocaria a numeração da lista em silêncio
    /// — daí a comparação dos níveis, e não só do número.
    /// </remarks>
    private static int? Reusable(Numbering numbering, string kind, JsonObject? definition, JsonArray? levels)
    {
        if (levels is null || definition?["abstractId"] is not JsonValue value ||
            !value.TryGetValue<int>(out var abstractId))
        {
            return null;
        }

        var found = numbering.Elements<AbstractNum>()
            .FirstOrDefault(candidate => candidate.AbstractNumberId?.Value == abstractId);
        if (found is null) return null;

        var existing = ListLevels.Defaults(kind);
        foreach (var level in found.Elements<Level>())
        {
            var index = level.LevelIndex?.Value ?? 0;
            if (index is >= 0 and < ListLevels.Count) existing[index] = NumberingReader.LevelJson(level);
        }

        return JsonNode.DeepEquals(existing, levels) ? abstractId : null;
    }

    private static int AddAbstract(Numbering numbering, string kind, JsonArray? levels)
    {
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
            var source = levels is not null && level < levels.Count ? levels[level] as JsonObject : null;
            abstractNum.AppendChild(ListLevels.ToOpenXml(source, level, kind));
        }

        // A ordem do `w:numbering` é sequência, não conjunto: primeiro os
        // `w:abstractNum`, depois os `w:num`. Trocá-los invalida a parte.
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
