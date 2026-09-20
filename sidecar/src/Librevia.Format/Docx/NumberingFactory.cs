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
/// Esta classe fecha a volta. Cria uma definição por tipo de lista e a reaproveita
/// no documento inteiro: quem tem três listas com marcador não precisa de três
/// definições iguais, e o Word as mostraria como três numerações diferentes na
/// caixa de diálogo.
///
/// `word/numbering.xml` é a única parte fora de `word/document.xml` que a
/// gravação passa a tocar — e só quando há lista nova. Ela entra na lista de
/// graváveis por fora, pelo mesmo mecanismo do cabeçalho em que se digitou: o
/// resto do arquivo continua voltando byte a byte.
/// </remarks>
internal sealed class NumberingFactory(MainDocumentPart part, HashSet<string> touched)
{
    /// <summary>Nove níveis, como o Word grava: é a profundidade que ele oferece.</summary>
    private const int Levels = 9;

    private const int IndentStep = 720;
    private const int Hanging = 360;

    private readonly Dictionary<string, int> _created = new(StringComparer.Ordinal);
    private Dictionary<int, int>? _declared;

    /// <summary>
    /// O `w:numId` a gravar para uma lista do modelo.
    /// </summary>
    /// <param name="kind">`bulletList` ou `orderedList`.</param>
    /// <param name="declared">O `numId` que o nó trouxe do arquivo, se trouxe.</param>
    public int NumberingIdFor(string kind, int? declared)
    {
        // O que veio do arquivo vale, desde que ainda exista: um `numId` que
        // ninguém define desenha lista sem marcador nenhum.
        if (declared is > 0 && Defined().ContainsKey(declared.Value)) return declared.Value;

        if (_created.TryGetValue(kind, out var existing)) return existing;

        var created = Create(kind);
        _created[kind] = created;
        return created;
    }

    private Dictionary<int, int> Defined() => _declared ??= Read();

    private Dictionary<int, int> Read()
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

    private int Create(string kind)
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

        var abstractId = numbering.Elements<AbstractNum>()
            .Select(existing => existing.AbstractNumberId?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

        var numberId = numbering.Elements<NumberingInstance>()
            .Select(existing => (int?)existing.NumberID?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

        var abstractNum = new AbstractNum { AbstractNumberId = abstractId };
        for (var level = 0; level < Levels; level++)
        {
            abstractNum.AppendChild(LevelOf(kind, level));
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

        numbering.AppendChild(new NumberingInstance(new AbstractNumId { Val = abstractId })
        {
            NumberID = numberId,
        });

        touched.Add(definitions.Uri.ToString().TrimStart('/'));
        if (_declared is not null) _declared[numberId] = abstractId;

        return numberId;
    }

    /// <summary>
    /// Um nível da definição, com a marca e o recuo que o Word usa por padrão.
    /// </summary>
    /// <remarks>
    /// A mesma marca em todos os níveis, e não as três que o Word alterna: o
    /// editor desenha uma só, e o arquivo deve sair parecido com o que a pessoa
    /// viu ao criar a lista. O recuo cresce meia polegada por nível, com o
    /// marcador pendurado a um quarto: as medidas do próprio Word.
    /// </remarks>
    private static Level LevelOf(string kind, int level)
    {
        var bullet = string.Equals(kind, "bulletList", StringComparison.Ordinal);

        var definition = new Level
        {
            LevelIndex = level,
            NumberingFormat = new NumberingFormat
            {
                Val = bullet ? NumberFormatValues.Bullet : NumberFormatValues.Decimal,
            },
            // `` é o ponto da fonte Symbol: o caractere que o Word grava
            // para "marcador redondo", e que o leitor traduz de volta para `•`.
            LevelText = new LevelText { Val = bullet ? "" : $"%{level + 1}." },
            LevelJustification = new LevelJustification { Val = LevelJustificationValues.Left },
            StartNumberingValue = new StartNumberingValue { Val = 1 },
            PreviousParagraphProperties = new PreviousParagraphProperties(new Indentation
            {
                Left = ((level + 1) * IndentStep).ToString(System.Globalization.CultureInfo.InvariantCulture),
                Hanging = Hanging.ToString(System.Globalization.CultureInfo.InvariantCulture),
            }),
        };

        if (bullet)
        {
            // A marca do Word vem da área de uso privado do Unicode, e só a fonte
            // Symbol a desenha. Sem declarar a fonte aparece a caixinha de
            // caractere ausente.
            definition.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(new RunFonts
            {
                Ascii = "Symbol",
                HighAnsi = "Symbol",
                Hint = FontTypeHintValues.Default,
            });
        }

        return definition;
    }
}
