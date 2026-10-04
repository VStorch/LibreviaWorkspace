using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>
/// <b>Invisível</b>: continua no arquivo, mas não aparece na tela. <b>Perda</b>:
/// some ao salvar, e só o que estava num bloco editado. <b>Estrutural</b>: o
/// invisível que some se o bloco que o ancora for editado; decide o somente
/// leitura. Misturadas, viram um aviso que se aprende a ignorar.
/// </summary>
/// <remarks>
/// Os rótulos são constantes porque a classificação é feita **por rótulo**: uma
/// frase mudada num leitor deixaria de casar e o documento abriria editável.
/// </remarks>
public sealed class Inventory
{
    /// <summary>O de cabeçalho, de nota ou de caixa de texto, que o editor não leva como nó.</summary>
    public const string Comments = "comentários";
    /// <summary>A revisão que o editor não representa: célula, numeração, seção ou tabela.</summary>
    public const string StructureRevisions = "revisões de estrutura";

    /// <summary><c>w:rPrChange</c> e <c>w:pPrChange</c>: no parágrafo editado, a do trecho se perde.</summary>
    public const string FormatRevisions = "revisões de formatação";
    /// <summary>A de dentro de uma caixa de texto e a do rascunho anterior às notas (<c>BeforeNotes</c>).</summary>
    public const string Footnotes = "notas de rodapé";
    public const string Endnotes = "notas de fim";
    public const string Fields = "campos calculados (como sumário e número de página)";
    public const string HeaderFields = "campos calculados no cabeçalho";
    /// <summary>
    /// A **moldura** que o CSS não faz: gradiente, textura, sombra, três dimensões,
    /// canto arredondado, ou a herdada de um tema que não resolvemos (ver
    /// <see cref="ShapeLook"/>).
    /// </summary>
    public const string Shapes = "moldura e preenchimento de formas";
    public const string ContentControls = "controles de conteúdo";
    /// <summary>A equação com construção que a tela não desenha: aparece travada, e o OMML volta inteiro.</summary>
    public const string Equations = "equações";

    private static readonly HashSet<string> StructuralLabels = new(StringComparer.Ordinal)
    {
        StructureRevisions, Fields, HeaderFields, ContentControls,
    };

    private readonly SortedSet<string> _invisible = new(StringComparer.Ordinal);
    private readonly SortedSet<string> _lost = new(StringComparer.Ordinal);
    private readonly SortedSet<string> _structural = new(StringComparer.Ordinal);

    [JsonPropertyName("invisible")]
    public IReadOnlyCollection<string> Invisible => _invisible;

    [JsonPropertyName("lost")]
    public IReadOnlyCollection<string> Lost => _lost;

    [JsonPropertyName("structural")]
    public IReadOnlyCollection<string> Structural => _structural;

    public void NoteInvisible(string message)
    {
        _invisible.Add(message);
        if (StructuralLabels.Contains(message)) _structural.Add(message);
    }

    public void NoteLoss(string message) => _lost.Add(message);

    /// <summary>
    /// Separado de <see cref="NoteInvisible"/>: a regra "nome desconhecido em
    /// minúscula é ruído" engoliria as frases em português.
    /// </summary>
    public void NoteInvisibleElement(string elementName)
    {
        var label = Describe(elementName);
        if (label is not null) NoteInvisible(label);
    }

    /// <summary>
    /// Ou <c>null</c> para o ruído: <c>w:bidi</c>, <c>w:textDirection</c> e <c>w:formProt</c>
    /// estão em toda seção gravada pelo LibreOffice.
    /// </summary>
    private static string? Describe(string what) => what switch
    {
        "bidi" or "textDirection" or "formProt" or "docGrid" or "rPr" or "pPr" => null,
        "proofErr" or "lastRenderedPageBreak" or "bookmarkStart" or "bookmarkEnd" => null,
        "sectPr" or "tabs" or "spacing" or "ind" or "jc" or "widowControl" => null,

        "commentRangeStart" or "commentRangeEnd" or "comentário" => Comments,
        "cellIns" or "cellDel" or "cellMerge" or "revisões de estrutura" => StructureRevisions,
        "footnoteReference" => Footnotes,
        "endnoteReference" => Endnotes,
        "fldChar" or "fldSimple" or "instrText" or "campo calculado" => Fields,
        "pict" or "object" or "AlternateContent" => Shapes,
        // Desenho e marcação inteligente aparecem como imagem e texto: não são estruturais.
        "drawing" => "desenhos",
        "smartTag" => "marcações inteligentes",
        "sdt" or "sdtBlock" => ContentControls,

        _ => what.Length > 0 && char.IsLower(what[0]) ? null : what,
    };

    [JsonIgnore]
    public bool IsEmpty => _invisible.Count == 0 && _lost.Count == 0;
}
