using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>
/// O que o documento tem e nós não damos conta — separado em categorias que
/// **não** são o mesmo problema.
/// </summary>
/// <remarks>
/// <b>Invisibilidade</b>: continua no arquivo depois de salvar, mas não aparece
/// na tela. O usuário precisa saber para não achar que sumiu.
///
/// <b>Perda</b>: some de verdade ao salvar. É o aviso grave, e a edição
/// cirúrgica torna esta lista curta — só entra o que estava dentro de um bloco
/// que o usuário editou.
///
/// <b>Estrutural</b>: um subconjunto da invisibilidade. São os recursos que
/// somem se — e só se — o usuário editar justamente o bloco que os ancora.
/// Comentário, revisão e nota de rodapé caem aqui; posicionamento de imagem e
/// decoração, não. A diferença decide se o documento abre em somente leitura,
/// que é a proteção mais forte contra perda de dados.
///
/// Misturar as categorias produz um aviso genérico que o usuário aprende a
/// ignorar em duas semanas, e aí ele deixa de proteger de qualquer coisa.
/// </remarks>
public sealed class Inventory
{
    /// <summary>
    /// Rótulos cujo desaparecimento é perda de conteúdo, não de aparência.
    /// </summary>
    /// <remarks>
    /// Constantes, e não literais espalhados pelos leitores, porque a
    /// classificação é feita **por rótulo**: uma frase mudada num leitor
    /// deixaria de casar com esta lista e o documento passaria a abrir editável
    /// sem que ninguém percebesse. Sendo constantes, o compilador não deixa.
    /// </remarks>
    public const string Comments = "comentários";
    public const string TrackedChanges = "controle de alterações";
    public const string Footnotes = "notas de rodapé";
    public const string Endnotes = "notas de fim";
    public const string Fields = "campos calculados (como sumário e número de página)";
    public const string HeaderFields = "campos calculados no cabeçalho";
    /// <summary>
    /// A **moldura** da forma, e não o conteúdo dela.
    /// </summary>
    /// <remarks>
    /// O texto de dentro e a posição são desenhados. A borda e o preenchimento
    /// também, quando são de cor sólida — ver <see cref="ShapeLook"/>. Sobra o
    /// que o CSS não faz por um retângulo: gradiente, textura, imagem de fundo,
    /// sombra, três dimensões, canto arredondado, e a forma que não declara
    /// preenchimento nem contorno e os herda de um tema que não resolvemos.
    ///
    /// É só disso que este aviso fala. Antes ele saía em toda forma, tivesse ela
    /// decoração ou não — e nos quatro documentos de evidências do corpus as
    /// caixas declaram `a:noFill` e linha de espessura zero, de modo que ele
    /// apontava para uma perda que não existia. Aviso que aparece sempre é aviso
    /// que se aprende a ignorar em duas semanas.
    ///
    /// Por isso este rótulo saiu da lista estrutural: desde que os objetos
    /// ancorados passaram a ser copiados do XML original para o parágrafo
    /// reescrito, editar o parágrafo não os apaga mais — e travar o documento
    /// inteiro deixou de proteger de coisa alguma.
    /// </remarks>
    public const string Shapes = "moldura e preenchimento de formas";
    public const string ContentControls = "controles de conteúdo";

    private static readonly HashSet<string> StructuralLabels = new(StringComparer.Ordinal)
    {
        Comments, TrackedChanges, Footnotes, Endnotes, Fields, HeaderFields, ContentControls,
    };

    private readonly SortedSet<string> _invisible = new(StringComparer.Ordinal);
    private readonly SortedSet<string> _lost = new(StringComparer.Ordinal);
    private readonly SortedSet<string> _structural = new(StringComparer.Ordinal);

    [JsonPropertyName("invisible")]
    public IReadOnlyCollection<string> Invisible => _invisible;

    [JsonPropertyName("lost")]
    public IReadOnlyCollection<string> Lost => _lost;

    /// <summary>Subconjunto de <see cref="Invisible"/>: o que se perde se o bloco for editado.</summary>
    [JsonPropertyName("structural")]
    public IReadOnlyCollection<string> Structural => _structural;

    /// <summary>Registra uma frase já escrita para o usuário.</summary>
    public void NoteInvisible(string message)
    {
        _invisible.Add(message);
        if (StructuralLabels.Contains(message)) _structural.Add(message);
    }

    public void NoteLoss(string message) => _lost.Add(message);

    /// <summary>
    /// Registra a partir do nome de um elemento OOXML, traduzindo — ou
    /// engolindo, quando não vale aviso.
    /// </summary>
    /// <remarks>
    /// Separado de <see cref="NoteInvisible"/> de propósito. Quando os dois
    /// eram o mesmo método, a regra "nome desconhecido em minúscula é ruído do
    /// formato" engolia também as frases em português — e o inventário voltava
    /// vazio para documentos que tinham o que avisar.
    /// </remarks>
    public void NoteInvisibleElement(string elementName)
    {
        var label = Describe(elementName);
        if (label is not null) NoteInvisible(label);
    }

    /// <summary>
    /// Traduz nome de elemento OOXML para algo que o usuário reconheça — ou
    /// devolve <c>null</c> para o que não merece aviso nenhum.
    /// </summary>
    /// <remarks>
    /// A lista de silêncio importa tanto quanto a de tradução. `w:bidi`,
    /// `w:textDirection` e `w:formProt` aparecem em toda seção gravada pelo
    /// LibreOffice — e o corpus inteiro foi gravado por ele. Avisar sobre eles
    /// encheria a tela de ruído em documentos perfeitamente normais.
    /// </remarks>
    private static string? Describe(string what) => what switch
    {
        // Ruído do LibreOffice e do próprio formato: nada a dizer ao usuário.
        "bidi" or "textDirection" or "formProt" or "docGrid" or "rPr" or "pPr" => null,
        "proofErr" or "lastRenderedPageBreak" or "bookmarkStart" or "bookmarkEnd" => null,
        "sectPr" or "tabs" or "spacing" or "ind" or "jc" or "widowControl" => null,

        "commentRangeStart" or "commentRangeEnd" or "comentário" => Comments,
        "ins" or "del" or "controle de alterações" => TrackedChanges,
        "footnoteReference" => Footnotes,
        "endnoteReference" => Endnotes,
        "fldChar" or "fldSimple" or "instrText" or "campo calculado" => Fields,
        "pict" or "object" or "AlternateContent" => Shapes,
        // Desenho e marcação inteligente aparecem, respectivamente, como imagem
        // e como texto comum: o conteúdo continua na tela, então não são
        // estruturais.
        "drawing" => "desenhos",
        "smartTag" => "marcações inteligentes",
        "sdt" or "sdtBlock" => ContentControls,

        _ => what.Length > 0 && char.IsLower(what[0]) ? null : what,
    };

    [JsonIgnore]
    public bool IsEmpty => _invisible.Count == 0 && _lost.Count == 0;
}
