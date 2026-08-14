using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>
/// O que o documento tem e nós não damos conta — separado em duas categorias
/// que **não** são o mesmo problema.
/// </summary>
/// <remarks>
/// <b>Invisibilidade</b>: continua no arquivo depois de salvar, mas não aparece
/// na tela. O usuário precisa saber para não achar que sumiu.
///
/// <b>Perda</b>: some de verdade ao salvar. É o aviso grave, e a edição
/// cirúrgica torna esta lista curta — só entra o que estava dentro de um bloco
/// que o usuário editou.
///
/// Misturar os dois produz um aviso genérico que o usuário aprende a ignorar em
/// duas semanas, e aí ele deixa de proteger de qualquer coisa.
/// </remarks>
public sealed class Inventory
{
    private readonly SortedSet<string> _invisible = new(StringComparer.Ordinal);
    private readonly SortedSet<string> _lost = new(StringComparer.Ordinal);

    [JsonPropertyName("invisible")]
    public IReadOnlyCollection<string> Invisible => _invisible;

    [JsonPropertyName("lost")]
    public IReadOnlyCollection<string> Lost => _lost;

    /// <summary>Registra uma frase já escrita para o usuário.</summary>
    public void NoteInvisible(string message) => _invisible.Add(message);

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
        if (label is not null) _invisible.Add(label);
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

        "commentRangeStart" or "commentRangeEnd" or "comentário" => "comentários",
        "ins" or "del" or "controle de alterações" => "controle de alterações",
        "footnoteReference" => "notas de rodapé",
        "endnoteReference" => "notas de fim",
        "fldChar" or "fldSimple" or "instrText" or "campo calculado" => "campos calculados (como sumário e número de página)",
        "pict" or "object" or "AlternateContent" => "formas e caixas de texto",
        "drawing" => "desenhos",
        "smartTag" => "marcações inteligentes",
        "sdt" or "sdtBlock" => "controles de conteúdo",

        _ => what.Length > 0 && char.IsLower(what[0]) ? null : what,
    };

    [JsonIgnore]
    public bool IsEmpty => _invisible.Count == 0 && _lost.Count == 0;
}
