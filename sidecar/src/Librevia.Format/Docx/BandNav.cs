using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A travessia dos parágrafos de um cabeçalho ou rodapé, uma só.
/// </summary>
/// <remarks>
/// Leitor e escritor precisam ver os mesmos parágrafos na mesma ordem: é essa
/// ordem que dá o endereço de cada peça editável, e dois percursos parecidos
/// escritos em dois arquivos é como o texto digitado numa linha acaba noutra.
/// Mesmo remédio de <see cref="TextBoxNav"/>, um andar acima.
///
/// O ramo de reserva do `mc:AlternateContent` fica de fora porque repete o
/// mesmo conteúdo: contá-lo dobraria os índices e desalinharia tudo depois
/// dele. Quem escreve nele é o espelhamento, depois, a partir do ramo que vale.
/// </remarks>
internal static class BandNav
{
    /// <summary>Os parágrafos da parte, na ordem que dá o endereço.</summary>
    internal static List<Paragraph> ParagraphsOf(OpenXmlElement root) =>
        root.Descendants<Paragraph>()
            .Where(paragraph => !paragraph.Ancestors<AlternateContentFallback>().Any())
            .ToList();

    /// <summary>O endereço de cada parágrafo, para o leitor carimbar as peças.</summary>
    internal static Dictionary<Paragraph, int> IndexOf(OpenXmlElement root)
    {
        // Identidade, e não igualdade: dois parágrafos com o mesmo texto são
        // dois endereços diferentes, e é o objeto que o leitor tem em mãos.
        var index = new Dictionary<Paragraph, int>(
            (IEqualityComparer<Paragraph>)ReferenceEqualityComparer.Instance);

        var paragraphs = ParagraphsOf(root);
        for (var at = 0; at < paragraphs.Count; at++) index[paragraphs[at]] = at;
        return index;
    }

    /// <summary>
    /// A parte apontada por uma relação, e a raiz dela.
    /// </summary>
    /// <remarks>
    /// Cabeçalho e rodapé são partes irmãs com o mesmo formato por dentro; quem
    /// escreve o texto de volta não precisa saber qual das duas está olhando.
    /// </remarks>
    internal static (OpenXmlPart Owner, OpenXmlPartRootElement Root)? PartOf(
        MainDocumentPart part,
        string relationshipId)
    {
        if (string.IsNullOrEmpty(relationshipId)) return null;

        return part.GetPartById(relationshipId) switch
        {
            HeaderPart header when header.Header is { } root => (header, root),
            FooterPart footer when footer.Footer is { } root => (footer, root),
            _ => null,
        };
    }

    /// <summary>
    /// O caminho da parte dentro do pacote, como o zip o nomeia.
    /// </summary>
    /// <remarks>
    /// A URI da parte vem com barra na frente — `/word/header1.xml` — e a
    /// entrada do zip não. É por este nome que a gravação decide o que devolver
    /// intacto, então errar a barra devolveria o cabeçalho antigo por cima do
    /// texto recém-digitado, sem erro nenhum.
    /// </remarks>
    internal static string PathOf(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');

    /// <summary>
    /// As caixas de texto da parte, na ordem que dá o endereço.
    /// </summary>
    /// <remarks>
    /// O cabeçalho corporativo do corpus não é feito de parágrafos: é um grupo
    /// de formas, e o título mora dentro de uma caixa. Ela não cabe na conta de
    /// peças — a caixa inteira é regenerada quando o texto muda, porque digitar
    /// pode abrir e fechar parágrafos dentro dela.
    ///
    /// Caixa dentro de caixa fica de fora: o conteúdo da de dentro já vai junto
    /// com o da de fora, e contá-la duas vezes desalinharia os endereços.
    /// </remarks>
    internal static List<TextBoxContent> BoxesOf(OpenXmlElement root) =>
        root.Descendants<TextBoxContent>()
            .Where(box => !box.Ancestors<AlternateContentFallback>().Any())
            .Where(box => !box.Ancestors<TextBoxContent>().Any())
            .ToList();

    /// <summary>O endereço de cada caixa, para o leitor carimbar os objetos.</summary>
    internal static Dictionary<TextBoxContent, int> BoxIndexOf(OpenXmlElement root)
    {
        var index = new Dictionary<TextBoxContent, int>(
            (IEqualityComparer<TextBoxContent>)ReferenceEqualityComparer.Instance);

        var boxes = BoxesOf(root);
        for (var at = 0; at < boxes.Count; at++) index[boxes[at]] = at;
        return index;
    }

    /// <summary>
    /// O endereço de uma caixa: a relação e a caixa dentro dela.
    /// </summary>
    /// <remarks>
    /// Separador diferente do endereço de peça de propósito: são duas coisas
    /// que não se substituem, e confundi-las escreveria um parágrafo inteiro
    /// dentro de um `w:t`.
    /// </remarks>
    internal static string BoxAddress(string relationshipId, int box) => $"{relationshipId}#{box}";

    /// <summary>Desmonta o endereço de uma caixa.</summary>
    internal static (string RelationshipId, int Box)? ParseBox(string? address)
    {
        if (string.IsNullOrEmpty(address)) return null;

        var parts = address.Split('#');
        if (parts.Length != 2 || parts[0].Length == 0) return null;
        if (!int.TryParse(parts[1], out var box) || box < 0) return null;

        return (parts[0], box);
    }

    /// <summary>O endereço de uma peça: a relação, o parágrafo e a peça nele.</summary>
    internal static string Address(string relationshipId, int paragraph, int piece) =>
        $"{relationshipId}:{paragraph}:{piece}";

    /// <summary>Desmonta o endereço. Devolve `null` para qualquer coisa fora do formato.</summary>
    internal static (string RelationshipId, int Paragraph, int Piece)? Parse(string? address)
    {
        if (string.IsNullOrEmpty(address)) return null;

        var parts = address.Split(':');
        if (parts.Length != 3) return null;
        if (!int.TryParse(parts[1], out var paragraph) || paragraph < 0) return null;
        if (!int.TryParse(parts[2], out var piece) || piece < 0) return null;
        if (parts[0].Length == 0) return null;

        return (parts[0], paragraph, piece);
    }
}
