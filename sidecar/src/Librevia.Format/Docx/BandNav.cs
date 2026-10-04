using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Leitor e escritor percorrem os mesmos parágrafos na mesma ordem: é ela que dá o
/// endereço de cada peça. O ramo de reserva do <c>mc:AlternateContent</c> fica de fora,
/// porque repete o conteúdo; o espelhamento escreve nele depois.
/// </summary>
internal static class BandNav
{
    internal static List<Paragraph> ParagraphsOf(OpenXmlElement root) =>
        root.Descendants<Paragraph>()
            .Where(paragraph => !paragraph.Ancestors<AlternateContentFallback>().Any())
            .ToList();

    internal static Dictionary<Paragraph, int> IndexOf(OpenXmlElement root)
    {
        // Identidade, e não igualdade: dois parágrafos de mesmo texto são dois endereços.
        var index = new Dictionary<Paragraph, int>(
            (IEqualityComparer<Paragraph>)ReferenceEqualityComparer.Instance);

        var paragraphs = ParagraphsOf(root);
        for (var at = 0; at < paragraphs.Count; at++) index[paragraphs[at]] = at;
        return index;
    }

    /// <summary>
    /// Nulo quando a relação não existe **neste** pacote, como no <c>.sdoc</c> reaberto num
    /// pacote mínimo: perda declarada, e não um <c>ArgumentOutOfRangeException</c>.
    /// </summary>
    internal static (OpenXmlPart Owner, OpenXmlPartRootElement Root)? PartOf(
        MainDocumentPart part,
        string relationshipId)
    {
        if (string.IsNullOrEmpty(relationshipId)) return null;

        var owner = part.Parts
            .Where(pair => string.Equals(pair.RelationshipId, relationshipId, StringComparison.Ordinal))
            .Select(pair => pair.OpenXmlPart)
            .FirstOrDefault();

        return owner switch
        {
            HeaderPart header when header.Header is { } root => (header, root),
            FooterPart footer when footer.Footer is { } root => (footer, root),
            _ => null,
        };
    }

    /// <summary>Sem a barra da URI: é pelo nome do zip que a gravação decide o que devolver intacto.</summary>
    internal static string PathOf(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');

    /// <summary>A caixa inteira é regenerada quando muda; caixa dentro de caixa vai com a de fora.</summary>
    internal static List<TextBoxContent> BoxesOf(OpenXmlElement root) =>
        root.Descendants<TextBoxContent>()
            .Where(box => !box.Ancestors<AlternateContentFallback>().Any())
            .Where(box => !box.Ancestors<TextBoxContent>().Any())
            .ToList();

    internal static Dictionary<TextBoxContent, int> BoxIndexOf(OpenXmlElement root)
    {
        var index = new Dictionary<TextBoxContent, int>(
            (IEqualityComparer<TextBoxContent>)ReferenceEqualityComparer.Instance);

        var boxes = BoxesOf(root);
        for (var at = 0; at < boxes.Count; at++) index[boxes[at]] = at;
        return index;
    }

    /// <summary>Outro separador: confundir com o endereço de peça escreveria um parágrafo num <c>w:t</c>.</summary>
    internal static string BoxAddress(string relationshipId, int box) => $"{relationshipId}#{box}";

    internal static (string RelationshipId, int Box)? ParseBox(string? address)
    {
        if (string.IsNullOrEmpty(address)) return null;

        var parts = address.Split('#');
        if (parts.Length != 2 || parts[0].Length == 0) return null;
        if (!int.TryParse(parts[1], out var box) || box < 0) return null;

        return (parts[0], box);
    }

    /// <summary>A relação, o parágrafo e a peça nele.</summary>
    internal static string Address(string relationshipId, int paragraph, int piece) =>
        $"{relationshipId}:{paragraph}:{piece}";

    /// <summary>Nulo para qualquer coisa fora do formato.</summary>
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
