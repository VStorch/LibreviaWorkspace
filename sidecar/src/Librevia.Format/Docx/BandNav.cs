using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Reader and writer walk the same paragraphs in the same order: that order gives each piece its
/// address. The fallback branch of <c>mc:AlternateContent</c> is left out because it repeats the
/// content; mirroring writes into it afterwards.
/// </summary>
internal static class BandNav
{
    internal static List<Paragraph> ParagraphsOf(OpenXmlElement root) =>
        root.Descendants<Paragraph>()
            .Where(paragraph => !paragraph.Ancestors<AlternateContentFallback>().Any())
            .ToList();

    internal static Dictionary<Paragraph, int> IndexOf(OpenXmlElement root)
    {
        // Identity, not equality: two paragraphs with the same text are two addresses.
        var index = new Dictionary<Paragraph, int>(
            (IEqualityComparer<Paragraph>)ReferenceEqualityComparer.Instance);

        var paragraphs = ParagraphsOf(root);
        for (var at = 0; at < paragraphs.Count; at++) index[paragraphs[at]] = at;
        return index;
    }

    /// <summary>
    /// Null when the relationship does not exist in **this** package, as with a <c>.sdoc</c>
    /// reopened in a minimal package: a declared loss, not an <c>ArgumentOutOfRangeException</c>.
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

    /// <summary>
    /// Without the URI slash: saving decides by zip entry name what to return untouched.
    /// </summary>
    internal static string PathOf(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');

    /// <summary>
    /// A box is regenerated whole when it changes; a box inside a box goes with the outer one.
    /// </summary>
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

    /// <summary>
    /// A different separator: confusing it with a piece address would write a paragraph into a
    /// <c>w:t</c>.
    /// </summary>
    internal static string BoxAddress(string relationshipId, int box) => $"{relationshipId}#{box}";

    internal static (string RelationshipId, int Box)? ParseBox(string? address)
    {
        if (string.IsNullOrEmpty(address)) return null;

        var parts = address.Split('#');
        if (parts.Length != 2 || parts[0].Length == 0) return null;
        if (!int.TryParse(parts[1], out var box) || box < 0) return null;

        return (parts[0], box);
    }

    /// <summary>The relationship, the paragraph and the piece in it.</summary>
    internal static string Address(string relationshipId, int paragraph, int piece) =>
        $"{relationshipId}:{paragraph}:{piece}";

    /// <summary>Null for anything outside the format.</summary>
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
