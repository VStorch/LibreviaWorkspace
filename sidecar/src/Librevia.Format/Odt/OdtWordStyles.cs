namespace Librevia.Format.Odt;

/// <summary>
/// LibreOffice's built-in styles and Word's: the same style under each program's id, so a heading
/// opened from ODT is the <c>Heading1</c> Word and the editor know. The name is Word's internal
/// English one (<c>heading 1</c>), which the editor translates.
/// </summary>
internal static class OdtWordStyles
{
    private sealed record Correspondence(string Family, string OdfName, string WordId, string WordName);

    private const int HeadingLevels = 9;

    private static readonly Correspondence[] Correspondences =
    [
        new(OdtStyles.ParagraphFamily, OdtStyles.DefaultParagraphStyle, "Normal", "Normal"),
        .. Enumerable.Range(1, HeadingLevels).Select(level =>
            new Correspondence(OdtStyles.ParagraphFamily, $"Heading_20_{level}", $"Heading{level}", $"heading {level}")),
        new(OdtStyles.ParagraphFamily, "Title", "Title", "Title"),
        new(OdtStyles.ParagraphFamily, "Subtitle", "Subtitle", "Subtitle"),
        new(OdtStyles.ParagraphFamily, "Text_20_body", "BodyText", "Body Text"),
        new(OdtStyles.ParagraphFamily, "Quotations", "Quote", "Quote"),
        new(OdtStyles.ParagraphFamily, "List_20_Paragraph", "ListParagraph", "List Paragraph"),
        new(OdtStyles.TextFamily, "Strong_20_Emphasis", "Strong", "Strong"),
        new(OdtStyles.TextFamily, "Emphasis", "Emphasis", "Emphasis"),
    ];

    private static readonly Dictionary<(string Family, string OdfName), Correspondence> ByOdfName =
        Correspondences.ToDictionary(entry => (entry.Family, entry.OdfName));

    private static readonly Dictionary<(string Family, string WordId), Correspondence> ByWordId =
        Correspondences.ToDictionary(entry => (entry.Family, entry.WordId));

    /// <summary>The inverse of <see cref="IdOf"/>: the file's name for the editor's id.</summary>
    public static string OdfNameOf(string family, string id) =>
        ByWordId.TryGetValue((family, id), out var entry) ? entry.OdfName : id;

    /// <summary>Word's id for a LibreOffice built-in; any other style keeps its own name.</summary>
    public static string IdOf(string family, string odfName) =>
        ByOdfName.TryGetValue((family, odfName), out var entry) ? entry.WordId : odfName;

    /// <summary>Word's name for a LibreOffice built-in, or <c>null</c> for a style of the file's own.</summary>
    public static string? NameOf(string family, string odfName) =>
        ByOdfName.TryGetValue((family, odfName), out var entry) ? entry.WordName : null;
}
