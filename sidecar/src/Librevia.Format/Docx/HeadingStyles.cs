using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// By <c>w:name</c>, which is not translated (<c>heading 1</c>), and with the id the document gave
/// (<c>Ttulo1</c> in Portuguese Word). Without a heading style, the definition comes from <see
/// cref="TemplateStyles"/>, and <c>word/styles.xml</c> becomes a touched part.
/// </summary>
/// <param name="touched"><c>null</c> to only query, as for the header text box.</param>
internal sealed class HeadingStyles(MainDocumentPart part, HashSet<string>? touched)
{
    private readonly Dictionary<int, string> _chosen = [];

    public string IdFor(int level)
    {
        if (_chosen.TryGetValue(level, out var known)) return known;

        var chosen = Declared(level) ?? Copied(level) ?? "Heading" + level;
        _chosen[level] = chosen;
        return chosen;
    }

    /// <summary>
    /// The reader's criterion (<c>StyleResolver.HeadingLevelByName</c>), so a demoted heading does
    /// not come back as a heading.
    /// </summary>
    public int? LevelByName(string? styleId) => LevelOfName(Definition(styleId)?.StyleName?.Val?.Value);

    /// <summary>Word draws a <c>w:pStyle</c> with a nonexistent id as Normal.</summary>
    public bool Defines(string? styleId) => Definition(styleId) is not null;

    private Style? Definition(string? styleId) =>
        styleId is null
            ? null
            : part.StyleDefinitionsPart?.Styles?.Elements<Style>()
                .FirstOrDefault(style =>
                    style.StyleId?.Value == styleId &&
                    (style.Type is null || style.Type.Value == StyleValues.Paragraph));

    /// <c>heading 1</c> to <c>heading 6</c>, case-insensitive: LibreOffice writes <c>Heading 1</c>.
    public static int? LevelOfName(string? name)
    {
        if (name is null) return null;

        const string Prefix = "heading ";
        if (!name.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase)) return null;

        return int.TryParse(name.AsSpan(Prefix.Length), out var level) && level >= 1 &&
               level <= TemplateStyles.HeadingLevels
            ? level
            : null;
    }

    private string? Declared(int level) =>
        part.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .FirstOrDefault(style =>
                style.Type?.Value == StyleValues.Paragraph && LevelOfName(style.StyleName?.Val?.Value) == level)
            ?.StyleId?.Value;

    private string? Copied(int level) => CopyOf(BuiltinStyles.Heading(level), $"Heading{level}_");

    /// <summary>
    /// What the package does not define is looked up by internal name, and otherwise the builtin
    /// definition is copied. An id even the builtin set does not know comes back as it came.
    /// </summary>
    public string IdForDeclared(string declared)
    {
        if (Defines(declared)) return declared;

        var builtin = BuiltinStyles.All.FirstOrDefault(style =>
            !style.Character && string.Equals(style.Id, declared, StringComparison.Ordinal));
        if (builtin is null) return declared;

        if (_copied.TryGetValue(declared, out var known)) return known;

        var byName = part.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .FirstOrDefault(style =>
                (style.Type is null || style.Type.Value == StyleValues.Paragraph) &&
                string.Equals(style.StyleName?.Val?.Value, builtin.Name, StringComparison.OrdinalIgnoreCase))
            ?.StyleId?.Value;

        var id = byName ?? CopyOf(builtin, declared + "_") ?? declared;
        _copied[declared] = id;
        return id;
    }

    private readonly Dictionary<string, string> _copied = new(StringComparer.Ordinal);

    private string? CopyOf(BuiltinStyle builtin, string collisionPrefix)
    {
        if (touched is null) return null;

        var definitions = part.StyleDefinitionsPart ?? part.AddNewPart<StyleDefinitionsPart>();
        var styles = definitions.Styles ??= new Styles();

        var ids = styles.Elements<Style>()
            .Select(style => style.StyleId?.Value)
            .OfType<string>()
            .ToHashSet(StringComparer.Ordinal);

        var style = TemplateStyles.Of(builtin);

        // A `Heading1` with another name is another style: the copy gets a free id.
        var id = style.StyleId!.Value!;
        for (var suffix = 2; ids.Contains(id); suffix++) id = $"{collisionPrefix}{suffix}";
        style.StyleId = id;
        // Two `w:default`s confuse Word.
        style.Default = null;

        // Inheriting from a `Normal` that does not exist would be a dangling reference.
        var normal = styles.Elements<Style>()
            .FirstOrDefault(candidate =>
                candidate.Type?.Value == StyleValues.Paragraph && candidate.Default?.Value == true)
            ?.StyleId?.Value;
        if (normal is null)
        {
            style.BasedOn = null;
            style.NextParagraphStyle = null;
        }
        else
        {
            style.BasedOn = new BasedOn { Val = normal };
            style.NextParagraphStyle = new NextParagraphStyle { Val = normal };
        }

        styles.AppendChild(style);
        styles.Save();
        touched.Add(definitions.Uri.OriginalString.TrimStart('/'));
        return id;
    }
}
