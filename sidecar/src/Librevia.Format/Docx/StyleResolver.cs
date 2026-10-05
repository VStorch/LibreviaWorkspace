using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Document defaults, style (with its chain) and direct formatting, in that order: in the corpus
/// almost every paragraph has <c>w:pStyle</c>, and <c>Heading1</c> is a red bar that lives in
/// <c>styles.xml</c>. This is the flattened reading (lists, cells, old drafts).
/// </summary>
public sealed class StyleResolver
{
    /// <summary>Against a circular <c>basedOn</c>.</summary>
    private const int MaxChainDepth = 16;

    private readonly Dictionary<string, Style> _byId;
    private readonly ParagraphPropertiesBaseStyle? _defaultParagraph;
    private readonly RunPropertiesBaseStyle? _defaultRun;

    /// <c>w:default="1"</c>, for a paragraph without <c>w:pStyle</c>.
    private readonly string? _defaultParagraphStyleId;
    private readonly Dictionary<string, (ParagraphProperties P, RunProperties R)> _cache = new(StringComparer.Ordinal);

    public StyleResolver(MainDocumentPart part)
    {
        var styles = part.StyleDefinitionsPart?.Styles;
        _byId = styles?.Elements<Style>()
            .Where(style => style.StyleId?.Value is not null)
            .GroupBy(style => style.StyleId!.Value!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal)
            ?? new Dictionary<string, Style>(StringComparer.Ordinal);

        _defaultParagraph = styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
        _defaultRun = styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;

        _defaultParagraphStyleId = _byId.Values
            .FirstOrDefault(style =>
                style.Type?.Value == StyleValues.Paragraph && style.Default?.Value == true)
            ?.StyleId?.Value;
    }

    /// <summary>
    /// The id is translated and the name is not (<c>Überschrift1</c> is <c>heading 1</c>): the same
    /// criterion as <see cref="HeadingStyles"/>.
    /// </summary>
    public int? HeadingLevelByName(string? styleId) =>
        styleId is not null && _byId.TryGetValue(styleId, out var style)
            ? HeadingStyles.LevelOfName(style.StyleName?.Val?.Value)
            : null;

    public (ParagraphProperties Paragraph, RunProperties Run) Resolve(ParagraphProperties? direct)
    {
        var styleId = direct?.ParagraphStyleId?.Val?.Value;
        var (paragraph, run) = FromStyle(styleId);

        var mergedParagraph = (ParagraphProperties)paragraph.CloneNode(true);
        var mergedRun = (RunProperties)run.CloneNode(true);

        if (direct is not null) Overlay(mergedParagraph, direct);

        // `w:pPr/w:rPr` formats the paragraph mark, not the runs.
        return (mergedParagraph, mergedRun);
    }

    /// <summary>
    /// Without direct formatting: so the reader knows whether a direct silence undoes something
    /// from the style.
    /// </summary>
    public ParagraphProperties StyleParagraphOf(ParagraphProperties? direct) =>
        FromStyle(direct?.ParagraphStyleId?.Val?.Value).Item1;

    public RunProperties ResolveRun(RunProperties inherited, RunProperties? direct)
    {
        var merged = (RunProperties)inherited.CloneNode(true);
        if (direct is not null) Overlay(merged, direct);
        return merged;
    }

    /// <summary>
    /// The font Word measures the line with and gives an empty paragraph its height.
    /// </summary>
    public RunProperties ResolveMark(RunProperties inherited, ParagraphProperties? direct)
    {
        var merged = (RunProperties)inherited.CloneNode(true);
        if (direct?.ParagraphMarkRunProperties is { } mark) Overlay(merged, mark);
        return merged;
    }

    private (ParagraphProperties, RunProperties) FromStyle(string? styleId)
    {
        var key = styleId ?? string.Empty;
        if (_cache.TryGetValue(key, out var cached)) return cached;

        var paragraph = new ParagraphProperties();
        var run = new RunProperties();

        if (_defaultParagraph is not null) Overlay(paragraph, _defaultParagraph);
        if (_defaultRun is not null) Overlay(run, _defaultRun);

        // From the farthest ancestor to the nearest, which has the last word.
        foreach (var style in ChainOf(styleId ?? _defaultParagraphStyleId))
        {
            if (style.StyleParagraphProperties is not null) Overlay(paragraph, style.StyleParagraphProperties);
            if (style.StyleRunProperties is not null) Overlay(run, style.StyleRunProperties);
        }

        var resolved = (paragraph, run);
        _cache[key] = resolved;
        return resolved;
    }

    private List<Style> ChainOf(string? styleId)
    {
        var chain = new List<Style>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;

        while (!string.IsNullOrEmpty(current) && seen.Add(current) && chain.Count < MaxChainDepth)
        {
            if (!_byId.TryGetValue(current, out var style)) break;
            chain.Add(style);
            current = style.BasedOn?.Val?.Value;
        }

        chain.Reverse();
        return chain;
    }

    /// <summary>
    /// The OOXML rule is to replace the whole property; these four merge attribute by attribute: a
    /// paragraph declaring only <c>w:after="0"</c> does not erase the style's line spacing.
    /// </summary>
    private static readonly HashSet<string> AttributeByAttribute = new(StringComparer.Ordinal)
    {
        "spacing",
        "ind",
        "rFonts",
        "lang",
    };

    /// <summary>By element name: what we cannot read yet passes through.</summary>
    private static void Overlay(OpenXmlElement target, OpenXmlElement source)
    {
        foreach (var incoming in source.ChildElements)
        {
            if (incoming is ParagraphStyleId) continue;

            var existing = target.ChildElements
                .FirstOrDefault(child => child.LocalName == incoming.LocalName &&
                                         child.NamespaceUri == incoming.NamespaceUri);

            if (existing is not null && AttributeByAttribute.Contains(incoming.LocalName))
            {
                foreach (var attribute in incoming.GetAttributes()) existing.SetAttribute(attribute);
                continue;
            }

            if (existing is not null) target.RemoveChild(existing);
            target.AppendChild(incoming.CloneNode(true));
        }
    }
}
