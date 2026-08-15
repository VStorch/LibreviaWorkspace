using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Resolve a formatação efetiva de um parágrafo: padrões do documento, estilo
/// (com a cadeia de heranças) e formatação direta, nessa ordem.
/// </summary>
/// <remarks>
/// Sem isto, um documento que usa estilos abre praticamente sem formatação —
/// e é o caso do corpus real, onde quase todo parágrafo tem `w:pStyle`. O
/// exemplo que doeu: `Heading1` neste corpus não é um título grande, é uma
/// **barra vermelha com texto branco em Arial 10 pt**. Nada disso está no
/// parágrafo; está todo em `styles.xml`.
///
/// O editor não tem noção de estilo. Então resolvemos aqui e emitimos valores
/// concretos — é o que permite a tela mostrar o documento como ele é.
/// </remarks>
public sealed class StyleResolver
{
    /// <summary>Teto de segurança contra `basedOn` circular.</summary>
    private const int MaxChainDepth = 16;

    private readonly Dictionary<string, Style> _byId;
    private readonly ParagraphPropertiesBaseStyle? _defaultParagraph;
    private readonly RunPropertiesBaseStyle? _defaultRun;
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
    }

    /// <summary>Propriedades efetivas do parágrafo e dos seus runs.</summary>
    public (ParagraphProperties Paragraph, RunProperties Run) Resolve(ParagraphProperties? direct)
    {
        var styleId = direct?.ParagraphStyleId?.Val?.Value;
        var (paragraph, run) = FromStyle(styleId);

        var mergedParagraph = (ParagraphProperties)paragraph.CloneNode(true);
        var mergedRun = (RunProperties)run.CloneNode(true);

        if (direct is not null) Overlay(mergedParagraph, direct);

        // `w:rPr` dentro de `w:pPr` formata **a marca de parágrafo**, não os
        // runs — é a formatação que o Word usa para o texto digitado no fim da
        // linha. Aplicá-la aos runs põe negrito em parágrafos que não têm.
        return (mergedParagraph, mergedRun);
    }

    /// <summary>Propriedades efetivas de um run: estilo do parágrafo + diretas.</summary>
    public RunProperties ResolveRun(RunProperties inherited, RunProperties? direct)
    {
        var merged = (RunProperties)inherited.CloneNode(true);
        if (direct is not null) Overlay(merged, direct);
        return merged;
    }

    private (ParagraphProperties, RunProperties) FromStyle(string? styleId)
    {
        var key = styleId ?? string.Empty;
        if (_cache.TryGetValue(key, out var cached)) return cached;

        var paragraph = new ParagraphProperties();
        var run = new RunProperties();

        // Padrões do documento primeiro: é a base sobre a qual tudo se aplica.
        if (_defaultParagraph is not null) Overlay(paragraph, _defaultParagraph);
        if (_defaultRun is not null) Overlay(run, _defaultRun);

        // Do ancestral mais distante para o mais próximo, para que o mais
        // próximo tenha a última palavra.
        foreach (var style in ChainOf(styleId))
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
    /// Copia as propriedades de <paramref name="source"/> por cima de
    /// <paramref name="target"/>, substituindo as de mesmo nome.
    /// </summary>
    /// <remarks>
    /// Por nome de elemento, e não por propriedade tipada: assim não é preciso
    /// enumerar as dezenas de propriedades do OOXML, e o que ainda não sabemos
    /// ler atravessa junto sem esforço.
    /// </remarks>
    private static void Overlay(OpenXmlElement target, OpenXmlElement source)
    {
        foreach (var incoming in source.ChildElements)
        {
            if (incoming is ParagraphStyleId) continue;

            var existing = target.ChildElements
                .FirstOrDefault(child => child.LocalName == incoming.LocalName &&
                                         child.NamespaceUri == incoming.NamespaceUri);

            if (existing is not null) target.RemoveChild(existing);
            target.AppendChild(incoming.CloneNode(true));
        }
    }
}
