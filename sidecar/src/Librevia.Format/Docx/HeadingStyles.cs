using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Pelo <c>w:name</c>, que não se traduz (<c>heading 1</c>), e com o id que o documento
/// deu (<c>Ttulo1</c> num Word em português). Sem estilo de título, a definição vem
/// de <see cref="TemplateStyles"/>, e <c>word/styles.xml</c> passa a ser parte tocada.
/// </summary>
/// <param name="touched"><c>null</c> para só consultar, como na caixa de texto do cabeçalho.</param>
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

    /// <summary>O critério do leitor (<c>StyleResolver.HeadingLevelByName</c>), para o título rebaixado não voltar título.</summary>
    public int? LevelByName(string? styleId) => LevelOfName(Definition(styleId)?.StyleName?.Val?.Value);

    /// <summary>Um <c>w:pStyle</c> com id inexistente o Word desenha como Normal.</summary>
    public bool Defines(string? styleId) => Definition(styleId) is not null;

    private Style? Definition(string? styleId) =>
        styleId is null
            ? null
            : part.StyleDefinitionsPart?.Styles?.Elements<Style>()
                .FirstOrDefault(style =>
                    style.StyleId?.Value == styleId &&
                    (style.Type is null || style.Type.Value == StyleValues.Paragraph));

    /// <summary><c>heading 1</c> a <c>heading 6</c>, sem caixa: o LibreOffice grava <c>Heading 1</c>.</summary>
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
    /// O que o pacote não define é procurado pelo nome interno, e senão a definição
    /// embutida é copiada. Id que nem o embutido conhece volta como veio.
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

        // Um `Heading1` com outro nome é outro estilo: a cópia ganha um id livre.
        var id = style.StyleId!.Value!;
        for (var suffix = 2; ids.Contains(id); suffix++) id = $"{collisionPrefix}{suffix}";
        style.StyleId = id;
        // Dois `w:default` confundem o Word.
        style.Default = null;

        // Herdar de um `Normal` que não existe seria uma referência pendurada.
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
