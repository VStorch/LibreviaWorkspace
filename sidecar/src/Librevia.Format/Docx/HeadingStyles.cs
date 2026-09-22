using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O estilo de título que **este** documento define para cada nível.
/// </summary>
/// <remarks>
/// O escritor gravava `"Heading" + nível`, sem conferir se o estilo existia. Num
/// Word em português o id é `Ttulo1` — o id é traduzido, e sem acento —, e o
/// título criado aqui dentro apontava um estilo que o documento não define: no
/// Word ele saía com a cara do Normal.
///
/// O que não se traduz é o **nome**: `heading 1` em qualquer idioma, porque é o
/// nome interno que o Word usa para reconhecer os títulos embutidos. Então o
/// estilo é escolhido pelo `w:name`, e o id é o que o documento tiver dado a ele.
///
/// Quando o documento não tem estilo de título nenhum, a definição vem do modelo
/// de documento novo (<see cref="TemplateStyles"/>) e `word/styles.xml` passa a
/// ser parte tocada — sem isso a restauração byte a byte da gravação cirúrgica
/// desfaria a cópia, e o título apontaria o vazio.
/// </remarks>
/// <param name="touched">
/// As partes que a gravação pode alterar, ou <c>null</c> para só consultar —
/// quem grava dentro de uma caixa de texto do cabeçalho não copia estilo nenhum.
/// </param>
internal sealed class HeadingStyles(MainDocumentPart part, HashSet<string>? touched)
{
    private readonly Dictionary<int, string> _chosen = [];

    /// <summary>O id do estilo a gravar no `w:pStyle` de um título.</summary>
    public string IdFor(int level)
    {
        if (_chosen.TryGetValue(level, out var known)) return known;

        var chosen = Declared(level) ?? Copied(level) ?? "Heading" + level;
        _chosen[level] = chosen;
        return chosen;
    }

    /// <summary>
    /// O nível de título que **este** pacote dá ao estilo, pelo nome dele.
    /// </summary>
    /// <remarks>
    /// O mesmo critério do leitor (<c>StyleResolver.HeadingLevelByName</c>): um
    /// `Überschrift1` de nome `heading 1` é título para os dois lados. Com dois
    /// critérios diferentes, o título que a pessoa rebaixa a parágrafo na tela
    /// continuaria título no arquivo e voltaria título ao reabrir.
    /// </remarks>
    public int? LevelByName(string? styleId) => LevelOfName(Definition(styleId)?.StyleName?.Val?.Value);

    /// <summary>O pacote define um estilo de parágrafo com este id?</summary>
    /// <remarks>
    /// Um `w:pStyle` que aponta id inexistente é silêncio: o Word desenha o
    /// parágrafo como Normal. Por isso o id que o modelo carrega só vale quando
    /// este pacote o define — o `.sdoc` que veio de um `.docx` em português traz
    /// `Ttulo1`, e o pacote mínimo não tem esse estilo.
    /// </remarks>
    public bool Defines(string? styleId) => Definition(styleId) is not null;

    private Style? Definition(string? styleId) =>
        styleId is null
            ? null
            : part.StyleDefinitionsPart?.Styles?.Elements<Style>()
                .FirstOrDefault(style =>
                    style.StyleId?.Value == styleId &&
                    (style.Type is null || style.Type.Value == StyleValues.Paragraph));

    /// <summary>
    /// O nível de um estilo pelo nome dele: `heading 1` a `heading 6`, sem
    /// importar maiúsculas — o LibreOffice grava `Heading 1`.
    /// </summary>
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

    private string? Copied(int level)
    {
        if (touched is null) return null;

        var definitions = part.StyleDefinitionsPart ?? part.AddNewPart<StyleDefinitionsPart>();
        var styles = definitions.Styles ??= new Styles();

        var ids = styles.Elements<Style>()
            .Select(style => style.StyleId?.Value)
            .OfType<string>()
            .ToHashSet(StringComparer.Ordinal);

        var style = TemplateStyles.Heading(level);

        // Um `Heading1` que não se chama `heading 1` é estilo de outra coisa, e
        // não pode ser sobrescrito: o título ganha um id que ninguém usa.
        var id = style.StyleId!.Value!;
        for (var suffix = 2; ids.Contains(id); suffix++) id = $"Heading{level}_{suffix}";
        style.StyleId = id;

        // A cadeia do modelo parte do `Normal`. Num documento que chama o estilo
        // padrão de outro jeito, herdar de um id que não existe é herdar de nada,
        // e o SDK acusa a referência pendurada.
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
