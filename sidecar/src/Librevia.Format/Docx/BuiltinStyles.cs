namespace Librevia.Format.Docx;

/// <summary>
/// Um estilo do documento novo, em **dados** — nas unidades do editor.
/// </summary>
/// <remarks>
/// Pontos, milímetros e o fator da entrelinha, e não twips nem meios-pontos: são
/// estas as unidades que o modelo do documento usa (`src/services/document/styles.ts`),
/// e é a tabela abaixo que o teste de contrato compara com o `BUILTIN_STYLES` de
/// lá. A conversão para o OOXML mora em <see cref="TemplateStyles"/>, de um lado
/// só, para que a tabela possa ser lida pelos dois.
/// </remarks>
/// <param name="Id">O `w:styleId`.</param>
/// <param name="Name">O `w:name` — o nome interno, que não se traduz.</param>
/// <param name="Character">Estilo de caractere; falso é estilo de parágrafo.</param>
/// <param name="BasedOn">De quem herda (`w:basedOn`).</param>
/// <param name="Next">O estilo do parágrafo seguinte (`w:next`).</param>
/// <param name="Link">O estilo de caractere ligado a este (`w:link`).</param>
/// <param name="UiPriority">A ordem na galeria do Word (`w:uiPriority`).</param>
/// <param name="QFormat">Estilo recomendado, que aparece na galeria (`w:qFormat`).</param>
/// <param name="Hidden">Escondido sempre (`w:hidden`).</param>
/// <param name="SemiHidden">Escondido até ser usado (`w:semiHidden`).</param>
/// <param name="UnhideWhenUsed">Reaparece ao ser usado (`w:unhideWhenUsed`).</param>
/// <param name="Default">O estilo que vale sem `w:pStyle` (`w:default="1"`).</param>
/// <param name="SizePt">Tamanho da fonte em pontos.</param>
/// <param name="Bold">Negrito.</param>
/// <param name="Color">Cor do texto, em hexadecimal com `#`.</param>
/// <param name="Underline">Sublinhado simples.</param>
/// <param name="BeforePt">Espaço antes do parágrafo, em pontos.</param>
/// <param name="AfterPt">Espaço depois, em pontos.</param>
/// <param name="LineFactor">Entrelinha como múltiplo da altura natural da linha.</param>
/// <param name="IndentMm">Recuo esquerdo, em milímetros.</param>
/// <param name="KeepNext">Não fica sozinho no pé da página.</param>
/// <param name="ContextualSpacing">Sem espaço entre parágrafos do mesmo estilo.</param>
/// <param name="OutlineLevel">Nível na estrutura do documento, de 0 a 8.</param>
public sealed record BuiltinStyle(
    string Id,
    string Name,
    bool Character = false,
    string? BasedOn = null,
    string? Next = null,
    string? Link = null,
    int? UiPriority = null,
    bool QFormat = false,
    bool Hidden = false,
    bool SemiHidden = false,
    bool UnhideWhenUsed = false,
    bool Default = false,
    double? SizePt = null,
    bool Bold = false,
    string? Color = null,
    bool Underline = false,
    double? BeforePt = null,
    double? AfterPt = null,
    double? LineFactor = null,
    double? IndentMm = null,
    bool KeepNext = false,
    bool ContextualSpacing = false,
    int? OutlineLevel = null);

/// <summary>
/// Os estilos de parágrafo e de caractere do documento novo.
/// </summary>
/// <remarks>
/// Não é um gosto: é a **tela**. O editor desenha o documento novo com o CSS de
/// `src/services/document/content-styles.ts` e com o padrão do navegador para o
/// que o CSS não diz, e o arquivo tem de reproduzir exatamente isso — se tela e
/// arquivo discordassem, salvar em DOCX mudaria a paginação do que a pessoa
/// acabou de escrever. Os números foram medidos no editor (estilo computado de
/// `p` e de `h1` a `h6` num documento novo), e não copiados do Word.
///
/// A mesma tabela existe em `src/services/document/styles.ts` como
/// `BUILTIN_STYLES`, porque é ela que um `.sdoc` gravado antes da versão 3 do
/// formato recebe ao ser aberto — assim o documento antigo abre idêntico. As
/// duas são comparadas a cada execução por `src/main/sidecar/builtin-styles.test.ts`:
/// uma medida mudada de um lado só faria tela e arquivo discordarem em silêncio.
///
/// Os estilos de tabela e de numeração (`TableNormal`, `NoList`, `TableGrid`) não
/// estão aqui: o modelo do documento só representa estilo de parágrafo e de
/// caractere, e o que eles carregam — bordas, margem de célula — não tem onde
/// morar nele. Continuam montados à mão em <see cref="TemplateStyles"/>.
/// </remarks>
public static class BuiltinStyles
{
    /// <summary>A fonte do documento novo — a primeira da pilha do CSS.</summary>
    public const string BodyFont = "Times New Roman";

    /// <summary>O tamanho do corpo do texto, em pontos (`font-size: 12pt`).</summary>
    public const double BodySizePt = 12;

    /// <summary>A fonte do cabeçalho e do rodapé de texto simples (`page-setup.ts`).</summary>
    public const string BandFont = "Calibri";

    /// <summary>
    /// A entrelinha 1,5 do CSS, em múltiplos da altura natural da linha.
    /// </summary>
    /// <remarks>
    /// 1,5 ÷ 1,1499 (a altura da Liberation Serif, ver <see cref="LineMetrics"/>)
    /// = 1,3042. O múltiplo do OOXML é medido sobre a altura da fonte, e não
    /// sobre o tamanho dela; tratar um pelo outro encurtava cada linha em 15 %.
    ///
    /// Arredondado a quatro casas de propósito: o `w:line` vive numa grade de
    /// 240-avos (1,3042 × 240 = 313), e é nessa grade que o número volta do
    /// arquivo. Guardar 1,304461 aqui faria o valor lido divergir do declarado.
    /// </remarks>
    public const double BodyLineFactor = 1.3042;

    /// <summary>
    /// Os estilos, na ordem em que entram no `word/styles.xml`.
    /// </summary>
    /// <remarks>
    /// O antes do título é o `margin-top: 1em` do CSS (0,6em no `h5` e no `h6`,
    /// que não têm regra própria); o depois é a margem de baixo do navegador —
    /// 0,67em no `h1`, 0,83em no `h2`, 1em no `h3`, 1,33em no `h4`, 1,67em e
    /// 2,33em nos dois últimos. O "manter com o próximo" é o `break-after: avoid`
    /// da impressão, que só os quatro primeiros têm.
    ///
    /// Duas aproximações, ambas abaixo do que o olho vê: o tamanho do `h5` e do
    /// `h6` (9,96 pt e 8,04 pt no navegador) vai para o meio-ponto mais próximo,
    /// que é a precisão do `w:sz`; e o `#111111` do texto fica de fora, porque
    /// gravado como cor ele voltaria como cor explícita em cada trecho reaberto.
    /// </remarks>
    public static readonly BuiltinStyle[] All =
    [
        new(
            "Normal", "Normal", QFormat: true, Default: true,
            BeforePt: 7.2, AfterPt: 12, LineFactor: BodyLineFactor),
        new(
            "DefaultParagraphFont", "Default Paragraph Font", Character: true,
            UiPriority: 1, SemiHidden: true, UnhideWhenUsed: true, Default: true),
        new(
            "Heading1", "heading 1", BasedOn: "Normal", Next: "Normal", UiPriority: 9, QFormat: true,
            SizePt: 22, Bold: true, BeforePt: 22, AfterPt: 14.75, KeepNext: true, OutlineLevel: 0),
        new(
            "Heading2", "heading 2", BasedOn: "Normal", Next: "Normal", UiPriority: 9, QFormat: true,
            SizePt: 17, Bold: true, BeforePt: 17, AfterPt: 14.1, KeepNext: true, OutlineLevel: 1),
        new(
            "Heading3", "heading 3", BasedOn: "Normal", Next: "Normal", UiPriority: 9, QFormat: true,
            SizePt: 14, Bold: true, BeforePt: 14, AfterPt: 14, KeepNext: true, OutlineLevel: 2),
        new(
            "Heading4", "heading 4", BasedOn: "Normal", Next: "Normal", UiPriority: 9, QFormat: true,
            SizePt: 12, Bold: true, BeforePt: 12, AfterPt: 15.95, KeepNext: true, OutlineLevel: 3),
        new(
            "Heading5", "heading 5", BasedOn: "Normal", Next: "Normal", UiPriority: 9, QFormat: true,
            SizePt: 10, Bold: true, BeforePt: 6, AfterPt: 16.65, OutlineLevel: 4),
        new(
            "Heading6", "heading 6", BasedOn: "Normal", Next: "Normal", UiPriority: 9, QFormat: true,
            SizePt: 8, Bold: true, BeforePt: 4.8, AfterPt: 18.75, OutlineLevel: 5),
        new(
            "ListParagraph", "List Paragraph", BasedOn: "Normal", UiPriority: 34, QFormat: true,
            IndentMm: 12.7, ContextualSpacing: true),
        new(
            "Hyperlink", "Hyperlink", Character: true, BasedOn: "DefaultParagraphFont",
            UiPriority: 99, UnhideWhenUsed: true, Color: "#0563c1", Underline: true),
    ];

    /// <summary>Quantos níveis de título o documento novo define.</summary>
    /// <remarks>
    /// Pelo nome interno, que é o critério com que o leitor e o escritor
    /// reconhecem um título (<see cref="HeadingStyles.LevelOfName"/>) — e não
    /// pela contagem de estilos, que inclui o resto da tabela.
    /// </remarks>
    public static int HeadingLevels =>
        All.Count(style => style.Name.StartsWith("heading ", StringComparison.Ordinal));

    /// <summary>A definição de um nível de título, pelo nome interno dele.</summary>
    public static BuiltinStyle Heading(int level) =>
        All.First(style => style.Name == $"heading {level}");
}
