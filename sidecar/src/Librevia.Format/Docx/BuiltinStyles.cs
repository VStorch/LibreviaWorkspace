namespace Librevia.Format.Docx;

/// <summary>
/// Um estilo em **dados**, nas unidades do editor; a conversão para OOXML é de
/// <see cref="TemplateStyles"/>.
/// </summary>
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
    int? OutlineLevel = null,
    bool Italic = false,
    bool KeepLines = false);

/// <summary>
/// O CSS que o editor desenhava antes dos estilos, em dados: os títulos que o
/// escritor acrescenta a um DOCX sem eles e os estilos do <c>.sdoc</c> anterior à versão
/// 3. É a <c>LEGACY_STYLES</c> do lado TS, comparada por
/// <c>src/main/sidecar/builtin-styles.test.ts</c>. Estilos de tabela e de numeração
/// ficam em <see cref="TemplateStyles"/>.
/// </summary>
public static class BuiltinStyles
{
    public const string BodyFont = "Times New Roman";

    /// <summary>Em pontos (<c>font-size: 12pt</c>).</summary>
    public const double BodySizePt = 12;

    /// <summary>A da faixa de texto simples (<c>page-setup.ts</c>).</summary>
    public const string BandFont = "Calibri";

    /// <summary>
    /// 1,5 ÷ 1,1499 (Liberation Serif, <see cref="LineMetrics"/>), a quatro casas: a
    /// grade de 240-avos do <c>w:line</c>.
    /// </summary>
    public const double BodyLineFactor = 1.3042;

    /// <summary>
    /// Na ordem de <c>word/styles.xml</c>. Antes e depois do título são as margens do
    /// navegador; "manter com o próximo" só nos quatro primeiros. O <c>#111111</c> do
    /// texto fica de fora: voltaria como cor explícita.
    /// </summary>
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

    /// <summary>Pelo nome interno (<see cref="HeadingStyles.LevelOfName"/>).</summary>
    public static int HeadingLevels =>
        All.Count(style => style.Name.StartsWith("heading ", StringComparison.Ordinal));

    public static BuiltinStyle Heading(int level) =>
        All.First(style => style.Name == $"heading {level}");
}
