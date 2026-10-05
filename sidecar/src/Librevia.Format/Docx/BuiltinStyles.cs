namespace Librevia.Format.Docx;

/// <summary>
/// A style as **data**, in editor units; converting to OOXML is <see cref="TemplateStyles"/>'s job.
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
/// The CSS the editor drew before styles, as data: the headings the writer adds to a DOCX without
/// them and the styles of a <c>.sdoc</c> before version 3. It is the TS side's
/// <c>LEGACY_STYLES</c>, compared by <c>src/main/sidecar/builtin-styles.test.ts</c>. Table and
/// numbering styles live in <see cref="TemplateStyles"/>.
/// </summary>
public static class BuiltinStyles
{
    public const string BodyFont = "Times New Roman";

    /// <summary>In points (<c>font-size: 12pt</c>).</summary>
    public const double BodySizePt = 12;

    /// <summary>The plain text band's (<c>page-setup.ts</c>).</summary>
    public const string BandFont = "Calibri";

    /// <summary>
    /// 1.5 ÷ 1.1499 (Liberation Serif, <see cref="LineMetrics"/>), to four decimals: the
    /// <c>w:line</c> 240ths grid.
    /// </summary>
    public const double BodyLineFactor = 1.3042;

    /// <summary>
    /// In <c>word/styles.xml</c> order. A heading's space before and after are the browser's
    /// margins; "keep with next" only on the first four. The text's <c>#111111</c> is left out: it
    /// would come back as an explicit color.
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

    /// <summary>By internal name (<see cref="HeadingStyles.LevelOfName"/>).</summary>
    public static int HeadingLevels =>
        All.Count(style => style.Name.StartsWith("heading ", StringComparison.Ordinal));

    public static BuiltinStyle Heading(int level) =>
        All.First(style => style.Name == $"heading {level}");
}
