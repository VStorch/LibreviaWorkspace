using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>text:p</c> and <c>text:h</c> → <c>paragraph</c> and <c>heading</c>, with the attributes
/// <see cref="BodyReader"/> gives a DOCX paragraph. A loose body paragraph carries only its
/// automatic style (the direct formatting): its named style draws the rest. In a list or a cell,
/// where the style CSS does not reach, the block carries its effective formatting
/// (<paramref name="flat"/>).
/// </summary>
internal sealed class OdtParagraphReader(OdtReadContext context)
{
    private const int MaxHeadingLevel = 6;

    private readonly OdtInlineReader _inline = new(context);

    private sealed record Formats(
        OdtStyle? Named,
        ParagraphFormat Direct,
        CharacterFormat DirectCharacter,
        ParagraphFormat Effective,
        CharacterFormat EffectiveCharacter,
        CharacterFormat NamedCharacter);

    /// <summary>The block, preceded by a <c>pageBreak</c> when it asks for one before it.</summary>
    public IEnumerable<Node> Read(XElement element, bool flat)
    {
        var styleName = element.Attr(Ns.Text + "style-name");
        var style = context.Styles.Find(OdtStyles.ParagraphFamily, styleName);
        var breakBefore = !flat && !context.IsFirstBlock && style is { Automatic: true } &&
                          ParagraphFormat.Of(style.Properties("paragraph-properties")).PageBreakBefore;
        context.SeeBlock(style?.MasterPage);

        var formats = FormatsOf(styleName, style);
        var node = NodeOf(element, formats);
        if (flat) Flattened(node, formats);
        else DirectOnly(node, formats);

        // What the block draws: the style CSS where it reaches, and the node's own font and size.
        var drawn = flat ? context.Styles.DefaultCharacter() : formats.NamedCharacter;
        var inherited = drawn with
        {
            FontFamily = formats.EffectiveCharacter.FontFamily,
            FontSizePt = formats.EffectiveCharacter.FontSizePt,
        };
        var content = _inline.Read(element, formats.EffectiveCharacter, inherited);
        node.Content = content.Count == 0 ? null : content;

        if (breakBefore) yield return Node.Of("pageBreak");
        yield return node;
    }

    private Formats FormatsOf(string? styleName, OdtStyle? style)
    {
        var styles = context.Styles;
        var chain = styles.Chain(OdtStyles.ParagraphFamily, styleName);
        var namedChain = chain.Where(link => !link.Automatic).ToList();
        var defaults = styles.DefaultCharacter();
        var automatic = style is { Automatic: true } ? style : null;
        return new Formats(
            styles.NamedOf(OdtStyles.ParagraphFamily, styleName),
            ParagraphFormat.Of(automatic?.Properties("paragraph-properties")),
            CharacterFormat.Of(automatic?.Properties("text-properties"), styles.Fonts),
            styles.DefaultParagraph().Over(styles.Paragraph(chain)),
            styles.Character(chain, defaults),
            styles.Character(namedChain, defaults));
    }

    /// <summary>A heading past level 6 is a paragraph: the editor stops there, as Word's styles.</summary>
    private Node NodeOf(XElement element, Formats formats)
    {
        var level = element.Name.LocalName == "h"
            ? OdtLength.Integer(element.Attr(Ns.Text + "outline-level")) ?? 1
            : 0;
        var node = level is >= 1 and <= MaxHeadingLevel ? Node.Of("heading").With("level", level) : Node.Of("paragraph");

        if (formats.Named is { } named && named.Name != OdtStyles.DefaultParagraphStyle) node.With("styleId", named.Id);
        if (formats.Direct.BreakAfter == ParagraphFormat.Page) node.With("breakAfter", true);
        if (formats.Direct.BreakAfter == ParagraphFormat.Column) node.With("columnBreakAfter", true);
        return node;
    }

    private static void DirectOnly(Node node, Formats formats)
    {
        var direct = formats.Direct;
        var effective = formats.Effective;
        if (direct.TextAlign is not null) node.With("textAlign", effective.TextAlign);
        if (direct.IndentMm is not null) node.With("indentMm", Math.Max(0, effective.IndentMm ?? 0));
        if (direct.IndentRightMm is not null) node.With("indentRightMm", Math.Max(0, effective.IndentRightMm ?? 0));
        if (direct.FirstLineMm is not null) node.With("firstLineMm", effective.FirstLineMm ?? 0);
        if (direct.Background is { } background)
        {
            if (background != CharacterFormat.Transparent) node.With("background", background);
            else if (effective.Background is not null) node.With("background", CharacterFormat.Transparent);
        }

        if (direct.SpaceBefore is not null) node.With("spaceBefore", effective.SpaceBefore ?? 0);
        if (direct.SpaceAfter is not null) node.With("spaceAfter", effective.SpaceAfter ?? 0);
        if (DirectLineHeight(formats) is { } lineHeight) node.With("lineHeight", lineHeight);
        if (formats.DirectCharacter.FontFamily is not null && formats.EffectiveCharacter.FontFamily is { } font)
        {
            node.With("fontFamily", font);
        }

        if ((formats.DirectCharacter.FontSizePt ?? formats.DirectCharacter.FontSizeFactor) is not null &&
            formats.EffectiveCharacter.FontSizePt is { } size)
        {
            node.With("fontSize", RunReader.FormatPoints(size));
        }

        if (direct.KeepNext is not null) node.With("keepNext", effective.KeepNext == true);
        if (direct.KeepLines is not null) node.With("keepLines", effective.KeepLines == true);
        if (direct.WidowControl is not null) node.With("widowControl", effective.WidowControl != false);
    }

    /// <summary>
    /// Declared, or changed by the block's own font: the multiple is measured against the font
    /// that draws the line.
    /// </summary>
    private static string? DirectLineHeight(Formats formats)
    {
        var spacing = formats.Effective.LineSpacing;
        var own = ParagraphFormat.CssLineHeight(spacing, formats.EffectiveCharacter.FontFamily);
        if (formats.Direct.LineSpacing is not null) return own;
        if (formats.DirectCharacter.FontFamily is null || spacing is null) return null;
        return own != ParagraphFormat.CssLineHeight(spacing, formats.NamedCharacter.FontFamily) ? own : null;
    }

    /// <summary>As <see cref="BodyReader"/> flattens: spacing and line height always stated.</summary>
    private static void Flattened(Node node, Formats formats)
    {
        var effective = formats.Effective;
        var font = formats.EffectiveCharacter.FontFamily;
        if (effective.TextAlign is { } align) node.With("textAlign", align);
        if (effective.IndentMm is > 0) node.With("indentMm", effective.IndentMm.Value);
        if (effective.IndentRightMm is > 0) node.With("indentRightMm", effective.IndentRightMm.Value);
        if (effective.FirstLineMm is { } firstLine and not 0) node.With("firstLineMm", firstLine);
        if (effective.Background is { } background && background != CharacterFormat.Transparent)
        {
            node.With("background", background);
        }

        node.With("spaceBefore", effective.SpaceBefore ?? 0);
        node.With("spaceAfter", effective.SpaceAfter ?? 0);
        node.With("lineHeight", ParagraphFormat.CssLineHeight(
            effective.LineSpacing ?? ParagraphFormat.SingleSpacing, font));
        if (font is not null) node.With("fontFamily", font);
        if (formats.EffectiveCharacter.FontSizePt is { } size) node.With("fontSize", RunReader.FormatPoints(size));
        if (effective.KeepNext == true) node.With("keepNext", true);
        if (effective.KeepLines == true) node.With("keepLines", true);
        if (effective.WidowControl == false) node.With("widowControl", false);
    }

    /// <summary>An empty <c>text:p</c>, for a list item or cell that has none.</summary>
    public static Node Empty() => Node.Of("paragraph");
}
