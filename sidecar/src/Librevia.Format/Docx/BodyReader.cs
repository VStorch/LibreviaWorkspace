using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;

namespace Librevia.Format.Docx;

/// <summary>A top-level body block, with identity.</summary>
/// <param name="Oid">A stable id in document order: b1, b2, …</param>
/// <param name="Source">The original element, to write back untouched.</param>
/// <param name="Extracted">What the editor sees.</param>
public sealed record Block(string Oid, OpenXmlElement Source, Node Extracted)
{
    /// <summary>
    /// What the body keeps **before** the block that is not a block: a loose bookmark after a
    /// table, an unknown content control. Saving returns it before the block, byte for byte;
    /// without this it would silently drop out.
    /// </summary>
    public List<OpenXmlElement> Leading { get; } = [];

    /// <summary>The same, after the body's last block.</summary>
    public List<OpenXmlElement> Trailing { get; } = [];

    /// <summary>
    /// A table of contents without a content control, whose field closes in one of the paragraphs
    /// below.
    /// </summary>
    public List<OpenXmlElement> Continuation { get; } = [];

    /// <summary>A field shown only by its result: rewritten, it becomes plain text.</summary>
    public bool UnrepresentedField { get; init; }
}

/// <summary>
/// Document body → editor nodes, **one way**: it feeds the screen and the PDF, never saving, which
/// starts from the original XML (<see cref="DocxWriter"/>). So a mistake here is cosmetic, and
/// reading can be lenient.
/// </summary>
/// <remarks>
/// Each switch turned off reproduces the reading from before the feature, the reference reading for
/// drafts of that time (<see cref="DocumentModelDto.BeforeReferences"/>, <c>BeforeSections</c>,
/// <c>BeforeComments</c>, <c>BeforeRevisions</c>, <c>BeforeNotes</c>, <c>BeforeMath</c>).
/// </remarks>
public sealed class BodyReader(
    MainDocumentPart part,
    Inventory inventory,
    bool flatten = false,
    bool references = true,
    bool sections = true,
    bool comments = true,
    bool revisions = true,
    bool notes = true,
    bool math = true)
{
    /// <summary>The <c>w:footnote</c>/<c>w:endnote</c> and its blocks.</summary>
    public sealed record NoteRead(OpenXmlElement Source, List<Block> Blocks);

    /// <summary>
    /// By address (<c>fn:3</c>, <c>en:1</c>). A note the body does not reference stays as it is.
    /// </summary>
    public Dictionary<string, NoteRead> Notes { get; } = new(StringComparer.Ordinal);

    /// <summary>
    /// Another instance: the state of the paragraph being read cannot be the note's.
    /// </summary>
    private BodyReader? _noteReader;

    /// <summary>
    /// The part owning the relationships: the document, or <c>footnotes.xml</c> when reading a
    /// note.
    /// </summary>
    private OpenXmlPart _owner = part;

    /// <summary>
    /// Null in the body (<c>b1</c>…); <c>fn:3</c> in a note, whose blocks come out <c>fn:3/p1</c>…
    /// </summary>
    private string? _oidPrefix;

    /// <summary>The one from each surrounding <c>w:ins</c>/<c>w:del</c>, outside in.</summary>
    private readonly List<Mark> _revision = [];

    /// <summary>
    /// The move name of each `w:moveFrom`/`w:moveTo`; see <see cref="MoveNamesOf"/>.
    /// </summary>
    private Dictionary<OpenXmlElement, string>? _moveNames;

    /// <summary>
    /// Reply ends do not become nodes: the editor carries one anchor per thread, the same on open
    /// and in the reference reading, and the fingerprint stays stable.
    /// </summary>
    private readonly HashSet<string> _replies = comments
        ? CommentsReader.RepliesOf(part).Values.SelectMany(ids => ids).ToHashSet(StringComparer.Ordinal)
        : [];

    /// <summary>
    /// A comment without <c>w:commentRangeEnd</c> is a point comment: the reference becomes the
    /// <c>commentEnd</c>.
    /// </summary>
    private readonly HashSet<string> _rangeEnds = comments
        ? new OpenXmlElement?[] { part.Document, part.FootnotesPart?.Footnotes, part.EndnotesPart?.Endnotes }
            .SelectMany(root => root?.Descendants<CommentRangeEnd>() ?? [])
            .Select(end => end.Id?.Value).OfType<string>().ToHashSet(StringComparer.Ordinal)
        : [];

    /// <summary>
    /// Inside a text box the anchor does not become a node: the box is not rewritten from the
    /// model.
    /// </summary>
    private int _textBoxDepth;

    /// <summary>The same id <see cref="PageReader"/> gives the section.</summary>
    private Dictionary<SectionProperties, string> _sectionIds = new(ReferenceEqualityComparer.Instance);

    private readonly NumberingReader _numbering = new(part);
    private readonly StyleResolver _styles = new(part);
    private readonly FontTable _fonts = new(part);
    private int _nextId = 1;

    /// <summary>
    /// Decides whether a text box starts a new line: a paragraph's boxes are in different
    /// <c>w:r</c>s, and none sees what the previous one wrote.
    /// </summary>
    private bool _paragraphHasContent;

    /// <summary>
    /// Only on a loose body paragraph, which the style CSS draws. Lists, cells, boxes and old
    /// drafts carry the whole formatting.
    /// </summary>
    private bool _directRuns;

    /// <summary>
    /// Found while reading lines, but belonging to the **block**: attached to the paragraph node at
    /// the end.
    /// </summary>
    private readonly List<FloatDto> _paragraphFloats = [];

    /// <summary>See <see cref="TopAnchoredFirst"/>.</summary>
    private readonly HashSet<Node> _topAnchored = new(ReferenceEqualityComparer.Instance);

    /// <summary>
    /// Rewriting the block would lose the field that did not become a node; see <see
    /// cref="Block.UnrepresentedField"/>.
    /// </summary>
    private bool _unrepresentedField;

    /// <summary>
    /// The tree and the block list come out together because blocks **point to tree nodes**: a list
    /// item is a <c>w:p</c> in the file and a nested <c>listItem</c>.
    /// </summary>
    public (List<Node> Content, List<Block> Blocks) Read(Body body)
    {
        if (sections) _sectionIds = PageReader.SectionIds(body);
        return ReadBlocks(body);
    }

    /// <summary>
    /// Flattened, like lists and cells. The <c>w:footnoteRef</c> run, the number at the start of
    /// the note, does not become a node: saving rebuilds it.
    /// </summary>
    private (List<Node> Content, List<Block> Blocks) ReadNote(OpenXmlElement note, string address, OpenXmlPart owner)
    {
        _owner = owner;
        _oidPrefix = address;
        _nextId = 1;
        return ReadBlocks(note);
    }

    private (List<Node> Content, List<Block> Blocks) ReadBlocks(OpenXmlElement container)
    {
        var content = new List<Node>();
        var blocks = new List<Block>();

        // Consecutive numbered paragraphs become a single list; one open list per level.
        var openLists = new List<(Node List, string Kind, int Level, int NumId)>();

        // See Block.Leading.
        var loose = new List<OpenXmlElement>();

        void Add(Block block)
        {
            block.Leading.AddRange(loose);
            loose.Clear();
            blocks.Add(block);
        }

        var elements = container.ChildElements.ToList();
        for (var at = 0; at < elements.Count; at++)
        {
            var element = elements[at];

            // See ReadTableOfContents.
            if (references && TableOfContentsAt(elements, at) is { } toc)
            {
                openLists.Clear();
                var block = NewBlock(element, toc.Node);
                block.Continuation.AddRange(toc.Continuation);
                Add(block);
                content.Add(toc.Node);
                at += toc.Continuation.Count;
                continue;
            }

            switch (element)
            {
                case Paragraph paragraph:
                {
                    // Only a loose paragraph in the body stops being flattened (`.page__content >
                    // p`).
                    var numbered = _numbering.ListKindOf(paragraph.ParagraphProperties);
                    var node = ReadParagraph(paragraph, flat: flatten || _oidPrefix is not null || numbered is not null);
                    var list = node.Type == "pageBreak" ? null : numbered;

                    if (list is null)
                    {
                        openLists.Clear();
                        Add(NewBlock(element, node));
                        content.Add(node);
                        break;
                    }

                    var (kind, level) = (list.Kind, list.Level);
                    while (openLists.Count > 0 && openLists[^1].Level > level)
                    {
                        openLists.RemoveAt(openLists.Count - 1);
                    }

                    // At the same level, another numbering is another list, as in Word.
                    if (openLists.Count > 0 && openLists[^1].Level == level &&
                        (openLists[^1].Kind != kind || openLists[^1].NumId != list.NumberingId))
                    {
                        openLists.RemoveAt(openLists.Count - 1);
                    }

                    if (openLists.Count == 0 || openLists[^1].Level < level)
                    {
                        var listNode = Node.Of(kind);
                        listNode.Content = [];
                        var parent = openLists.Count > 0 ? openLists[^1] : default;
                        var depth = openLists.Count;

                        // The marker and the indent belong to the level. The `numId` travels on the
                        // node: without it the writer would write `w:numId w:val="0"`, "no
                        // numbering".
                        listNode.With("numId", list.NumberingId);

                        if (list.Marker is { } marker) listNode.With("marker", marker);

                        // The `start` the editor materializes on every numbered list: it enters the
                        // outer item's fingerprint.
                        if (kind == "orderedList") listNode.With("start", 1);
                        if (list.IndentMm is { } indent) listNode.With("indentMm", indent);
                        if (list.HangingMm is { } hanging) listNode.With("hangingMm", hanging);

                        // The definition only where the numbering changes.
                        if (depth == 0 || parent.NumId != list.NumberingId)
                        {
                            listNode.With("numbering", list.Definition);
                        }

                        // A list starting at level 2 lives at the top of the tree: without this it
                        // would go back to the file at level 0.
                        if (level != (depth == 0 ? 0 : parent.Level + 1)) listNode.With("level", level);

                        if (depth > 0)
                        {
                            var parentItems = parent.List.Content!;
                            (parentItems[^1].Content ??= []).Add(listNode);
                        }
                        else
                        {
                            content.Add(listNode);
                        }

                        openLists.Add((listNode, kind, level, list.NumberingId));
                    }

                    // On the `listItem`: the item is what corresponds to a `w:p`.
                    var item = Node.Of("listItem");
                    item.Content = [node];
                    openLists[^1].List.Content!.Add(item);
                    CarrySpacing(openLists[^1].List, node);
                    Add(NewBlock(element, item));

                    // The section ends at the item, as in Word: the next one opens another list.
                    if (node.Attrs?.ContainsKey("sectionBreak") == true) openLists.Clear();
                    break;
                }

                case Table table:
                {
                    openLists.Clear();
                    var node = ReadTable(table);
                    Add(NewBlock(element, node));
                    content.Add(node);
                    break;
                }

                case SectionProperties:
                    break;

                default:
                    inventory.NoteInvisibleElement(element.LocalName);
                    loose.Add(element);
                    break;
            }
        }

        // What is left after the last block stays with it.
        if (blocks.Count > 0) blocks[^1].Trailing.AddRange(loose);

        if (content.Count == 0) content.Add(Node.Of("paragraph"));

        return (content, blocks);
    }

    /// <summary>
    /// In the tree the list is an element and would get the editor's spacing: in a document with
    /// six lists, fifteen extra millimetres. Before the first item the item's space before applies;
    /// after the last, its space after, as in Word.
    /// </summary>
    private static void CarrySpacing(Node list, Node paragraph)
    {
        if (paragraph.Attrs is not { } attrs) return;

        if (list.Attrs?.ContainsKey("spaceBefore") != true && attrs.TryGetValue("spaceBefore", out var before))
        {
            list.With("spaceBefore", before?.DeepClone());
        }

        if (attrs.TryGetValue("spaceAfter", out var after)) list.With("spaceAfter", after?.DeepClone());
    }

    private Block NewBlock(OpenXmlElement source, Node extracted)
    {
        var oid = _oidPrefix is null ? "b" + _nextId++ : $"{_oidPrefix}/p{_nextId++}";
        extracted.With("oid", oid);
        var block = new Block(oid, source, extracted) { UnrepresentedField = _unrepresentedField };
        _unrepresentedField = false;
        return block;
    }


    private Node ReadParagraph(Paragraph paragraph, bool flat)
    {
        var direct = paragraph.ParagraphProperties;

        // Effective formatting: defaults, style and direct.
        var (effective, inheritedRun) = _styles.Resolve(direct);

        _paragraphHasContent = false;
        _paragraphFloats.Clear();
        _topAnchored.Clear();
        var outer = _directRuns;
        _directRuns = !flat;
        var content = ReadInline(paragraph, inheritedRun);
        _directRuns = outer;

        // A lone page break is the `pageBreak` node, **except** when the paragraph anchors
        // something: the object would fall onto the sheet below. Then the break becomes a paragraph
        // property.
        if (content.Count == 1 && content[0].Type == "pageBreak" && _paragraphFloats.Count == 0)
        {
            return content[0];
        }

        // A break mid-paragraph (`w:br w:type="page"` in a `w:r`) becomes a block property: as a
        // node at line position it would be invalid, and in the HTML it would misalign indexes
        // between screen and paper. Exact when the break ends the paragraph; approximate when text
        // follows it.
        var breakAfter = content.RemoveAll(child => child.Type == "pageBreak") > 0;
        var columnBreakAfter = content.RemoveAll(child => child.Type == "columnBreak") > 0;
        TopAnchoredFirst(content);

        var node = (HeadingLevelOf(direct) ?? _styles.HeadingLevelByName(direct?.ParagraphStyleId?.Val?.Value)) is { } level
            ? Node.Of("heading").With("level", level)
            : Node.Of("paragraph");

        if (breakAfter) node.With("breakAfter", true);
        if (columnBreakAfter) node.With("columnBreakAfter", true);

        WithFloats(node);

        // So the edited paragraph keeps pointing to the original style.
        if (direct?.ParagraphStyleId?.Val?.Value is { Length: > 0 } styleId)
        {
            node.With("styleId", styleId);
        }

        var alignment = AlignmentOf(effective);

        // Tabs at the start of the line position, they are not text; see TabAlignmentOf.
        var viaTabs = TabAlignmentOf(effective, content);
        if (viaTabs is not null)
        {
            alignment = viaTabs;
            while (content.Count > 0 && content[0].Type == "text" && content[0].Text == "\t")
            {
                content.RemoveAt(0);
            }
        }

        if (flat) Flattened(node, alignment, effective, direct, inheritedRun);
        else DirectOnly(node, viaTabs, effective, direct, inheritedRun);

        // The empty paragraph holding `w:sectPr` **is** the section mark, without height, as in
        // LibreOffice. It stays in the model: saving returns it.
        if (direct?.SectionProperties is not null && content.Count == 0)
        {
            node.With("sectionMark", true);
        }

        // The setup lives outside the nodes, in `sections`; the id is the link.
        if (direct?.SectionProperties is { } marked && _sectionIds.TryGetValue(marked, out var sectionId))
        {
            node.With("sectionBreak", sectionId);
        }

        // A revised paragraph mark: accepting the deletion joins this one to the next.
        if (revisions && Revisions.BlockRevisionOf(direct?.ParagraphMarkRunProperties) is { } markRevision)
        {
            node.With("markRevision", markRevision);
        }

        node.Content = content.Count == 0 ? null : content;
        return node;
    }
    /// <summary>
    /// Defaults, style and direct, flattened: where style rules do not reach (lists, cells) and in
    /// <c>flatten</c> mode, the reference reading for the flattened draft.
    /// </summary>
    private void Flattened(
        Node node,
        string? alignment,
        ParagraphProperties effective,
        ParagraphProperties? direct,
        RunProperties inheritedRun)
    {
        if (alignment is not null) node.With("textAlign", alignment);


        // Always zero: the file's indent comes in millimetres just below, and the editor returns
        // `indent` on every paragraph, so omitting it would make the block look changed.
        node.With("indent", 0);

        // In millimetres, as the file declares it.
        Measure(node, "indentMm", effective.Indentation?.Left?.Value);
        Measure(node, "indentRightMm", effective.Indentation?.Right?.Value);

        // `w:firstLine` pushes and `w:hanging` pulls: the same CSS `text-indent`.
        var firstLine = TwipsToMm(effective.Indentation?.FirstLine?.Value);
        var hanging = TwipsToMm(effective.Indentation?.Hanging?.Value);
        if (firstLine is > 0) node.With("firstLineMm", firstLine.Value);
        else if (hanging is > 0) node.With("firstLineMm", -hanging.Value);

        // The background is what turns the corpus `Heading1` into a colored bar.
        if (ShadingOf(effective) is { } background) node.With("background", background);

        // **Always** written: silence in the file is zero, not the editor default.
        var spacing = effective.SpacingBetweenLines;
        node.With("spaceBefore", TwipsToPt(spacing?.Before?.Value) ?? 0);
        node.With("spaceAfter", TwipsToPt(spacing?.After?.Value) ?? 0);
        // The font that measures the line is the paragraph mark's.
        var markFont = _styles.ResolveMark(inheritedRun, direct).RunFonts?.Ascii?.Value;
        node.With("lineHeight", LineHeightOf(spacing, LineMetrics.Of(markFont)));

        // The **block's** font, from the paragraph mark (`w:pPr/w:rPr`): line height comes from the
        // element's font, and the mark gives an empty paragraph its height, as in Word.
        var mark = _styles.ResolveMark(inheritedRun, direct);
        if (FontOf(mark) is { } font) node.With("fontFamily", font);
        if (FontSizeOf(mark) is { } size) node.With("fontSize", size);

        if (RunReader.IsOn(effective.KeepNext)) node.With("keepNext", true);

        if (RunReader.IsOn(effective.KeepLines)) node.With("keepLines", true);

        // On when the file is silent, as in Word: only off is stated.
        if (effective.WidowControl is { } widow && !RunReader.IsOn(widow)) node.With("widowControl", false);
    }

    /// <summary>
    /// Only **direct** formatting; the inherited part comes from the style CSS.
    /// </summary>
    /// <remarks>
    /// Does the file declare the property in <c>w:pPr</c> (or in the mark, for the font)? The block
    /// carries its **effective** value, as when flattening; if the file is silent, so is the block.
    /// A declared zero is a declaration. Two exceptions: tab alignment (<see
    /// cref="TabAlignmentOf"/>) and line spacing with the direct mark font, which is measured
    /// against that font.
    /// </remarks>
    private void DirectOnly(
        Node node,
        string? viaTabs,
        ParagraphProperties effective,
        ParagraphProperties? direct,
        RunProperties inheritedRun)
    {
        if (viaTabs is not null) node.With("textAlign", viaTabs);
        else if (direct?.Justification is not null && AlignmentOf(effective) is { } alignment)
        {
            node.With("textAlign", alignment);
        }

        // Zero for the same reason as flattening.
        node.With("indent", 0);

        // A negative indent comes out as zero: the block does not draw outside the margin.
        var indentation = direct?.Indentation;
        if (indentation?.Left is not null) Declared(node, "indentMm", effective.Indentation?.Left?.Value);
        if (indentation?.Right is not null) Declared(node, "indentRightMm", effective.Indentation?.Right?.Value);
        if (indentation?.FirstLine is not null || indentation?.Hanging is not null)
        {
            var firstLine = TwipsToMm(effective.Indentation?.FirstLine?.Value);
            var hanging = TwipsToMm(effective.Indentation?.Hanging?.Value);
            node.With("firstLineMm", firstLine is > 0 ? firstLine.Value : hanging is > 0 ? -hanging.Value : 0);
        }

        if (direct?.Shading is not null)
        {
            // A colorless `w:shd` over a style with a background erases the background: otherwise
            // the rule would paint it.
            if (ShadingOf(effective) is { } background) node.With("background", background);
            else if (ShadingOf(_styles.StyleParagraphOf(direct)) is not null) node.With("background", "transparent");
        }

        var spacing = effective.SpacingBetweenLines;
        var declared = direct?.SpacingBetweenLines;
        if (declared?.Before is not null) node.With("spaceBefore", TwipsToPt(spacing?.Before?.Value) ?? 0);
        if (declared?.After is not null) node.With("spaceAfter", TwipsToPt(spacing?.After?.Value) ?? 0);

        var mark = _styles.ResolveMark(inheritedRun, direct);
        var lineHeight = LineHeightOf(spacing, LineMetrics.Of(mark.RunFonts?.Ascii?.Value), 4);
        if (declared?.Line is not null ||
            lineHeight != LineHeightOf(spacing, LineMetrics.Of(inheritedRun.RunFonts?.Ascii?.Value), 4))
        {
            node.With("lineHeight", lineHeight);
        }

        var markDirect = direct?.ParagraphMarkRunProperties;
        if (markDirect?.GetFirstChild<RunFonts>() is { } fonts &&
            (fonts.Ascii is not null || fonts.HighAnsi is not null) &&
            FontOf(mark) is { } font)
        {
            node.With("fontFamily", font);
        }

        if (markDirect?.GetFirstChild<FontSize>() is not null && FontSizeOf(mark) is { } size)
        {
            node.With("fontSize", size);
        }

        // `w:keepNext w:val="0"` exists to undo the style's.
        if (direct?.KeepNext is not null) node.With("keepNext", RunReader.IsOn(effective.KeepNext));
        if (direct?.KeepLines is not null) node.With("keepLines", RunReader.IsOn(effective.KeepLines));
        if (direct?.WidowControl is not null) node.With("widowControl", RunReader.IsOn(effective.WidowControl));
    }

    /// <summary>A declared measure: zero counts, negative becomes zero.</summary>
    private static void Declared(Node node, string name, string? twips)
    {
        if (TwipsToMm(twips) is { } value) node.With(name, Math.Max(0, value));
    }

    private string? FontOf(RunProperties properties)
    {
        var font = properties.RunFonts?.Ascii?.Value ?? properties.RunFonts?.HighAnsi?.Value;
        return string.IsNullOrWhiteSpace(font) ? null : _fonts.Stack(font);
    }

    private static string? FontSizeOf(RunProperties properties)
    {
        var value = properties.FontSize?.Val?.Value;
        if (!double.TryParse(value, out var halfPoints) || halfPoints <= 0) return null;
        var points = Unit.HalfPointsToPoints(halfPoints);
        return points == Math.Floor(points)
            ? $"{(int)points}pt"
            : points.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "pt";
    }

    private static string? ShadingOf(ParagraphProperties properties)
    {
        var fill = properties.Shading?.Fill?.Value;
        if (string.IsNullOrWhiteSpace(fill)) return null;
        if (fill.Equals("auto", StringComparison.OrdinalIgnoreCase)) return null;
        // "FFFFFF" is real white; only "auto" means "no color".
        return "#" + fill.TrimStart('#').ToLowerInvariant();
    }

    /// <summary>
    /// Twips → points. An **explicit** zero means "no space before", not absence.
    /// </summary>
    private static double? TwipsToPt(string? twips) =>
        int.TryParse(twips, out var value) && value >= 0 ? Math.Round(Unit.TwipsToPoints(value), 1) : null;

    /// <summary>
    /// Line height in CSS. <c>w:line</c> with the <c>auto</c> rule comes in 240ths of **times the
    /// natural height**, not of the font size. A number comes out whenever the font is known,
    /// single spacing included, because Chromium rounds <c>normal</c> to a whole pixel; with an
    /// unknown font, <c>normal</c>. <c>exact</c> and <c>atLeast</c> become points.
    /// </summary>
    /// <param name="decimals">
    /// Two on a flattened block, as always; four on one carrying only direct formatting, the 240ths
    /// grid in which <see cref="StyleReader"/> also delivers the style's.
    /// </param>
    private static string LineHeightOf(SpacingBetweenLines? spacing, double? natural, int decimals = 2)
    {
        var rule = spacing?.LineRule?.Value;
        var declared = spacing?.Line?.Value;

        if (declared is not null && int.TryParse(declared, out var value) && value > 0)
        {
            if (rule is not null && rule != LineSpacingRuleValues.Auto)
            {
                return Math.Round(Unit.TwipsToPoints(value), 1)
                    .ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "pt";
            }

            var factor = Math.Round(value / 240.0, decimals);

            // Outside this range it is file garbage.
            if (factor is > 0.5 and < 4) return Multiple(factor, natural);
        }

        return Multiple(1, natural);
    }

    /// <summary>
    /// With an unknown font and a declared multiple, the 1.15 guess: the height of almost every
    /// Latin font and of LibreOffice's substitutes.
    /// </summary>
    private static string Multiple(double factor, double? natural)
    {
        if (natural is not null) return Text(Math.Round(factor * natural.Value, 4));
        return factor == 1 ? "normal" : Text(Math.Round(factor * 1.1499, 4));
    }

    private static string Text(double value) =>
        value.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);

    /// <c>Heading1</c>, <c>heading 1</c> and LibreOffice's <c>Ttulo1</c>, without the accent
    /// because style ids do not accept it.
    private static int? HeadingLevelOf(ParagraphProperties? properties) =>
        HeadingLevelOfStyle(properties?.ParagraphStyleId?.Val?.Value);

    /// <summary>
    /// So <see cref="ParagraphFormat"/> does not replace the document's <c>Ttulo1</c> with a
    /// <c>Heading1</c>.
    /// </summary>
    internal static int? HeadingLevelOfStyle(string? style)
    {
        if (string.IsNullOrEmpty(style)) return null;

        var normalized = style.Replace(" ", string.Empty).ToLowerInvariant();
        foreach (var prefix in (string[])["heading", "ttulo", "titulo"])
        {
            if (normalized.StartsWith(prefix, StringComparison.Ordinal) &&
                int.TryParse(normalized[prefix.Length..], out var level) &&
                level is >= 1 and <= 6)
            {
                return level;
            }
        }

        return null;
    }

    /// <summary>
    /// In the corpus, the first heading comes with left <c>w:jc</c>, tabs and a centered stop: in
    /// HTML the tab would collapse. Only fires when the line **starts** with a tab.
    /// </summary>
    private static string? TabAlignmentOf(ParagraphProperties properties, List<Node> content)
    {
        if (content.Count == 0 || content[0].Type != "text" || content[0].Text != "\t") return null;

        var stops = properties.Tabs?.Elements<TabStop>().ToList();
        if (stops is null || stops.Count == 0) return null;

        if (stops.Any(stop => stop.Val is not null && stop.Val.Value == TabStopValues.Center)) return "center";
        if (stops.Any(stop => stop.Val is not null && stop.Val.Value == TabStopValues.Right)) return "right";
        return null;
    }

    private static string? AlignmentOf(ParagraphProperties? properties)
    {
        var value = properties?.Justification?.Val;
        if (value is null) return null;

        if (value == JustificationValues.Center) return "center";
        if (value == JustificationValues.Right) return "right";
        if (value == JustificationValues.Both || value == JustificationValues.Distribute) return "justify";
        return "left";
    }

    private Node WithFloats(Node node)
    {
        if (_paragraphFloats.Count > 0)
        {
            node.With("floats", JsonSerializer.SerializeToNode(_paragraphFloats, DocxJson.Options));
        }

        return node;
    }

    private static void Measure(Node node, string name, string? twips)
    {
        if (TwipsToMm(twips) is { } value and > 0) node.With(name, value);
    }

    private static double? TwipsToMm(string? twips) =>
        int.TryParse(twips, out var value) ? Math.Round(Unit.TwipsToMillimeters(value), 2) : null;


    private List<Node> ReadInline(
        OpenXmlElement container,
        RunProperties inherited,
        string? hyperlink = null)
    {
        var nodes = new List<Node>();
        var children = container.ChildElements.ToList();

        for (var index = 0; index < children.Count; index++)
        {
            var element = children[index];

            // The whole field becomes a node (ReadField); what does not fit goes run by run to the
            // inventory.
            if (references && element is Run begin && IsFieldBegin(begin) &&
                ReadField(children, index, inherited, hyperlink) is { } field)
            {
                nodes.Add(field.Node);
                index = field.Last;
                continue;
            }

            // Its own mark (`w:customMarkFollows`) in the next run is the mark, not text.
            if (notes && _textBoxDepth == 0 && element is Run noted && NoteReferenceOf(noted) is { } noteReference)
            {
                var read = ReadRun(noted, inherited, hyperlink).ToList();
                if (read.LastOrDefault(node => node.Type == "noteRef") is { } noteRef &&
                    IsCustomMark(noteReference) && noteRef.Attrs?.ContainsKey("mark") != true &&
                    index + 1 < children.Count && children[index + 1] is Run follower &&
                    string.Concat(follower.Elements<Text>().Select(text => text.Text)) is { Length: > 0 } mark)
                {
                    noteRef.With("mark", mark);
                    index++;
                }

                nodes.AddRange(read);
                continue;
            }

            if (references && element is SimpleField simple &&
                ReadSimpleField(simple, inherited, hyperlink) is { } plain)
            {
                nodes.Add(plain);
                continue;
            }

            switch (element)
            {
                // The ends become zero-width nodes, like bookmark ends; replies do not
                // (`_replies`).
                case CommentRangeStart start when comments && _textBoxDepth == 0:
                    if (start.Id?.Value is { } startId && !_replies.Contains(startId))
                    {
                        nodes.Add(Node.Of("commentStart").With("cid", startId));
                    }

                    break;

                case CommentRangeEnd end when comments && _textBoxDepth == 0:
                    if (end.Id?.Value is { } endId && !_replies.Contains(endId))
                    {
                        nodes.Add(Node.Of("commentEnd").With("cid", endId));
                    }

                    break;

                // Saving rebuilds the reference run with the `commentEnd`; only a point comment
                // gets its end here.
                case Run reference when comments && ReferenceOnly(reference) is { } referenced:
                    if (_textBoxDepth == 0 && !_replies.Contains(referenced) && !_rangeEnds.Contains(referenced))
                    {
                        nodes.Add(Node.Of("commentEnd").With("cid", referenced));
                    }

                    break;

                case Run run:
                    nodes.AddRange(ReadRun(run, inherited, hyperlink));
                    break;

                case Hyperlink link:
                    nodes.AddRange(ReadInline(link, inherited, HyperlinkTargetOf(link) ?? hyperlink));
                    break;

                // The ends become zero-width nodes and survive editing the paragraph; hidden ones
                // too, because Word cites them.
                case BookmarkStart start when references:
                    nodes.Add(Node.Of("bookmarkStart")
                        .With("name", start.Name?.Value ?? string.Empty)
                        .With("bid", start.Id?.Value ?? string.Empty));
                    break;

                case BookmarkEnd end when references:
                    nodes.Add(Node.Of("bookmarkEnd").With("bid", end.Id?.Value ?? string.Empty));
                    break;

                // A simple field the node does not represent (ReadSimpleField).
                case SimpleField:
                    inventory.NoteInvisible(Inventory.Fields);
                    _unrepresentedField = true;
                    break;

                case ParagraphProperties:
                case BookmarkStart:
                case BookmarkEnd:
                case ProofError:
                    break;

                // Those in a text box and in a draft older than comments go back through the
                // original XML: invisibility, not loss.
                case CommentRangeStart:
                case CommentRangeEnd:
                    inventory.NoteInvisible(Inventory.Comments);
                    break;

                // The range carries the revision mark, in any order with the link.
                case InsertedRun or DeletedRun or MoveFromRun or MoveToRun when revisions:
                    _revision.Add(Revisions.MarkOf(element, MoveNameOf(element)));
                    nodes.AddRange(ReadInline(element, inherited, hyperlink));
                    _revision.RemoveAt(_revision.Count - 1);
                    break;

                // A draft older than revisions (`BeforeRevisions`) reads insertions as plain text.
                case InsertedRun inserted:
                    nodes.AddRange(ReadInline(inserted, inherited, hyperlink));
                    break;

                case DeletedRun:
                    break;

                // The OMML goes back to the file; the MathML is what the screen draws.
                case OfficeMath or MathParagraph when math && _textBoxDepth == 0:
                    nodes.Add(ReadMath(element));
                    break;

                // One inside a box goes back with the box XML.
                case OfficeMath or MathParagraph when math:
                    inventory.NoteInvisible(Inventory.Equations);
                    break;

                default:
                    inventory.NoteInvisibleElement(element.LocalName);
                    break;
            }
        }

        return nodes;
    }

    /// <c>omml</c> is the identity: the fingerprint only sees it, because the MathML changes when
    /// the conversion improves. What the conversion does not draw is invisibility, not loss.
    private Node ReadMath(OpenXmlElement element)
    {
        var omml = element.OuterXml;
        var converted = OmmlMath.Convert(omml);
        if (converted.Lossy.Count > 0) inventory.NoteInvisible(Inventory.Equations);

        var node = new Node { Type = "math", Marks = RevisionMarks() }
            .With("omml", omml)
            .With("mathml", converted.MathMl)
            .With("latex", string.Empty)
            .With("display", converted.Display)
            .With("lossy", new JsonArray([.. converted.Lossy.Select(name => (JsonNode?)JsonValue.Create(name))]))
            .With("editable", converted.Lossy.Count == 0);
        if (converted.Jc is { } jc) node.With("jc", jc);
        return node;
    }

    /// <summary>Null when the run carries more than the reference.</summary>
    internal static string? ReferenceOnly(Run run)
    {
        string? id = null;
        foreach (var child in run.ChildElements)
        {
            if (child is RunProperties) continue;
            if (child is not CommentReference reference || id is not null) return null;
            id = reference.Id?.Value ?? string.Empty;
        }

        return id;
    }

    internal static OpenXmlElement? NoteReferenceOf(Run run) =>
        run.ChildElements.FirstOrDefault(child => child is FootnoteReference or EndnoteReference);

    /// <c>w:customMarkFollows</c>: the mark is the text that follows.
    internal static bool IsCustomMark(OpenXmlElement reference) =>
        (reference as FootnoteEndnoteReferenceType)?.CustomMarkFollows?.Value == true;

    private Node ReadNoteRef(OpenXmlElement reference)
    {
        var endnote = reference is EndnoteReference;
        var id = (reference as FootnoteEndnoteReferenceType)?.Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var node = Node.Of("noteRef")
            .With("kind", endnote ? NotesWriter.Endnote : NotesWriter.Footnote)
            .With("nid", id);
        node.Marks = RevisionMarks();

        var source = id is null ? null : NotesWriter.NoteOf(part, endnote, id);
        if (source is null)
        {
            node.Content = [Node.Of("paragraph")];
            return node;
        }

        var owner = endnote ? (OpenXmlPart)part.EndnotesPart! : part.FootnotesPart!;
        var address = NotesWriter.Address(endnote, id!);
        _noteReader ??= new BodyReader(part, inventory, flatten, references, sections: false, comments, revisions, notes: false, math);
        var (content, blocks) = _noteReader.ReadNote(source, address, owner);
        Notes[address] = new NoteRead(source, blocks);
        node.Content = content;
        return node;
    }

    /// <summary>
    /// It lives in <c>w:moveFromRangeStart</c>/<c>w:moveToRangeStart</c>, not in the run.
    /// </summary>
    private string? MoveNameOf(OpenXmlElement revision)
    {
        if (revision is not (MoveFromRun or MoveToRun)) return null;
        _moveNames ??= MoveNamesOf(part);
        return _moveNames.GetValueOrDefault(revision);
    }

    private static Dictionary<OpenXmlElement, string> MoveNamesOf(MainDocumentPart part)
    {
        var names = new Dictionary<OpenXmlElement, string>(ReferenceEqualityComparer.Instance);
        var open = new List<(string Id, string Name, bool From)>();
        foreach (var element in part.Document?.Descendants() ?? [])
        {
            switch (element)
            {
                case MoveFromRangeStart start:
                    open.Add((start.Id?.Value ?? string.Empty, start.Name?.Value ?? string.Empty, true));
                    break;
                case MoveToRangeStart start:
                    open.Add((start.Id?.Value ?? string.Empty, start.Name?.Value ?? string.Empty, false));
                    break;
                case MoveFromRangeEnd end:
                    open.RemoveAll(range => range.From && range.Id == end.Id?.Value);
                    break;
                case MoveToRangeEnd end:
                    open.RemoveAll(range => !range.From && range.Id == end.Id?.Value);
                    break;
                case MoveFromRun or MoveToRun:
                    var from = element is MoveFromRun;
                    var range = open.LastOrDefault(candidate => candidate.From == from);
                    if (range.Name is { Length: > 0 } name) names[element] = name;
                    break;
            }
        }

        return names;
    }

    /// <summary>For a node that is not text, like a line break.</summary>
    private List<Mark>? RevisionMarks() => _revision.Count == 0 ? null : [.. _revision];

    private List<Mark>? MarksOfRun(Run run, RunProperties inherited, string? hyperlink)
    {
        // The inherited part goes along so a direct "off" becomes a mark; old drafts do not know
        // them.
        var marks = RunReader.MarksOf(
            _styles.ResolveRun(inherited, run.RunProperties),
            hyperlink,
            _fonts,
            flatten ? null : inherited,
            directOnly: _directRuns);

        // `w:rStyle`: the style CSS draws it, and the mark brings it back to the file.
        if (!flatten && run.RunProperties?.RunStyle?.Val?.Value is { Length: > 0 } characterStyle)
        {
            (marks ??= []).Add(Mark.Of("charStyle", "styleId", characterStyle));
        }

        if (_revision.Count > 0) (marks ??= []).AddRange(_revision);

        return marks;
    }


    private sealed record TableOfContentsRead(Node Node, List<OpenXmlElement> Continuation);

    /// <summary>
    /// A "Table of Contents" <c>w:sdt</c>, as Word and LibreOffice write it, or the paragraphs
    /// themselves, when the <c>TOC</c> field opens in the first and closes in one of those below
    /// (<see cref="Block.Continuation"/>).
    /// </summary>
    private TableOfContentsRead? TableOfContentsAt(List<OpenXmlElement> elements, int at)
    {
        if (elements[at] is SdtBlock sdt)
        {
            var inside = sdt.SdtContentBlock?.ChildElements.ToList() ?? [];
            if (inside.Count == 0 || inside.Any(child => child is not Paragraph)) return null;
            var node = ReadTableOfContents([.. inside.Cast<Paragraph>()], sdt: true);
            return node is null ? null : new TableOfContentsRead(node, []);
        }

        if (elements[at] is not Paragraph first || TocBegin(first) is null) return null;

        // Up to the paragraph where the field closes; a table of contents does not cross a table.
        var depth = 0;
        for (var index = at; index < elements.Count; index++)
        {
            if (elements[index] is not Paragraph paragraph) return null;
            foreach (var mark in paragraph.Descendants<FieldChar>())
            {
                if (mark.FieldCharType?.Value == FieldCharValues.Begin) depth++;
                else if (mark.FieldCharType?.Value == FieldCharValues.End) depth--;
            }

            if (depth > 0) continue;

            var group = elements.Skip(at).Take(index - at + 1).Cast<Paragraph>().ToList();
            var node = ReadTableOfContents(group, sdt: false);
            return node is null ? null : new TableOfContentsRead(node, [.. group.Skip(1)]);
        }

        return null;
    }

    /// <summary>
    /// The index of the run opening the top-level <c>TOC</c>, the separator's, and the instruction.
    /// </summary>
    private static (int Begin, int Separate, string Instruction)? TocBegin(Paragraph paragraph)
    {
        var children = paragraph.ChildElements.ToList();
        var begin = children.FindIndex(child => child is Run run && IsFieldBegin(run));
        if (begin < 0) return null;

        var instruction = new System.Text.StringBuilder();
        for (var index = begin + 1; index < children.Count; index++)
        {
            if (children[index] is not Run run) return null;
            if (FieldCharOf(run) is { } mark)
            {
                if (mark.FieldCharType?.Value != FieldCharValues.Separate) return null;
                var text = instruction.ToString();
                return text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase)
                    ? (begin, index, text)
                    : null;
            }

            foreach (var code in run.Elements<FieldCode>()) instruction.Append(code.Text);
        }

        return null;
    }

    /// <summary>
    /// The <c>TOC</c> field leaves the paragraphs and goes to the node; what remains are plain
    /// paragraphs, editable by hand as in Word. <c>head</c> counts the paragraphs before the field.
    /// A field that does not close in the last paragraph is not represented.
    /// </summary>
    private Node? ReadTableOfContents(List<Paragraph> paragraphs, bool sdt)
    {
        var head = paragraphs.FindIndex(paragraph => paragraph.Descendants<FieldChar>().Any());
        if (head < 0 || TocBegin(paragraphs[head]) is not { } begin) return null;

        // Counting the PAGEREFs inside.
        var depth = 0;
        FieldChar? end = null;
        var endParagraph = -1;
        for (var index = head; index < paragraphs.Count && end is null; index++)
        {
            foreach (var mark in paragraphs[index].Descendants<FieldChar>())
            {
                if (mark.FieldCharType?.Value == FieldCharValues.Begin) depth++;
                else if (mark.FieldCharType?.Value == FieldCharValues.End && --depth == 0)
                {
                    end = mark;
                    endParagraph = index;
                    break;
                }
            }
        }

        if (end is null || endParagraph != paragraphs.Count - 1) return null;
        if (end.Parent is not Run endRun || endRun.Parent != paragraphs[endParagraph] || FieldCharOf(endRun) is null)
        {
            return null;
        }

        var endIndex = paragraphs[endParagraph].ChildElements.ToList().IndexOf(endRun);
        var clones = paragraphs.Select(paragraph => (Paragraph)paragraph.CloneNode(true)).ToList();

        // The end before the start, so the index holds for a single-paragraph table of contents.
        clones[endParagraph].ChildElements[endIndex].Remove();
        var opening = clones[head].ChildElements.Skip(begin.Begin).Take(begin.Separate - begin.Begin + 1).ToList();
        foreach (var piece in opening) piece.Remove();

        var content = new List<Node>();
        foreach (var clone in clones)
        {
            var node = ReadParagraph(clone, flat: flatten);
            if (node.Type is not ("paragraph" or "heading")) return null;
            content.Add(node);
        }

        var toc = Node.Of("tableOfContents")
            .With("instr", begin.Instruction)
            .With("head", head)
            .With("sdt", sdt);
        toc.Content = content;
        return toc;
    }


    /// <summary>Only the opening <c>w:fldChar</c>.</summary>
    private static bool IsFieldBegin(Run run) =>
        FieldCharOf(run) is { } mark && mark.FieldCharType?.Value == FieldCharValues.Begin;

    /// <summary>When it is the only thing the run carries.</summary>
    private static FieldChar? FieldCharOf(Run run)
    {
        var content = run.ChildElements.Where(child => child is not RunProperties).ToList();
        return content is [FieldChar only] ? only : null;
    }

    private sealed record ReadFieldResult(Node Node, int Last);

    /// <summary>
    /// A whole complex field as a <c>field</c> node, to survive editing the paragraph. Only the
    /// lossless case: one container, one run per piece, a text-and-tab result, no nested field.
    /// Otherwise returns null, and the field is read by its result. The formatting is the result's
    /// first run's.
    /// </summary>
    private ReadFieldResult? ReadField(List<OpenXmlElement> siblings, int start, RunProperties inherited, string? hyperlink)
    {
        // What the node does not carry (`w:ffData`, `w:fldData`, `w:fldLock`, `w:dirty`) stays out
        // of it: rewritten from the node, the field would come back without them.
        if (FieldCharOf((Run)siblings[start]) is { } opening &&
            (opening.HasChildren || opening.FieldLock is not null || opening.Dirty is not null))
        {
            return null;
        }

        var instruction = new System.Text.StringBuilder();
        var result = new System.Text.StringBuilder();
        Run? formatted = null;
        var separated = false;

        for (var index = start + 1; index < siblings.Count; index++)
        {
            if (siblings[index] is ProofError) continue;
            if (siblings[index] is not Run run) return null;

            foreach (var child in run.ChildElements)
            {
                switch (child)
                {
                    case RunProperties:
                    case LastRenderedPageBreak:
                        break;

                    case FieldChar mark when mark.FieldCharType?.Value == FieldCharValues.Separate && !separated:
                        separated = true;
                        break;

                    case FieldChar mark when mark.FieldCharType?.Value == FieldCharValues.End:
                        if (FieldCharOf(run) is null) return null;
                        var node = Node.Of("field")
                            .With("instr", instruction.ToString())
                            .With("result", result.ToString());
                        node.Marks = MarksOfRun(formatted ?? (Run)siblings[start], inherited, hyperlink);
                        if (result.Length > 0) _paragraphHasContent = true;
                        return new ReadFieldResult(node, index);

                    case FieldCode code when !separated:
                        instruction.Append(code.Text);
                        break;

                    case DeletedFieldCode code when !separated && revisions:
                        instruction.Append(code.Text);
                        break;

                    case Text text when separated:
                        result.Append(text.Text);
                        formatted ??= run;
                        break;

                    case DeletedText text when separated && revisions:
                        result.Append(text.Text);
                        formatted ??= run;
                        break;

                    case TabChar when separated:
                        result.Append('\t');
                        formatted ??= run;
                        break;

                    default:
                        return null;
                }
            }
        }

        return null;
    }

    /// <c>w:fldSimple</c>: the same node, which goes back in complex form, the same to Word.
    private Node? ReadSimpleField(SimpleField field, RunProperties inherited, string? hyperlink)
    {
        // For the same reason as ReadField.
        if (field.FieldLock is not null || field.Dirty is not null || field.GetFirstChild<FieldData>() is not null)
        {
            return null;
        }

        var result = new System.Text.StringBuilder();
        Run? formatted = null;

        foreach (var child in field.ChildElements)
        {
            if (child is not Run run) return null;
            foreach (var piece in run.ChildElements)
            {
                switch (piece)
                {
                    case RunProperties: break;
                    case Text text: result.Append(text.Text); break;
                    case TabChar: result.Append('\t'); break;
                    default: return null;
                }
            }

            formatted ??= run;
        }

        var node = Node.Of("field")
            .With("instr", field.Instruction?.Value ?? string.Empty)
            .With("result", result.ToString());
        node.Marks = formatted is null ? null : MarksOfRun(formatted, inherited, hyperlink);
        if (result.Length > 0) _paragraphHasContent = true;
        return node;
    }

    private IEnumerable<Node> ReadRun(Run run, RunProperties inherited, string? hyperlink)
    {
        var marks = MarksOfRun(run, inherited, hyperlink);

        Node? customMark = null;

        foreach (var element in run.ChildElements)
        {
            switch (element)
            {
                case Text text when customMark is not null:
                    customMark.With("mark", text.Text);
                    customMark = null;
                    break;

                // The number is the document order; the note body goes inside the node.
                case FootnoteReference or EndnoteReference when notes && _textBoxDepth == 0:
                {
                    var noteRef = ReadNoteRef(element);
                    if (IsCustomMark(element)) customMark = noteRef;
                    _paragraphHasContent = true;
                    yield return noteRef;
                    break;
                }

                // Word draws the body's number, and saving rebuilds it (NotesWriter).
                case FootnoteReferenceMark or EndnoteReferenceMark when _oidPrefix is not null:
                    break;

                case Text text:
                    if (text.Text.Length > 0)
                    {
                        _paragraphHasContent = true;
                        yield return new Node { Type = "text", Text = text.Text, Marks = marks };
                    }

                    break;

                case DeletedText deleted when revisions:
                    if (deleted.Text.Length > 0)
                    {
                        _paragraphHasContent = true;
                        yield return new Node { Type = "text", Text = deleted.Text, Marks = marks };
                    }

                    break;

                case TabChar:
                    yield return new Node { Type = "text", Text = "\t", Marks = marks };
                    break;

                case Break br:
                    // A column break becomes a block property, like a page break; in drafts older
                    // than sections it was a line break.
                    if (br.Type is not null && br.Type.Value == BreakValues.Page)
                    {
                        yield return Node.Of("pageBreak");
                    }
                    else if (br.Type is not null && br.Type.Value == BreakValues.Column && sections)
                    {
                        yield return Node.Of("columnBreak");
                    }
                    else
                    {
                        var hardBreak = Node.Of("hardBreak");
                        hardBreak.Marks = RevisionMarks();
                        yield return hardBreak;
                    }

                    break;

                case DocumentFormat.OpenXml.Wordprocessing.Drawing:
                    foreach (var node in ReadShape(element)) yield return node;
                    break;

                case Picture picture:
                    // VML: a Word 2003 document carries images this way in the body.
                    foreach (var node in ReadShape(picture)) yield return node;
                    break;

                case AlternateContent alternate:
                    // `mc:Choice` and `mc:Fallback` are the same content: only one branch is read,
                    // or each box would appear twice.
                    var branch = (OpenXmlElement?)alternate.GetFirstChild<AlternateContentChoice>()
                                 ?? alternate.GetFirstChild<AlternateContentFallback>();
                    if (branch is not null)
                    {
                        foreach (var node in ReadShape(branch)) yield return node;
                    }

                    break;

                case RunProperties:
                case LastRenderedPageBreak:
                    break;

                // A reference sharing its run with text: rare, and saving declares it.
                case CommentReference when comments:
                    break;

                case FieldChar:
                case FieldCode:
                    // A field the node does not represent: the cached result comes in the next run,
                    // and rewritten it would become plain text.
                    inventory.NoteInvisible(Inventory.Fields);
                    _unrepresentedField = true;
                    break;

                default:
                    inventory.NoteInvisibleElement(element.LocalName);
                    break;
            }
        }
    }


    /// <summary>
    /// The image and the text inside the shape. A text box is content: the manual template cover
    /// keeps its title in one. Without an anchor, the text enters the line the shape is on, the
    /// best approximation without paginating.
    /// </summary>
    private IEnumerable<Node> ReadShape(OpenXmlElement shape)
    {
        // Anchored objects leave the flow and become a property of the anchor paragraph, as in
        // Word; `wp:inline` stays in the flow. Except anchored ones LibreOffice places where the
        // flow already would; see AnchorReader.FlowsWithText.
        if (AnchorReader.AnchorOf(shape) is { } anchor && !AnchorReader.FlowsWithText(anchor))
        {
            foreach (var floating in DescribeAnchored(shape, anchor)) _paragraphFloats.Add(floating);
            yield break;
        }

        if (ReadImage(shape) is { } image)
        {
            _paragraphHasContent = true;
            yield return image;
        }

        // A box without an anchor has no position of its own: the text enters the line.
        foreach (var node in ReadTextBoxesInline(shape)) yield return node;
    }

    /// <summary>An image, a text box, or nothing.</summary>
    private IEnumerable<FloatDto> DescribeAnchored(
        OpenXmlElement shape,
        Drawing.Wordprocessing.Anchor anchor)
    {
        if (ImageSourceOf(shape) is { } src)
        {
            yield return AnchorReader.Describe(anchor, "image", src, null);
            yield break;
        }

        foreach (var box in OutermostTextBoxes(shape))
        {
            // Only what is not drawn goes to the inventory.
            var look = ShapeLook.Of(box);
            if (!look.Complete) inventory.NoteInvisible(Inventory.Shapes);

            var content = new List<Node>();
            foreach (var paragraph in box.Descendants<Paragraph>())
            {
                if (paragraph.Ancestors<TextBoxContent>().First() != box) continue;
                content.Add(ReadParagraphOf(paragraph));
            }

            if (content.Count > 0)
            {
                yield return AnchorReader.Describe(anchor, "text", null, content) with
                {
                    Fill = look.Fill,
                    Line = look.Line,
                    LineWidthPt = look.LineWidthPt,
                    Dash = look.Dashed,
                };
            }
        }
    }

    /// <summary>A box is its own flow: its paragraphs are paragraphs, not joined lines.</summary>
    private Node ReadParagraphOf(Paragraph paragraph)
    {
        var (effective, inheritedRun) = _styles.Resolve(paragraph.ParagraphProperties);

        var node = Node.Of("paragraph");
        if (AlignmentOf(effective) is { } alignment) node.With("textAlign", alignment);

        var outer = _directRuns;
        _directRuns = false;
        _textBoxDepth++;
        var content = ReadInline(paragraph, inheritedRun);
        _textBoxDepth--;
        _directRuns = outer;
        if (content.Count > 0) node.Content = content;
        return node;
    }

    /// <summary>Without an anchor, the content enters the line.</summary>
    private IEnumerable<Node> ReadTextBoxesInline(OpenXmlElement shape)
    {
        foreach (var box in OutermostTextBoxes(shape))
        {
            // The box is not drawn: its frame is out of sight, and the warning says so.
            var look = ShapeLook.Of(box);
            if (!look.Complete || look.Draws) inventory.NoteInvisible(Inventory.Shapes);

            foreach (var paragraph in box.Descendants<Paragraph>())
            {
                if (paragraph.Ancestors<TextBoxContent>().First() != box) continue;

                var (_, inheritedRun) = _styles.Resolve(paragraph.ParagraphProperties);
                var afterSomething = _paragraphHasContent;

                var outer = _directRuns;
                _directRuns = false;
                _textBoxDepth++;
                var inline = ReadInline(paragraph, inheritedRun);
                _textBoxDepth--;
                _directRuns = outer;
                if (inline.Count == 0) continue;

                if (afterSomething) yield return Node.Of("hardBreak");
                _paragraphHasContent = true;

                foreach (var node in inline) yield return node;
            }
        }
    }

    /// <summary>
    /// Stops at the first box of each branch: the inner one is reached by the recursion of <see
    /// cref="ReadInline"/>, and both routes would write the text twice.
    /// </summary>
    private static IEnumerable<TextBoxContent> OutermostTextBoxes(OpenXmlElement root) =>
        TextBoxNav.Outermost(root);

    /// <summary>
    /// Separate from <see cref="ReadImage"/>: an anchored object needs the bytes without the node.
    /// </summary>
    private string? ImageSourceOf(OpenXmlElement drawing)
    {
        var blip = drawing.Descendants<Drawing.Blip>().FirstOrDefault();
        var relationshipId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(relationshipId)) return null;

        if (!_owner.TryGetPartById(relationshipId, out var found) || found is not ImagePart image)
        {
            inventory.NoteLoss("imagem em formato não suportado");
            return null;
        }

        using var stream = image.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);

        return $"data:{image.ContentType};base64,{Convert.ToBase64String(buffer.ToArray())}";
    }

    /// <summary>
    /// A capture anchored to the top of the paragraph goes to its start, as Word and LibreOffice
    /// draw it; anchored to the line, it stays where it is.
    /// </summary>
    private void TopAnchoredFirst(List<Node> content)
    {
        if (_topAnchored.Count == 0) return;
        var moved = content.Where(_topAnchored.Contains).ToList();
        if (moved.Count == 0 || content.Take(moved.Count).SequenceEqual(moved)) return;
        content.RemoveAll(_topAnchored.Contains);
        content.InsertRange(0, moved);
    }

    /// <summary>
    /// An inline image (<c>wp:inline</c>) and an anchored one where the flow would already place
    /// it; see <see cref="AnchorReader.FlowsWithText"/>.
    /// </summary>
    private Node? ReadImage(OpenXmlElement drawing)
    {
        if (ImageSourceOf(drawing) is not { } src) return null;

        var node = Node.Of("image").With("src", src);

        // Anchored to the paragraph, even if in its flow place: the anchoring paragraph takes its
        // own line besides the image (11.55 pt measured in LibreOffice).
        if (AnchorReader.AnchorOf(drawing) is { } anchor)
        {
            node.With("anchored", true);
            var vertical = anchor.GetFirstChild<Drawing.Wordprocessing.VerticalPosition>();
            var from = vertical?.RelativeFrom?.Value;
            // An offset up to 1 pt (635 EMU in the corpus) is the zero LibreOffice writes.
            var offset = long.TryParse(vertical?.PositionOffset?.Text, out var emus) ? Math.Abs(emus) : 0;
            if ((from is null || from == Drawing.Wordprocessing.VerticalRelativePositionValues.Paragraph) &&
                offset <= Unit.EmusPerPoint)
            {
                _topAnchored.Add(node);
            }
        }

        // Alt text (`wp:docPr/@descr`) becomes `alt`; `descr=""` is the default of someone who
        // never filled it in, and does not become an attribute.
        var properties = drawing.Descendants<Drawing.Wordprocessing.DocProperties>().FirstOrDefault();
        if (properties?.Description?.Value is { Length: > 0 } description)
        {
            node.With("alt", description);
        }

        var extent = drawing.Descendants<Drawing.Wordprocessing.Extent>().FirstOrDefault();
        if (extent?.Cx?.Value is { } wide && extent.Cy?.Value is { } tall && wide > 0 && tall > 0)
        {
            // `wp:extent` measures before rotating: on a quarter turn, width and height swap.
            var (across, down) = IsQuarterTurned(drawing) ? (tall, wide) : (wide, tall);

            // Both measures: without the height, pagination would measure the sheet without the
            // image, and the ratio would be the file's.
            node.With("width", (int)Math.Round(Unit.EmusToPixels(across)));
            node.With("height", (int)Math.Round(Unit.EmusToPixels(down)));
        }

        return node;
    }

    /// <c>a:rot</c> comes in 60000ths of a degree, and may be negative. Only a quarter turn swaps
    /// width and height.
    internal static bool IsQuarterTurned(OpenXmlElement drawing)
    {
        var rotation = drawing.Descendants<Drawing.Transform2D>().FirstOrDefault()?.Rotation?.Value;
        if (rotation is null) return false;

        var degrees = ((rotation.Value / 60000.0) % 360 + 360) % 360;
        return degrees is (> 45 and < 135) or (> 225 and < 315);
    }

    private string? HyperlinkTargetOf(Hyperlink link)
    {
        var id = link.Id?.Value;

        // An internal link has `w:anchor`, not a relationship: in the editor it becomes `#name`.
        if (string.IsNullOrEmpty(id))
        {
            return references && link.Anchor?.Value is { Length: > 0 } anchor ? "#" + anchor : null;
        }

        try
        {
            return _owner.HyperlinkRelationships
                .FirstOrDefault(relationship => relationship.Id == id)?.Uri.ToString();
        }
        catch (UriFormatException)
        {
            return null;
        }
    }


    private Node ReadTable(Table table)
    {
        var rows = new List<Node>();

        // The table grid, which Word uses and the editor mirrors in `colgroup`.
        var grid = TableLook.GridWidths(table);

        foreach (var row in table.Elements<TableRow>())
        {
            var cells = new List<Node>();

            // `w:tblHeader` is the editor's header row; present without `w:val` already means
            // "yes".
            var repeat = row.TableRowProperties?.GetFirstChild<TableHeader>();
            var header = repeat is not null
                         && !(repeat.Val is { } declared && declared.Value == OnOffOnlyValues.Off);

            // `w:gridBefore`: a row may start some columns ahead.
            var column = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value is > 0 and var before
                ? before
                : 0;

            foreach (var cell in row.Elements<TableCell>())
            {
                // A nested table is a block, and the editor's `tableCell` accepts blocks.
                var contents = new List<Node>();
                foreach (var child in cell.ChildElements)
                {
                    switch (child)
                    {
                        case Paragraph paragraph: contents.Add(ReadParagraph(paragraph, flat: true)); break;
                        case Table nested: contents.Add(ReadTable(nested)); break;
                        case TableCellProperties: break;
                        default: inventory.NoteInvisibleElement(child.LocalName); break;
                    }
                }

                if (contents.Count == 0) contents.Add(Node.Of("paragraph"));

                var node = Node.Of(header ? "tableHeader" : "tableCell");
                node.Content = contents;

                // **Always** written: the editor returns `colspan` and `rowspan` on every cell, and
                // omitting them would regenerate every table.
                var span = cell.TableCellProperties?.GridSpan?.Val?.Value;
                var spanned = span is > 1 ? span.Value : 1;
                node.With("colspan", spanned);

                // Vertical merging stays in the file's `w:vMerge`.
                node.With("rowspan", 1);

                // One measure per column, in CSS pixels, as TableKit uses.
                if (grid.Count >= column + spanned)
                {
                    var widths = new JsonArray();
                    for (var index = 0; index < spanned; index++) widths.Add(grid[column + index]);
                    node.With("colwidth", widths);
                }

                column += spanned;

                // Only when the file declares them: absent and null are the same statement.
                if (TableLook.Shading(cell.TableCellProperties) is { } fill) node.With("shading", fill);
                if (TableLook.Borders(cell.TableCellProperties) is { } borders) node.With("borders", borders);

                cells.Add(node);
            }

            var rowNode = Node.Of("tableRow");
            rowNode.Content = cells;

            if (revisions && Revisions.BlockRevisionOf(row.TableRowProperties) is { } rowRevision)
            {
                rowNode.With("rowRevision", rowRevision);
            }

            rows.Add(rowNode);
        }

        var tableNode = Node.Of("table");
        tableNode.Content = rows;

        // So the screen measures the row like paper; the `w:tblPr` goes back as it was.
        tableNode.With("cellMargins", string.Join(' ', TableLook.CellMargins(table, part)));
        return tableNode;
    }

}
