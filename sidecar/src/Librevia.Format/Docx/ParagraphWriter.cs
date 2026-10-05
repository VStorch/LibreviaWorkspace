using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Editor node → OOXML, only for **edited** blocks: what this file cannot generate is a real loss.
/// </summary>
public sealed class ParagraphWriter
{
    private readonly MainDocumentPart _part;

    /// <summary>
    /// The document, or <c>footnotes.xml</c> in a note: a note link's <c>r:id</c> belongs to it.
    /// </summary>
    private readonly OpenXmlPart _owner;

    /// <summary>See NotesWriter.ReferenceRunsOf.</summary>
    private Dictionary<string, Run>? _noteRuns;

    private readonly Inventory _inventory;
    private readonly ParagraphFormat _format;
    private readonly StyleResolver _styles;

    /// <summary>See <see cref="DropWhatRepeatsTheStyle"/>.</summary>
    private RunProperties _paragraphRun = new();
    private readonly TableWriter _tables;
    private readonly ImageWriter _images;

    /// <summary>
    /// When off, in a draft older than bookmarks, the original paragraph's are copied.
    /// </summary>
    private readonly bool _references;

    /// <summary>
    /// Read at the first end, before the body is replaced; see <see cref="CommentAnchors"/>.
    /// </summary>
    private CommentAnchors? _comments;

    /// <summary>In twentieths of a point: 1 CSS px is 15.</summary>
    private readonly int _usableTwips;

    /// <param name="usableWidthPx">The ceiling, in CSS pixels, for an image arriving without a
    /// measure.</param>
    internal ParagraphWriter(
        MainDocumentPart part,
        Inventory inventory,
        int usableWidthPx = ImageWriter.DefaultWidthPx,
        HeadingStyles? headings = null,
        bool flatten = false,
        bool references = true,
        bool revisions = true,
        OpenXmlPart? owner = null)
    {
        _part = part;
        _owner = owner ?? part;
        _references = references;
        _inventory = inventory;
        _styles = new StyleResolver(part);
        _format = new ParagraphFormat(inventory, headings ?? new HeadingStyles(part, null), _styles, flatten, revisions);
        var usable = usableWidthPx > 0 ? usableWidthPx : ImageWriter.DefaultWidthPx;
        _tables = new TableWriter(inventory, (node, original) => Write(node, null, original), usable, revisions);
        _images = new ImageWriter(part, inventory, usable, _owner);
        _usableTwips = usable * Unit.TwipsPerPixel;
    }

    public IEnumerable<OpenXmlElement> Write(
        Node node,
        ListPlacement? list = null,
        OpenXmlElement? original = null)
    {
        switch (node.Type)
        {
            case "paragraph":
            case "heading":
            {
                var paragraph = WriteParagraph(node, list, original as Paragraph);
                CarryAnchored(paragraph, node, original);
                NoteEditedMathFormatting(node, original);
                yield return paragraph;
                break;
            }

            case "pageBreak":
                yield return new Paragraph(new Run(new Break { Type = BreakValues.Page }));
                break;

            case "table":
                yield return _tables.Write(node, original as Table);
                break;

            // An image inserted from the toolbar is an editor block, and in OOXML it travels inside
            // a paragraph.
            case "image":
            {
                var image = _images.Write(node);
                if (image is null)
                {
                    yield return new Paragraph();
                    break;
                }

                // The paragraph's `w:jc` aligns: OOXML has no centered image.
                var aligned = new Paragraph(image);
                if (ParagraphFormat.JustificationOf(Attr.String(node, "align")) is { } justification)
                {
                    aligned.ParagraphProperties = new ParagraphProperties(justification);
                }

                yield return aligned;
                break;
            }

            case "tableOfContents":
                foreach (var element in WriteTableOfContents(node, original)) yield return element;
                break;

            case "horizontalRule":
                yield return new Paragraph(new ParagraphProperties(
                    new ParagraphBorders(new BottomBorder
                    {
                        Val = BorderValues.Single,
                        Size = 6,
                        Color = "auto",
                    })));
                break;

            default:
                // What we cannot generate becomes an empty paragraph and goes to the inventory,
                // never silently.
                _inventory.NoteLoss($"bloco do tipo \"{node.Type}\"");
                yield return new Paragraph();
                break;
        }
    }

    /// <summary>Numbering inherited from the document, for list items.</summary>
    /// <param name="Kind">A sublist only inherits the outer numbering when it is the same
    /// kind.</param>
    public sealed record ListContext(string Kind, int NumberingId, int Level);

    /// <summary>
    /// Three states: the body knows the paragraph **is** an item (the context), that it **is not**
    /// (<c>null</c>, and the original <c>w:numPr</c> goes), and a table **does not know** (no
    /// wrapper, and the <c>w:numPr</c> stays).
    /// </summary>
    public readonly record struct ListPlacement(ListContext? List);

    /// <summary>
    /// The original's anchored objects stay in the rewritten paragraph, as the surgical save does
    /// with a whole block: this writer does not generate them. Only a <c>w:r</c> that **is** the
    /// drawing; a run with text would repeat the sentence and goes to the inventory.
    /// </summary>
    private void CarryAnchored(Paragraph paragraph, Node node, OpenXmlElement? original)
    {
        if (original is null) return;

        // An anchored object flowing with the text comes back as a paragraph image; copied here, it
        // would appear twice.
        var flowing = ImageWriter.FlowingImagesOf(original);

        foreach (var run in original.Elements<Run>())
        {
            if (!IsAnchoredOnly(run)) continue;
            if (run.Elements<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any(flowing.Contains)) continue;
            paragraph.AppendChild(run.CloneNode(true));
        }

        ApplyBoxText(paragraph, node);
    }

    /// <summary>
    /// Text typed in a box goes back to the <c>w:txbxContent</c>. A box without changes is not
    /// touched. If the count does not match, nothing is written: text swapped between boxes is
    /// worse than lost text.
    /// </summary>
    private void ApplyBoxText(Paragraph paragraph, Node node)
    {
        var wanted = BoxContentsOf(node);
        if (wanted.Count == 0) return;

        var boxes = TextBoxNav.AnchoredBoxesOf(paragraph).ToList();
        if (boxes.Count != wanted.Count)
        {
            _inventory.NoteLoss("texto de caixa num parágrafo que você editou");
            return;
        }

        var touched = false;
        for (var index = 0; index < boxes.Count; index++)
        {
            var box = boxes[index];
            var content = wanted[index];
            if (TextBoxNav.TextOf(box) == PlainTextOf(content)) continue;

            box.RemoveAllChildren();
            foreach (var block in content)
            {
                foreach (var element in Write(block)) box.AppendChild(element);
            }

            // An empty `w:txbxContent` invalidates the document.
            if (!box.HasChildren) box.AppendChild(new Paragraph());
            touched = true;
        }

        if (!touched) return;

        foreach (var alternate in paragraph.Descendants<AlternateContent>().ToList())
        {
            TextBoxNav.MirrorFallback(alternate);
        }
    }

    private static List<List<Node>> BoxContentsOf(Node node)
    {
        var contents = new List<List<Node>>();
        if (node.Attrs is null ||
            !node.Attrs.TryGetValue("floats", out var value) ||
            value is not JsonArray floats)
        {
            return contents;
        }

        foreach (var item in floats)
        {
            if (item is not JsonObject float_) continue;
            if (float_["kind"]?.GetValue<string>() != "text") continue;

            var content = float_["content"].Deserialize<List<Node>>(DocxJson.Options);
            contents.Add(content ?? []);
        }

        return contents;
    }

    private static string PlainTextOf(List<Node> content) =>
        string.Join("\n", content.Select(TextOfNode));

    private static string TextOfNode(Node node) =>
        node.Text ?? string.Concat((node.Content ?? []).Select(TextOfNode));

    /// <summary>
    /// By direct children: a box's text is also <c>w:t</c>, and in depth it would reject exactly
    /// the runs to copy.
    /// </summary>
    internal static bool IsAnchoredOnly(Run run) =>
        run.Descendants<WordDrawing.Anchor>().Any() && !run.Elements<Text>().Any();

    private Paragraph WriteParagraph(Node node, ListPlacement? list, Paragraph? original)
    {
        var paragraph = new Paragraph();

        // The `w:pPr` starts from the original (ParagraphFormat): style, spacing, mark and
        // `w:sectPr` survive.
        if (_format.Build(node, list, original) is { } properties)
        {
            paragraph.ParagraphProperties = properties;
        }

        foreach (var mark in _references ? [] : Bookmarks(original, leading: true))
        {
            paragraph.AppendChild(mark.CloneNode(true));
        }

        _paragraphRun = _styles.Resolve(paragraph.ParagraphProperties).Run;

        // Images that were already there come back with the original drawing; see
        // ImageWriter.Reuse.
        var images = ImageWriter.FlowingImagesOf(original);
        foreach (var child in node.Content ?? [])
        {
            // See Revisions.Wrap.
            var written = WriteInline(child, images);
            foreach (var element in Revisions.HasRevision(child) ? Revisions.Wrap([.. written], child.Marks) : written)
            {
                paragraph.AppendChild(element);
            }
        }

        MergeAnchorLinks(paragraph);
        Revisions.MergeNeighbours(paragraph);

        // The break goes back to the end of the paragraph, in a `w:r`, where the reader took it
        // from.
        if (Attr.Bool(node, "breakAfter"))
        {
            paragraph.AppendChild(new Run(new Break { Type = BreakValues.Page }));
        }

        if (Attr.Bool(node, "columnBreakAfter"))
        {
            paragraph.AppendChild(new Run(new Break { Type = BreakValues.Column }));
        }

        foreach (var mark in _references ? [] : Bookmarks(original, leading: false))
        {
            paragraph.AppendChild(mark.CloneNode(true));
        }

        return paragraph;
    }

    /// <summary>
    /// Neighbouring links to the same bookmark become a single <c>w:hyperlink</c>, as Word writes a
    /// table of contents entry.
    /// </summary>
    private static void MergeAnchorLinks(Paragraph paragraph)
    {
        Hyperlink? previous = null;
        foreach (var child in paragraph.ChildElements.ToList())
        {
            if (child is Hyperlink link && link.Id is null && link.Anchor?.Value is { } anchor &&
                previous?.Anchor?.Value == anchor && previous.Id is null)
            {
                foreach (var inner in link.ChildElements.ToList())
                {
                    inner.Remove();
                    previous.AppendChild(inner);
                }

                link.Remove();
                continue;
            }

            previous = child as Hyperlink;
        }
    }

    /// <summary>
    /// The original's bookmarks, which the model does not represent: without them, whatever cites
    /// them would point to nothing. On the side they were on, so the marked range does not shrink;
    /// the position mid-sentence is what is lost.
    /// </summary>
    /// <param name="leading">Those that came before any content.</param>
    private static IEnumerable<OpenXmlElement> Bookmarks(Paragraph? original, bool leading)
    {
        if (original is null) yield break;

        var started = false;
        foreach (var child in original.ChildElements)
        {
            if (child is BookmarkStart or BookmarkEnd)
            {
                if (started != leading) yield return child;
                continue;
            }

            if (child is not ParagraphProperties) started = true;
        }
    }

    private IEnumerable<OpenXmlElement> WriteInline(
        Node node,
        List<DocumentFormat.OpenXml.Wordprocessing.Drawing>? originalImages = null)
    {
        switch (node.Type)
        {
            case "text":
                foreach (var element in WriteTextRun(node)) yield return element;
                break;

            case "hardBreak":
                yield return new Run(new Break());
                break;

            case "pageBreak":
                yield return new Run(new Break { Type = BreakValues.Page });
                break;

            // The file's id and name, or those the editor gave the new bookmark.
            case "bookmarkStart":
                yield return new BookmarkStart
                {
                    Id = Attr.String(node, "bid") ?? "0",
                    Name = Attr.String(node, "name") ?? string.Empty,
                };
                break;

            case "bookmarkEnd":
                yield return new BookmarkEnd { Id = Attr.String(node, "bid") ?? "0" };
                break;

            // Reply ends go along: in Word the whole thread embraces the range.
            case "commentStart":
                _comments ??= new CommentAnchors(_part);
                foreach (var id in _comments.Thread(Attr.String(node, "cid") ?? "0"))
                {
                    _comments.Start(id);
                    yield return new CommentRangeStart { Id = id };
                }

                break;

            case "commentEnd":
                _comments ??= new CommentAnchors(_part);
                foreach (var id in _comments.Thread(Attr.String(node, "cid") ?? "0"))
                {
                    if (_comments.Ranged(id)) yield return new CommentRangeEnd { Id = id };
                    yield return _comments.Reference(id);
                }

                break;

            case "field":
                foreach (var element in WriteField(node)) yield return element;
                break;

            // The body goes to the notes part (NotesWriter); the run with the id stays here.
            case "noteRef":
                yield return WriteNoteReference(node);
                break;

            // The OMML as it came; a new or edited one derives it from the MathML
            // (OmmlMath.ToOmml).
            case "math":
                if (MathOf(node) is { } math) yield return math;
                else _inventory.NoteLoss("equação que não pôde ser gravada");
                break;

            case "image":
                if ((originalImages is null ? null : _images.Reuse(node, originalImages)) is { } kept)
                {
                    yield return kept;
                }
                else if (_images.Write(node) is { } image)
                {
                    yield return image;
                }

                break;

            default:
                _inventory.NoteLoss($"conteúdo do tipo \"{node.Type}\"");
                break;
        }
    }

    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>
    /// MathML keeps neither color nor font of its tokens: inner formatting is lost, and that is
    /// said once.
    /// </summary>
    private void NoteEditedMathFormatting(Node paragraph, OpenXmlElement? original)
    {
        if (original is null || paragraph.Content is not { } content) return;
        var maths = content.Where(child => child.Type == "math").ToList();
        if (!maths.Any(math => Attr.String(math, "omml") is not { Length: > 0 })) return;

        var kept = maths.Select(math => Attr.String(math, "omml")).OfType<string>().ToHashSet(StringComparer.Ordinal);
        var equations = original.Descendants()
            .Where(element => element is DocumentFormat.OpenXml.Math.Paragraph ||
                              (element is DocumentFormat.OpenXml.Math.OfficeMath &&
                               element.Ancestors<DocumentFormat.OpenXml.Math.Paragraph>().FirstOrDefault() is null));
        if (equations.Any(equation => FormattedInside(equation) && !kept.Contains(equation.OuterXml)))
        {
            _inventory.NoteLoss("formatação dentro da equação editada");
        }
    }

    private static bool FormattedInside(OpenXmlElement equation) =>
        equation.Descendants<RunProperties>().Any(properties => properties.ChildElements.Any(child => child switch
        {
            RunFonts fonts => (fonts.Ascii?.Value ?? MathFontName) != MathFontName || (fonts.HighAnsi?.Value ?? MathFontName) != MathFontName,
            Italic or ItalicComplexScript or Languages or NoProof => false,
            _ => true,
        }));

    private const string MathFontName = "Cambria Math";

    /// <summary>Or null.</summary>
    internal static OpenXmlElement? MathOf(Node node)
    {
        var omml = Attr.String(node, "omml") is { Length: > 0 } kept
            ? kept
            : Attr.String(node, "mathml") is { Length: > 0 } mathMl
                ? OmmlMath.ToOmml(mathMl, Attr.Bool(node, "display"), Attr.String(node, "jc"))
                : null;
        if (omml is null) return null;
        try
        {
            var root = System.Xml.Linq.XElement.Parse(omml);
            if (root.Name.Namespace != OmmlMath.M) return null;
            OpenXmlElement? math = root.Name.LocalName switch
            {
                "oMathPara" => new DocumentFormat.OpenXml.Math.Paragraph(omml),
                "oMath" => new DocumentFormat.OpenXml.Math.OfficeMath(omml),
                _ => null,
            };

            // The root's `w:` declaration is enough: the equation goes out written as it came in.
            foreach (var element in math?.Descendants() ?? [])
            {
                foreach (var (prefix, uri) in element.NamespaceDeclarations.ToList())
                {
                    if ((prefix, uri) is ("w", WordprocessingNamespace) or ("m", MathNamespace))
                    {
                        element.RemoveNamespaceDeclaration(prefix);
                    }
                }
            }

            return math;
        }
        catch (System.Xml.XmlException)
        {
            return null;
        }
    }

    /// <summary>
    /// The original run's <c>w:rPr</c>, or Word's for a new note; its own mark comes right after.
    /// </summary>
    private Run WriteNoteReference(Node node)
    {
        var endnote = Attr.String(node, "kind") == NotesWriter.Endnote;
        var id = long.TryParse(Attr.String(node, "nid"), System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var number) ? number : 0L;
        _noteRuns ??= NotesWriter.ReferenceRunsOf(_part);

        var run = new Run();
        var address = NotesWriter.Address(endnote, id.ToString(System.Globalization.CultureInfo.InvariantCulture));
        run.RunProperties = _noteRuns.TryGetValue(address, out var original) && original.RunProperties is { } kept
            ? (RunProperties)kept.CloneNode(true)
            : NotesWriter.ReferenceProperties(_part, endnote);

        FootnoteEndnoteReferenceType reference = endnote ? new EndnoteReference() : new FootnoteReference();
        reference.Id = id;
        var mark = Attr.String(node, "mark");
        if (mark is { Length: > 0 }) reference.CustomMarkFollows = true;
        run.AppendChild(reference);
        if (mark is { Length: > 0 })
        {
            foreach (var piece in XmlText.Of(mark)) run.AppendChild(piece);
        }

        return run;
    }

    private (RunProperties Properties, string? Hyperlink) FormatOf(Node node)
    {
        var properties = new RunProperties();
        string? hyperlink = null;

        foreach (var mark in node.Marks ?? [])
        {
            switch (mark.Type)
            {
                // `off` turns off what the paragraph style turns on; see RunReader.Off.
                case "bold": properties.Bold = IsOff(mark) ? new Bold { Val = false } : new Bold(); break;
                case "italic": properties.Italic = IsOff(mark) ? new Italic { Val = false } : new Italic(); break;
                case "strike": properties.Strike = IsOff(mark) ? new Strike { Val = false } : new Strike(); break;
                case "caps": properties.Caps = new Caps(); break;
                case "smallCaps": properties.SmallCaps = new SmallCaps(); break;
                case "underline":
                    properties.Underline = new Underline
                    {
                        Val = IsOff(mark) ? UnderlineValues.None : UnderlineValues.Single,
                    };
                    break;

                // A single property, `w:vertAlign`, with mutually exclusive values, as in the
                // editor.
                case "superscript":
                    properties.VerticalTextAlignment = new VerticalTextAlignment
                    {
                        Val = VerticalPositionValues.Superscript,
                    };
                    break;

                case "subscript":
                    properties.VerticalTextAlignment = new VerticalTextAlignment
                    {
                        Val = VerticalPositionValues.Subscript,
                    };
                    break;

                case "highlight":
                    if (Attr.MarkString(mark, "color") is { } fill) ApplyHighlight(properties, fill);
                    break;

                case "link":
                    hyperlink = Attr.MarkString(mark, "href");
                    break;

                case "charStyle":
                    if (Attr.MarkString(mark, "styleId") is { Length: > 0 } characterStyle)
                    {
                        properties.RunStyle = new RunStyle { Val = characterStyle };
                    }

                    break;

                case "textStyle":
                    ApplyTextStyle(properties, mark);
                    break;

                // It wraps the run; see WriteParagraph.
                case Revisions.Insertion:
                case Revisions.Deletion:
                    break;

                default:
                    _inventory.NoteLoss($"formatação \"{mark.Type}\"");
                    break;
            }
        }

        DropWhatRepeatsTheStyle(properties, _paragraphRun);
        return (properties, hyperlink);
    }

    private IEnumerable<OpenXmlElement> WriteTextRun(Node node)
    {
        var run = new Run();
        var (properties, hyperlink) = FormatOf(node);
        if (properties.HasChildren) run.RunProperties = properties;

        // A tab is `w:tab`, a break `w:br`, and control characters do not exist in XML 1.0
        // (XmlText).
        var pieces = XmlText.Of(node.Text).ToList();
        if (pieces.Count == 0)
        {
            pieces.Add(new Text(string.Empty) { Space = SpaceProcessingModeValues.Preserve });
        }

        foreach (var piece in pieces) run.AppendChild(piece);

        return Linked([run], hyperlink);
    }

    private IEnumerable<OpenXmlElement> Linked(List<Run> runs, string? hyperlink)
    {
        if (hyperlink is null) return runs;

        // An internal link: `w:anchor`, without a relationship; `w:history`, as Word writes it.
        if (hyperlink.StartsWith('#'))
        {
            return [new Hyperlink(runs) { Anchor = hyperlink[1..], History = true }];
        }

        Uri target;
        try
        {
            target = new Uri(hyperlink, UriKind.Absolute);
        }
        catch (UriFormatException)
        {
            _inventory.NoteLoss("endereço de link inválido");
            return runs;
        }

        var relationship = _owner.AddHyperlinkRelationship(target, true);
        return [new Hyperlink(runs) { Id = relationship.Id }];
    }

    /// <summary>Each piece in its own run, with the node's formatting, as Word writes it.</summary>
    private IEnumerable<OpenXmlElement> WriteField(Node node)
    {
        var (properties, hyperlink) = FormatOf(node);

        Run Piece(params OpenXmlElement[] content)
        {
            var run = new Run();
            if (properties.HasChildren) run.RunProperties = (RunProperties)properties.CloneNode(true);
            foreach (var element in content) run.AppendChild(element);
            return run;
        }

        var runs = new List<Run>
        {
            Piece(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            Piece(new FieldCode(Attr.String(node, "instr") ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }),
            Piece(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        };

        var result = XmlText.Of(Attr.String(node, "result")).ToArray();
        if (result.Length > 0) runs.Add(Piece(result));
        runs.Add(Piece(new FieldChar { FieldCharType = FieldCharValues.End }));

        return Linked(runs, hyperlink);
    }

    /// <summary>
    /// The paragraphs with the <c>TOC</c> field around them, in the content control if there was
    /// one. The title goes back with the original <c>w:pPr</c>; the entries do not, because "Update
    /// table" rebuilds them. An entry without a tab stop gets Word's, right-aligned with dot
    /// leaders.
    /// </summary>
    private IEnumerable<OpenXmlElement> WriteTableOfContents(Node node, OpenXmlElement? original)
    {
        var children = node.Content ?? [];
        if (children.Count == 0) children = [Node.Of("paragraph")];

        var head = Math.Clamp(Attr.Int(node, "head") ?? 0, 0, children.Count - 1);
        var originals = original switch
        {
            SdtBlock block => block.SdtContentBlock?.Elements<Paragraph>().ToList() ?? [],
            Paragraph first => [first],
            _ => [],
        };

        var paragraphs = new List<Paragraph>();
        for (var index = 0; index < children.Count; index++)
        {
            var source = index < head ? originals.ElementAtOrDefault(index) : null;
            foreach (var paragraph in Write(children[index], null, source).OfType<Paragraph>())
            {
                if (index >= head) WithPageTab(paragraph);
                paragraphs.Add(paragraph);
            }
        }

        var opening = paragraphs[Math.Min(head, paragraphs.Count - 1)];
        var at = opening.ParagraphProperties is null ? 0 : 1;
        opening.InsertAt(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), at);
        opening.InsertAt(
            new Run(new FieldCode(Attr.String(node, "instr") ?? " TOC \\o \"1-3\" \\h \\z \\u ")
            {
                Space = SpaceProcessingModeValues.Preserve,
            }),
            at + 1);
        opening.InsertAt(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), at + 2);
        paragraphs[^1].AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        if (!Attr.Bool(node, "sdt"))
        {
            foreach (var paragraph in paragraphs) yield return paragraph;
            yield break;
        }

        // The original's properties; a new table of contents gets Word's.
        var sdt = new SdtBlock();
        if (original is SdtBlock previous && previous.SdtProperties is { } kept)
        {
            sdt.AppendChild(kept.CloneNode(true));
            if (previous.SdtEndCharProperties is { } end) sdt.AppendChild(end.CloneNode(true));
        }
        else
        {
            sdt.AppendChild(new SdtProperties(
                new SdtContentDocPartObject(
                    new DocPartGallery { Val = "Table of Contents" },
                    new DocPartUnique())));
        }

        sdt.AppendChild(new SdtContentBlock(paragraphs));
        yield return sdt;
    }

    private void WithPageTab(Paragraph paragraph)
    {
        if (!paragraph.Descendants<FieldCode>().Any(code => code.Text.Contains("PAGEREF", StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
        if (properties.Tabs is not null) return;

        var tabs = new Tabs(new TabStop
        {
            Val = TabStopValues.Right,
            Leader = TabStopLeaderCharValues.Dot,
            Position = _usableTwips,
        });

        // In schema order: after style, numbering and borders, before spacing.
        var before = properties.ChildElements.LastOrDefault(child =>
            child is ParagraphStyleId or KeepNext or KeepLines or PageBreakBefore or FrameProperties
                or WidowControl or NumberingProperties or SuppressLineNumbers or ParagraphBorders or Shading);
        if (before is null) properties.PrependChild(tabs);
        else properties.InsertAfter(tabs, before);
    }

    /// <summary>
    /// Marks arrive flattened, with the style inside: written back, they would detach the paragraph
    /// from its style. Only what matches goes; an absent mark does not become "off".
    /// </summary>
    private static void DropWhatRepeatsTheStyle(RunProperties properties, RunProperties style)
    {
        // On or off, what matches goes.
        if (properties.Bold is not null && RunReader.IsOn(properties.Bold) == RunReader.IsOn(style.Bold))
        {
            properties.Bold = null;
        }

        if (properties.Italic is not null && RunReader.IsOn(properties.Italic) == RunReader.IsOn(style.Italic))
        {
            properties.Italic = null;
        }

        if (properties.Strike is not null && RunReader.IsOn(properties.Strike) == RunReader.IsOn(style.Strike))
        {
            properties.Strike = null;
        }

        if (properties.Caps is not null && RunReader.IsOn(style.Caps)) properties.Caps = null;
        if (properties.SmallCaps is not null && RunReader.IsOn(style.SmallCaps)) properties.SmallCaps = null;

        if (properties.Underline?.Val is { } underline &&
            (underline.Value != UnderlineValues.None) ==
            (style.Underline?.Val is { } line && line.Value != UnderlineValues.None))
        {
            properties.Underline = null;
        }

        if (Same(properties.RunFonts?.Ascii?.Value, style.RunFonts?.Ascii?.Value)) properties.RunFonts = null;
        if (Same(properties.FontSize?.Val?.Value, style.FontSize?.Val?.Value)) properties.FontSize = null;
        if (Same(properties.Color?.Val?.Value, style.Color?.Val?.Value)) properties.Color = null;
        if (Same(properties.VerticalTextAlignment?.Val?.InnerText, style.VerticalTextAlignment?.Val?.InnerText))
        {
            properties.VerticalTextAlignment = null;
        }
    }

    private static bool IsOff(Mark mark) =>
        mark.Attrs?.TryGetValue("off", out var off) == true &&
        off?.GetValueKind() == System.Text.Json.JsonValueKind.True;

    private static bool Same(string? a, string? b) =>
        a is not null && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    private void ApplyTextStyle(RunProperties properties, Mark mark)
    {
        if (Attr.MarkString(mark, "color") is { } color)
        {
            if (ColorValue.Hex(color) is { } hex) properties.Color = new Color { Val = hex };
            else _inventory.NoteLoss($"cor de texto \"{color}\"");
        }

        // The editor's `backgroundColor` is the same `w:shd` as the file's highlight.
        if (Attr.MarkString(mark, "backgroundColor") is { } background)
        {
            ApplyHighlight(properties, background);
        }

        if (Attr.MarkString(mark, "fontFamily") is { } font &&
            ParagraphFormat.FirstFont(font) is { } first)
        {
            properties.RunFonts = new RunFonts { Ascii = first, HighAnsi = first };
        }

        if (Attr.MarkString(mark, "fontSize") is { } size)
        {
            // `w:sz` is in half-points, and the measure may arrive in pixels, as CSS writes it.
            if (Attr.Points(size) is { } points && points > 0)
            {
                var halfPoints = (int)Math.Round(points * Unit.HalfPointsPerPoint);
                properties.FontSize = new FontSize
                {
                    Val = halfPoints.ToString(CultureInfo.InvariantCulture),
                };
            }
            else
            {
                _inventory.NoteLoss($"tamanho de fonte \"{size}\"");
            }
        }

        // Line spacing belongs to the paragraph (ParagraphFormat): in a run it has nowhere to go.
        if (Attr.MarkString(mark, "lineHeight") is { } lineHeight)
        {
            _inventory.NoteLoss($"entrelinha de um trecho de texto (\"{lineHeight}\")");
        }
    }

    /// <summary>
    /// Always six-digit hex: <c>rgb(...)</c> would make Word declare the document damaged, and a
    /// color name would come out black.
    /// </summary>
    private void ApplyHighlight(RunProperties properties, string color)
    {
        if (ColorValue.Hex(color) is not { } hex)
        {
            _inventory.NoteLoss($"cor de fundo de texto \"{color}\"");
            return;
        }

        properties.Shading = new Shading
        {
            Val = ShadingPatternValues.Clear,
            Color = "auto",
            Fill = hex,
        };
    }

    /// <summary>
    /// Each thread's replies, comments with a range (inventing a <c>w:commentRangeEnd</c> on a
    /// point comment would be wrong) and each reference's run.
    /// </summary>
    private sealed class CommentAnchors
    {
        private readonly Dictionary<string, List<string>> _replies;
        private readonly HashSet<string> _ranged;
        private readonly Dictionary<string, Run> _references = new(StringComparer.Ordinal);

        /// <summary>
        /// A comment created in the editor is not in the original: its start says it has a range.
        /// </summary>
        private readonly HashSet<string> _started = new(StringComparer.Ordinal);

        public CommentAnchors(MainDocumentPart part)
        {
            _replies = CommentsReader.RepliesOf(part);
            // Note paragraphs carry anchors too.
            var roots = CommentsWriter.AnchorRoots(part);
            _ranged = roots.SelectMany(root => root.Descendants<CommentRangeStart>())
                .Select(start => start.Id?.Value).OfType<string>().ToHashSet(StringComparer.Ordinal);
            foreach (var run in roots.SelectMany(root => root.Descendants<Run>()))
            {
                if (BodyReader.ReferenceOnly(run) is { } id) _references.TryAdd(id, run);
            }
        }

        public IEnumerable<string> Thread(string id) => [id, .. _replies.GetValueOrDefault(id) ?? []];

        public void Start(string id) => _started.Add(id);

        public bool Ranged(string id) => _ranged.Contains(id) || _started.Contains(id);

        public Run Reference(string id) =>
            _references.TryGetValue(id, out var original)
                ? (Run)original.CloneNode(true)
                : new Run(
                    new RunProperties(new RunStyle { Val = "CommentReference" }),
                    new CommentReference { Id = id });
    }
}
