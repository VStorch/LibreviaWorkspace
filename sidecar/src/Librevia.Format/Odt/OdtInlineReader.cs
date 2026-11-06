using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The content of a <c>text:p</c> or <c>text:h</c>: text with marks, line breaks and images. ODF
/// collapses white space as HTML does; <c>text:s</c>, <c>text:tab</c> and
/// <c>text:line-break</c> are the real ones.
/// </summary>
internal sealed class OdtInlineReader(OdtReadContext context)
{
    /// <summary>Against <c>text:c="2000000000"</c>.</summary>
    internal const int MaxSpaces = 1000;

    private readonly OdtFrameReader _frames = new(context);
    private OdtBodyReader? _noteBodies;
    private List<Node> _nodes = [];
    private CharacterFormat _inherited = CharacterFormat.Empty;
    private bool _afterSpace;
    private string? _lastMarks;

    private readonly record struct Run(CharacterFormat Format, string? Href, string? CharStyle);

    /// <param name="runBase">What every run starts from: the paragraph's effective formatting.</param>
    /// <param name="inherited">What the block already draws: marks only say what differs.</param>
    public List<Node> Read(XElement paragraph, CharacterFormat runBase, CharacterFormat inherited)
    {
        _nodes = [];
        _inherited = inherited;
        _afterSpace = true;
        _lastMarks = null;
        ReadChildren(paragraph, new Run(runBase, null, null));
        return _nodes;
    }

    private void ReadChildren(XElement container, Run run)
    {
        foreach (var child in container.Nodes())
        {
            if (child is XText text) AppendCollapsing(text.Value, run);
            else if (child is XElement element) ReadElement(element, run);
        }
    }

    private void ReadElement(XElement element, Run run)
    {
        if (element.Name.Namespace == Ns.Text && ReadText(element, run)) return;
        if (element.Name.Namespace == Ns.Draw)
        {
            ReadDrawing(element, run);
            return;
        }

        if (element.Name.Namespace == Ns.Office && ReadComment(element)) return;
        if (element.Name == Ns.Text + "note" && !context.InNote)
        {
            ReadNote(element);
            return;
        }

        if (OdtNotices.DroppedInline(element) is { } key)
        {
            context.NoteLoss(key);
            return;
        }

        // Unknown: named in the inventory, and its text stays.
        OdtNotices.Element(context.Inventory, element);
        ReadChildren(element, run);
    }

    /// <summary>The anchors are zero-width: the white space around them collapses as if they were not there.</summary>
    private bool ReadComment(XElement element)
    {
        if (element.Name.LocalName == "annotation-end")
        {
            if (context.Comments.End(element) is { } end) Add(end);
            return true;
        }

        if (element.Name.LocalName != "annotation") return false;
        var anchor = context.Comments.Annotation(element, out var rich);
        if (rich) context.NoteLoss(Inventory.Comments);
        if (anchor is not null) Add(anchor);
        return true;
    }

    /// <summary>The body flattened, like a list or a cell, as the DOCX reader gives it.</summary>
    private void ReadNote(XElement note)
    {
        var endnote = note.Attr(Ns.Text + "note-class") == OdtNoteCounter.Endnote;
        var node = Node.Of("noteRef")
            .With("kind", endnote ? OdtNoteCounter.Endnote : OdtNoteCounter.Footnote)
            .With("nid", note.Attr(Ns.Text + "id"));
        if (note.Element(Ns.Text + "note-citation")?.Attr(Ns.Text + "label") is { Length: > 0 } mark) node.With("mark", mark);

        _noteBodies ??= new OdtBodyReader(context);
        context.InNote = true;
        var body = note.Element(Ns.Text + "note-body")?.Elements().SelectMany(child => _noteBodies.ReadBlock(child, flat: true)).ToList() ?? [];
        context.InNote = false;
        node.Content = body.Count == 0 ? [OdtParagraphReader.Empty()] : body;
        Add(node);
        _afterSpace = false;
    }

    /// <summary><c>false</c> for what is not a text construct this reader handles.</summary>
    private bool ReadText(XElement element, Run run)
    {
        switch (element.Name.LocalName)
        {
            case "s":
                var count = Math.Clamp(OdtLength.Integer(element.Attr(Ns.Text + "c")) ?? 1, 1, MaxSpaces);
                AppendRaw(new string(' ', count), run);
                _afterSpace = true;
                return true;
            case "tab":
                AppendRaw("\t", run);
                _afterSpace = true;
                return true;
            case "line-break":
                Add(Node.Of("hardBreak"));
                _afterSpace = true;
                return true;
            case "soft-page-break":
                return true;
            case "span":
                ReadChildren(element, Span(run, element.Attr(Ns.Text + "style-name")));
                return true;
            case "a":
                ReadChildren(element, run with { Href = LinkOf(element.Attr(Ns.XLink + "href")) ?? run.Href });
                return true;
            case "meta":
                ReadChildren(element, run);
                return true;
            case "ruby":
                OdtNotices.Element(context.Inventory, element);
                if (element.Element(Ns.Text + "ruby-base") is { } rubyBase) ReadChildren(rubyBase, run);
                return true;
            default:
                return ReadField(element, run);
        }
    }

    /// <summary>
    /// One with a Word equivalent is a <c>field</c> node; any other leaves its text, and costs it
    /// only if the block is rewritten.
    /// </summary>
    private bool ReadField(XElement element, Run run)
    {
        if (!OdtNotices.IsTextField(element)) return false;
        if (OdtFields.InstructionOf(element) is not { } instruction)
        {
            context.NoteLoss(Inventory.PlainFields);
            ReadChildren(element, run);
            return true;
        }

        var field = Node.Of("field").With("instr", instruction).With("result", ResultOf(element));
        field.Marks = OdtMarks.Of(run.Format, _inherited, run.Href, run.CharStyle);
        Add(field);
        _afterSpace = false;
        return true;
    }

    /// <summary>The text as shown: <c>text:s</c> and <c>text:tab</c> count, collapsed white space does not.</summary>
    private static string ResultOf(XElement element)
    {
        var text = new StringBuilder();
        foreach (var node in element.DescendantNodes())
        {
            if (node is XText piece) text.Append(piece.Value);
            else if (node is XElement { Name.LocalName: "tab" }) text.Append('\t');
            else if (node is XElement { Name.LocalName: "s" } space)
            {
                text.Append(' ', Math.Clamp(OdtLength.Integer(space.Attr(Ns.Text + "c")) ?? 1, 1, MaxSpaces));
            }
        }

        return text.ToString();
    }

    private void ReadDrawing(XElement element, Run run)
    {
        switch (element.Name.LocalName)
        {
            case "frame":
                if (_frames.Read(element) is not { } framed) return;
                if (framed.Type == "math") framed.Marks = OdtMarks.Of(run.Format, _inherited, run.Href, run.CharStyle);
                else if (run.Href is { } href) framed.Marks = [Mark.Of("link", "href", href)];
                Add(framed);
                _afterSpace = false;
                return;
            case "a":
                ReadChildren(element, run with { Href = LinkOf(element.Attr(Ns.XLink + "href")) ?? run.Href });
                return;
            default:
                context.NoteLoss(OdtNotices.ShapeKey(element));
                return;
        }
    }

    /// <summary>
    /// The address goes as the file has it, as the DOCX reader does: the editor and main apply the
    /// scheme rule. An internal one (<c>#name</c>) would point to a bookmark the model does not keep.
    /// </summary>
    private static string? LinkOf(string? href) =>
        string.IsNullOrWhiteSpace(href) || href.StartsWith('#') ? null : href.Trim();

    /// <summary>The span's character style (or the named one under its automatic style) over the run.</summary>
    private Run Span(Run run, string? styleName)
    {
        var chain = context.Styles.Chain(OdtStyles.TextFamily, styleName);
        if (chain.Count == 0) return run;
        var named = context.Styles.NamedOf(OdtStyles.TextFamily, styleName)?.Id;
        return new Run(context.Styles.Character(chain, run.Format), run.Href, named ?? run.CharStyle);
    }

    private void AppendCollapsing(string value, Run run)
    {
        var text = new StringBuilder(value.Length);
        foreach (var letter in value)
        {
            if (letter is ' ' or '\t' or '\r' or '\n')
            {
                if (_afterSpace) continue;
                text.Append(' ');
                _afterSpace = true;
            }
            else
            {
                text.Append(letter);
                _afterSpace = false;
            }
        }

        AppendRaw(text.ToString(), run);
    }

    /// <summary>Neighbouring text with the same marks is one node, as ProseMirror keeps it.</summary>
    private void AppendRaw(string text, Run run)
    {
        if (text.Length == 0) return;

        var marks = OdtMarks.Of(run.Format, _inherited, run.Href, run.CharStyle);
        var key = marks is null ? string.Empty : JsonSerializer.Serialize(marks, DocxJson.Options);
        if (_lastMarks == key && _nodes.Count > 0 && _nodes[^1].Type == "text")
        {
            _nodes[^1].Text += text;
            return;
        }

        _nodes.Add(new Node { Type = "text", Text = text, Marks = marks });
        _lastMarks = key;
    }

    private void Add(Node node)
    {
        _nodes.Add(node);
        _lastMarks = null;
    }
}
