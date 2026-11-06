using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// Where each element of a part sits in its text, down to <see cref="MaxDepth"/>: what lets saving
/// cut an untouched block out of <c>content.xml</c> byte for byte. The XML is already known to be
/// well formed (<see cref="XmlParts"/>), so a lexer that skips comments, processing instructions,
/// CDATA and quoted attributes is enough.
/// </summary>
internal sealed class XmlSpan
{
    public required int Start { get; init; }

    public int ContentStart { get; set; }

    /// <summary>At the end tag's <c>&lt;</c>; equal to <see cref="End"/> when self-closing.</summary>
    public int ContentEnd { get; set; }

    public int End { get; set; }
    public bool SelfClosing { get; set; }
    public List<XmlSpan> Children { get; } = [];
}

internal static class OdtXmlSpans
{
    /// <summary>Root, <c>office:body</c>, <c>office:text</c>, a block: deeper is never cut.</summary>
    public const int MaxDepth = 3;

    public static XmlSpan Root(string text)
    {
        var stack = new Stack<XmlSpan?>();
        XmlSpan? root = null;
        var index = 0;
        while ((index = text.IndexOf('<', index)) >= 0)
        {
            if (Skip(text, index) is { } skipped)
            {
                index = skipped;
                continue;
            }

            if (text[index + 1] == '/')
            {
                index = CloseTag(text, index, stack);
                continue;
            }

            index = OpenTag(text, index, stack, ref root);
        }

        return root ?? throw new OdtException(OdtException.Unreadable);
    }

    /// <summary>The element under <paramref name="span"/> that is <paramref name="element"/>.</summary>
    public static XmlSpan ChildOf(XmlSpan span, XElement parent, XElement element)
    {
        var index = parent.Elements().TakeWhile(child => child != element).Count();
        return index < span.Children.Count ? span.Children[index] : throw new OdtException(OdtException.Unreadable);
    }

    /// <summary>The spans of a path of elements from the root down.</summary>
    public static XmlSpan Find(XmlSpan root, XElement element)
    {
        var path = element.AncestorsAndSelf().Reverse().ToList();
        var span = root;
        for (var level = 1; level < path.Count; level++) span = ChildOf(span, path[level - 1], path[level]);
        return span;
    }

    private static int? Skip(string text, int index)
    {
        if (Starts(text, index, "<!--")) return End(text, index, "-->");
        if (Starts(text, index, "<![CDATA[")) return End(text, index, "]]>");
        if (Starts(text, index, "<?")) return End(text, index, "?>");
        return Starts(text, index, "<!") ? End(text, index, ">") : null;
    }

    private static bool Starts(string text, int index, string token) =>
        string.CompareOrdinal(text, index, token, 0, token.Length) == 0;

    private static int End(string text, int index, string token)
    {
        var end = text.IndexOf(token, index, StringComparison.Ordinal);
        return end < 0 ? throw new OdtException(OdtException.Unreadable) : end + token.Length;
    }

    private static int CloseTag(string text, int index, Stack<XmlSpan?> stack)
    {
        var after = End(text, index, ">");
        if (stack.TryPop(out var span) && span is not null)
        {
            span.ContentEnd = index;
            span.End = after;
        }

        return after;
    }

    private static int OpenTag(string text, int index, Stack<XmlSpan?> stack, ref XmlSpan? root)
    {
        var after = TagEnd(text, index);
        var selfClosing = text[after - 2] == '/';
        var depth = stack.Count;
        XmlSpan? span = null;
        if (depth <= MaxDepth && (depth == 0 || stack.Peek() is not null))
        {
            span = new XmlSpan { Start = index, ContentStart = after, SelfClosing = selfClosing };
            if (depth == 0) root ??= span;
            else stack.Peek()!.Children.Add(span);
        }

        if (selfClosing)
        {
            if (span is not null)
            {
                span.ContentEnd = after;
                span.End = after;
            }
        }
        else
        {
            stack.Push(span);
        }

        return after;
    }

    /// <summary>Past the <c>&gt;</c> that closes the tag, quotes respected.</summary>
    private static int TagEnd(string text, int index)
    {
        char? quote = null;
        for (var position = index + 1; position < text.Length; position++)
        {
            var letter = text[position];
            if (quote is not null)
            {
                if (letter == quote) quote = null;
            }
            else if (letter is '"' or '\'')
            {
                quote = letter;
            }
            else if (letter == '>')
            {
                return position + 1;
            }
        }

        throw new OdtException(OdtException.Unreadable);
    }
}
