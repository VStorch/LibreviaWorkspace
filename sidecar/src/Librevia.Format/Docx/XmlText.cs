using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A node's text → the children of a <c>w:r</c>. XML 1.0 cannot represent most control characters,
/// not even escaped: a tab becomes <c>w:tab</c>, <c>\n</c> and Word's manual line break
/// <c>\u000B</c> become <c>w:br</c>, and the rest goes.
/// </summary>
internal static class XmlText
{
    public static IEnumerable<OpenXmlElement> Of(string? text)
    {
        var elements = new List<OpenXmlElement>();
        if (string.IsNullOrEmpty(text)) return elements;

        var pending = new StringBuilder();

        void Flush()
        {
            if (pending.Length == 0) return;

            // Without `preserve` Word swallows the spaces at the ends.
            elements.Add(new Text(pending.ToString()) { Space = SpaceProcessingModeValues.Preserve });
            pending.Clear();
        }

        for (var index = 0; index < text.Length; index++)
        {
            var letter = text[index];
            switch (letter)
            {
                case '\t':
                    Flush();
                    elements.Add(new TabChar());
                    break;

                case '\r':
                    // `\r\n` is a single break: the `\n` that follows provides it.
                    if (index + 1 < text.Length && text[index + 1] == '\n') break;
                    Flush();
                    elements.Add(new Break());
                    break;

                case '\n':
                case '\u000B':
                    Flush();
                    elements.Add(new Break());
                    break;

                default:
                    if (IsAllowed(text, index)) pending.Append(letter);
                    break;
            }
        }

        Flush();
        return elements;
    }

    /// <summary>
    /// A surrogate pair only counts whole: half of a cut emoji would make a file that does not
    /// reopen.
    /// </summary>
    private static bool IsAllowed(string text, int index)
    {
        var letter = text[index];

        if (letter < 0x20) return false;
        if (letter is '￾' or '￿') return false;

        if (char.IsHighSurrogate(letter))
        {
            return index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]);
        }

        if (char.IsLowSurrogate(letter))
        {
            return index > 0 && char.IsHighSurrogate(text[index - 1]);
        }

        return true;
    }
}
