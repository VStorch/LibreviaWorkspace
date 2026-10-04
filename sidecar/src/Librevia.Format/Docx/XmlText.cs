using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O texto de um nó → os filhos de um <c>w:r</c>. O XML 1.0 não representa a
/// maioria dos caracteres de controle, nem escapados: a tabulação vira <c>w:tab</c>,
/// <c>\n</c> e o <c>\u000B</c> da quebra manual do Word viram <c>w:br</c>, e o resto sai.
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

            // Sem `preserve` o Word engole os espaços das pontas.
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
                    // `\r\n` é uma quebra só: a do `\n` que vem a seguir.
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
    /// Par substituto só vale inteiro: a metade de um emoji cortado faria um arquivo
    /// que não se reabre.
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
