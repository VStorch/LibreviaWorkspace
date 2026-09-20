using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O texto de um nó do editor → os filhos de um `w:r`.
/// </summary>
/// <remarks>
/// Não é frescura de formato: o XML 1.0 **não tem** como representar a maioria
/// dos caracteres de controle, nem escapados. Um `\u0001` colado de um terminal
/// — ou um `\u000B` que qualquer editor de texto do Windows produz — derrubava a
/// gravação inteira no serializador, com uma mensagem sobre XML inválido que não
/// dizia nada a quem só queria salvar o documento. Perder o arquivo por causa de
/// um caractere invisível é o pior negócio possível.
///
/// Os que têm significado voltam a tê-lo: a tabulação é `w:tab`, e a quebra de
/// linha — tanto `\n` como o `\u000B` que o Word usa para "quebra de linha
/// manual" — é `w:br`. O resto é descartado, porque não existe nada a escrever
/// no lugar e o usuário não vê diferença.
/// </remarks>
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

            // `xml:space="preserve"` senão o Word engole espaço no começo e no
            // fim, e frases coladas aparecem sem separação.
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
    /// O caractere pode existir num documento XML 1.0?
    /// </summary>
    /// <remarks>
    /// Um par substituto é válido junto e inválido separado: o emoji passa, a
    /// metade de emoji que sobrou de um corte de string, não. Escrevê-la geraria
    /// um arquivo que nem nós conseguiríamos reabrir.
    /// </remarks>
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
