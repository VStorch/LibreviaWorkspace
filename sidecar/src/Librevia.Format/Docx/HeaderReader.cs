using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Cabeçalho e rodapé → texto para exibir.
/// </summary>
/// <remarks>
/// Extração **de mão única**, e é isso que a torna barata. O que vai para o
/// arquivo é a parte OOXML original, copiada intacta (decisão do Vinícius em
/// 2026-08-14 — ver docs/01-corpus-docx.md, Descoberta 4). Daqui sai só o que a
/// tela e o PDF precisam mostrar, então um erro aqui é cosmético.
///
/// O modelo guarda cabeçalho como uma linha de texto com `{n}` e `{total}`.
/// Logo, tudo que o cabeçalho tiver de imagem e posicionamento não aparece —
/// invisibilidade, não perda.
/// </remarks>
public static class HeaderReader
{
    public static string Read(SectionProperties section, MainDocumentPart part, Inventory inventory) =>
        ReadReferenced(
            section.Elements<HeaderReference>().Select(reference => reference.Id?.Value),
            part,
            inventory,
            "cabeçalho");

    public static string ReadFooter(SectionProperties section, MainDocumentPart part, Inventory inventory) =>
        ReadReferenced(
            section.Elements<FooterReference>().Select(reference => reference.Id?.Value),
            part,
            inventory,
            "rodapé");

    private static string ReadReferenced(
        IEnumerable<string?> relationshipIds,
        MainDocumentPart part,
        Inventory inventory,
        string what)
    {
        foreach (var id in relationshipIds)
        {
            if (string.IsNullOrEmpty(id)) continue;

            var root = part.GetPartById(id) switch
            {
                HeaderPart header => (OpenXmlPartRootElement?)header.Header,
                FooterPart footer => footer.Footer,
                _ => null,
            };

            if (root is null) continue;

            var text = Flatten(root, inventory);
            // A primeira página costuma ter cabeçalho vazio; o que interessa é
            // o primeiro que tenha conteúdo.
            if (text.Length > 0) return text;
        }

        return string.Empty;
    }

    private static string Flatten(OpenXmlPartRootElement root, Inventory inventory)
    {
        var builder = new StringBuilder();
        Walk(root, builder, inventory, new FieldState());
        return Collapse(builder.ToString());
    }

    /// <summary>Onde estamos dentro de um campo `PAGE`.</summary>
    /// <remarks>
    /// Um campo do OOXML tem três marcos: `begin`, `separate` e `end`. Entre
    /// `separate` e `end` está o **último valor calculado**, que o Word deixou
    /// em cache. Se copiarmos esse texto junto com o nosso `{n}`, o cabeçalho
    /// vira "{n}5" — o marcador e o número da página em que o arquivo foi salvo.
    /// </remarks>
    private sealed class FieldState
    {
        public int Depth;
        public bool InCachedResult;
    }

    private static void Walk(OpenXmlElement parent, StringBuilder builder, Inventory inventory, FieldState field)
    {
        foreach (var element in parent.ChildElements)
        {
            switch (element)
            {
                // Um desenho moderno vem embrulhado junto com um fallback VML
                // do mesmo conteúdo. Percorrer os dois duplicaria o cabeçalho
                // inteiro; ficamos com a primeira alternativa.
                case AlternateContent alternate:
                {
                    var branch = (OpenXmlElement?)alternate.GetFirstChild<AlternateContentChoice>()
                                 ?? alternate.GetFirstChild<AlternateContentFallback>();
                    if (branch is not null) Walk(branch, builder, inventory, field);
                    break;
                }

                case FieldChar marker:
                    switch (marker.FieldCharType?.Value)
                    {
                        case { } type when type == FieldCharValues.Begin:
                            field.Depth++;
                            break;
                        case { } type when type == FieldCharValues.Separate:
                            field.InCachedResult = true;
                            break;
                        case { } type when type == FieldCharValues.End:
                            field.Depth = Math.Max(0, field.Depth - 1);
                            field.InCachedResult = false;
                            break;
                    }

                    break;

                case FieldCode code:
                    if (code.Text.Contains("NUMPAGES", StringComparison.Ordinal)) builder.Append("{total}");
                    else if (code.Text.Contains("PAGE", StringComparison.Ordinal)) builder.Append("{n}");
                    break;

                case Text text:
                    if (!field.InCachedResult) builder.Append(text.Text);
                    break;

                case TabChar:
                    builder.Append(' ');
                    break;

                case DocumentFormat.OpenXml.Wordprocessing.Drawing:
                case Picture:
                    inventory.NoteInvisible("imagens no cabeçalho (o logotipo continua no arquivo)");
                    Walk(element, builder, inventory, field);
                    break;

                default:
                    Walk(element, builder, inventory, field);
                    break;
            }
        }
    }

    /// <summary>
    /// O texto sai picado em runs. Colapsar espaço junta as partes; se o
    /// resultado for a mesma frase duas vezes — cabeçalho par e ímpar com o
    /// mesmo conteúdo — fica só uma.
    /// </summary>
    private static string Collapse(string raw)
    {
        var parts = raw.Split((char[])[' ', '\t', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
        var joined = string.Join(' ', parts).Trim();

        if (joined.Length == 0) return string.Empty;

        if (joined.Length % 2 == 1)
        {
            var half = joined.Length / 2;
            if (joined[half] == ' ' &&
                string.Equals(joined[..half], joined[(half + 1)..], StringComparison.Ordinal))
            {
                return joined[..half];
            }
        }

        return joined;
    }
}
