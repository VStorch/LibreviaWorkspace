using System.IO.Compression;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

public sealed record SaveResult(
    [property: JsonPropertyName("inventory")] Inventory Inventory,
    [property: JsonPropertyName("preservedBlocks")] int PreservedBlocks,
    [property: JsonPropertyName("rewrittenBlocks")] int RewrittenBlocks);

/// <summary>
/// Gravação cirúrgica: reescreve só o que o usuário tocou.
/// </summary>
/// <remarks>
/// O desenho está em docs/02-docx-cirurgico.md. Em uma frase: a fidelidade não
/// vem de entender o OOXML, vem de **não mexer** no que não foi editado.
///
/// Só `word/document.xml` é reescrito. Estilos, numeração, cabeçalhos,
/// rodapés, mídia, comentários, notas, tema e configurações continuam
/// exatamente como estavam, porque ninguém os abre.
/// </remarks>
public static class DocxWriter
{
    public static (byte[] Bytes, SaveResult Result) Write(byte[] original, DocumentModelDto model)
    {
        var inventory = new Inventory();

        // Trabalha numa cópia gravável: `original` são os bytes que o processo
        // main guardou na abertura e não podem ser alterados.
        using var buffer = new MemoryStream();
        buffer.Write(original, 0, original.Length);
        buffer.Position = 0;

        using var document = OpenEditable(buffer);
        var part = document.MainDocumentPart
                   ?? throw new DocxException("O arquivo original não contém um documento do Word.");
        var body = part.Document?.Body
                   ?? throw new DocxException("O arquivo original está vazio ou danificado.");

        // Reindexa a partir dos mesmos bytes: os ids saem iguais aos da
        // abertura porque a numeração é posicional e determinística.
        var (_, blocks) = new BodyReader(part, new Inventory()).Read(body);
        var index = blocks.ToDictionary(block => block.Oid, StringComparer.Ordinal);

        var section = body.Elements<SectionProperties>().LastOrDefault();
        var replacement = BuildBody(model, part, index, inventory, out var preserved, out var rewritten);

        body.RemoveAllChildren();
        foreach (var element in replacement) body.AppendChild(element);

        // `w:sectPr` fecha o corpo e carrega a configuração de página.
        body.AppendChild(section is null ? new SectionProperties() : section);
        ApplyPageSetup(body.Elements<SectionProperties>().Last(), model.Page);

        part.Document!.Save();
        document.Dispose();

        return (RestoreUntouchedParts(original, buffer.ToArray()),
            new SaveResult(inventory, preserved, rewritten));
    }

    /// <summary>
    /// Partes que a gravação tem o direito de alterar. Todo o resto volta a ser
    /// exatamente o que era.
    /// </summary>
    private static readonly HashSet<string> Writable = new(StringComparer.Ordinal)
    {
        "word/document.xml",
        // Mudam quando o usuário insere imagem ou link num bloco editado.
        "word/_rels/document.xml.rels",
        "[Content_Types].xml",
    };

    /// <summary>
    /// Devolve às demais partes o conteúdo original, byte a byte.
    /// </summary>
    /// <remarks>
    /// Não é zelo excessivo: o SDK **reserializa toda parte cujo DOM tipado foi
    /// materializado**, mesmo sem alteração nenhuma. Basta o leitor de
    /// numeração tocar em `NumberingDefinitionsPart.Numbering` para
    /// `word/numbering.xml` sair diferente de um documento que ninguém editou —
    /// medido no corpus real.
    ///
    /// Reescrita sem intenção é a porta por onde a fidelidade escapa, e ela
    /// escapa em silêncio. Impor a invariante aqui é mais seguro do que confiar
    /// em ninguém nunca materializar um DOM sem querer.
    /// </remarks>
    private static byte[] RestoreUntouchedParts(byte[] original, byte[] produced)
    {
        using var originalArchive = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read);
        var pristine = originalArchive.Entries.ToDictionary(
            entry => entry.FullName,
            entry =>
            {
                using var stream = entry.Open();
                using var copy = new MemoryStream();
                stream.CopyTo(copy);
                return copy.ToArray();
            },
            StringComparer.Ordinal);

        using var source = new ZipArchive(new MemoryStream(produced), ZipArchiveMode.Read);
        using var result = new MemoryStream();

        using (var output = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in source.Entries)
            {
                var keepOriginal = !Writable.Contains(entry.FullName) &&
                                   pristine.ContainsKey(entry.FullName);

                using var target = output.CreateEntry(entry.FullName, CompressionLevel.Optimal).Open();

                if (keepOriginal)
                {
                    var bytes = pristine[entry.FullName];
                    target.Write(bytes, 0, bytes.Length);
                }
                else
                {
                    using var stream = entry.Open();
                    stream.CopyTo(target);
                }
            }
        }

        return result.ToArray();
    }

    private static List<OpenXmlElement> BuildBody(
        DocumentModelDto model,
        MainDocumentPart part,
        Dictionary<string, Block> index,
        Inventory inventory,
        out int preserved,
        out int rewritten)
    {
        var writer = new ParagraphWriter(part, inventory);
        var used = new HashSet<string>(StringComparer.Ordinal);
        var elements = new List<OpenXmlElement>();

        preserved = 0;
        rewritten = 0;

        foreach (var slot in Flatten(model.Doc))
        {
            var oid = OidOf(slot.Identity);

            // Um `oid` repetido é bloco colado: preservar o mesmo XML duas
            // vezes duplicaria âncoras de comentário e ids de revisão, então a
            // partir da segunda ocorrência ele é tratado como bloco novo.
            if (oid is not null && used.Add(oid) && index.TryGetValue(oid, out var block) &&
                string.Equals(
                    block.Extracted.Fingerprint(),
                    slot.Identity.Fingerprint(),
                    StringComparison.Ordinal))
            {
                elements.Add(block.Source.CloneNode(true));
                preserved++;
                continue;
            }

            // O XML original do bloco editado ainda serve para o que este
            // escritor não sabe gerar: os objetos ancorados seguem para o
            // parágrafo reescrito em vez de sumirem com ele.
            var source = oid is not null && index.TryGetValue(oid, out var edited) ? edited.Source : null;

            foreach (var element in writer.Write(slot.Content, slot.List, source)) elements.Add(element);
            rewritten++;

            if (source is not null) NoteWhatWasInside(source, inventory);
        }

        if (elements.Count == 0) elements.Add(new Paragraph());
        return elements;
    }

    /// <summary>
    /// Achata a árvore do editor na sequência de blocos que o corpo do DOCX
    /// espera, desembrulhando listas.
    /// </summary>
    /// <remarks>
    /// No editor uma lista é um nó com itens dentro; no OOXML são parágrafos
    /// irmãos, cada um apontando a mesma numeração. É por isso que o `oid` mora
    /// no `listItem`: ele é que corresponde a um `w:p`.
    /// </remarks>
    /// <param name="Identity">
    /// O nó extraído na abertura — é contra ele que a impressão digital é
    /// comparada. Para um item de lista é o `listItem`, não o parágrafo de
    /// dentro: comparar coisas de tipos diferentes nunca daria igual, e **nada
    /// seria preservado**, transformando a edição cirúrgica em regeneração
    /// completa sem que nada falhasse visivelmente.
    /// </param>
    /// <param name="Content">O nó a gravar quando não houver preservação.</param>
    private sealed record Slot(Node Identity, Node Content, ParagraphWriter.ListContext? List);

    private static IEnumerable<Slot> Flatten(
        Node doc,
        ParagraphWriter.ListContext? inherited = null)
    {
        foreach (var node in doc.Content ?? [])
        {
            switch (node.Type)
            {
                case "bulletList":
                case "orderedList":
                {
                    var level = (inherited?.Level ?? -1) + 1;
                    var numbering = NumberingOf(node) ?? inherited?.NumberingId ?? 0;

                    foreach (var item in node.Content ?? [])
                    {
                        if (item.Type != "listItem") continue;

                        var context = new ParagraphWriter.ListContext(numbering, level);
                        var paragraphs = (item.Content ?? [])
                            .Where(child => child.Type is "paragraph" or "heading").ToList();
                        var nested = (item.Content ?? [])
                            .Where(child => child.Type is "bulletList" or "orderedList");

                        // O item dá a identidade; o parágrafo de dentro dá o
                        // conteúdo.
                        if (paragraphs.Count > 0) yield return new Slot(item, paragraphs[0], context);

                        foreach (var extra in paragraphs.Skip(1))
                        {
                            yield return new Slot(extra, extra, context);
                        }

                        foreach (var child in nested)
                        {
                            var wrapper = Node.Of("doc");
                            wrapper.Content = [child];
                            foreach (var deeper in Flatten(wrapper, context)) yield return deeper;
                        }
                    }

                    break;
                }

                case "blockquote":
                {
                    // Sem citação no OOXML: vira recuo, que é o que o Word faz.
                    foreach (var child in node.Content ?? [])
                    {
                        yield return new Slot(child, child, inherited);
                    }

                    break;
                }

                default:
                    yield return new Slot(node, node, inherited);
                    break;
            }
        }
    }

    private static int? NumberingOf(Node list)
    {
        if (list.Attrs is not null && list.Attrs.TryGetValue("numId", out var value) && value is not null &&
            value.GetValueKind() == System.Text.Json.JsonValueKind.Number)
        {
            return value.GetValue<int>();
        }

        return null;
    }

    private static string? OidOf(Node node)
    {
        if (node.Attrs is null || !node.Attrs.TryGetValue("oid", out var value) || value is null) return null;
        return value.GetValueKind() == System.Text.Json.JsonValueKind.String ? value.GetValue<string>() : null;
    }

    /// <summary>
    /// O que havia dentro de um bloco editado e não sabemos regenerar.
    /// </summary>
    /// <remarks>
    /// Esta é a **perda de verdade**, e ela é detectada por comparação em vez
    /// de por palpite: o aviso pode dizer "você editou um parágrafo que tinha
    /// um comentário ancorado", em vez de um alerta genérico na abertura que o
    /// usuário aprende a ignorar.
    /// </remarks>
    private static void NoteWhatWasInside(OpenXmlElement original, Inventory inventory)
    {
        if (original.Descendants<CommentRangeStart>().Any() ||
            original.Descendants<CommentReference>().Any())
        {
            inventory.NoteLoss("comentário ancorado num parágrafo que você editou");
        }

        if (original.Descendants<InsertedRun>().Any() || original.Descendants<DeletedRun>().Any())
        {
            inventory.NoteLoss("marcas de revisão num parágrafo que você editou");
        }

        if (original.Descendants<FootnoteReference>().Any())
        {
            inventory.NoteLoss("nota de rodapé num parágrafo que você editou");
        }

        if (original.Descendants<FieldChar>().Any())
        {
            inventory.NoteLoss("campo calculado num parágrafo que você editou");
        }

        // Objeto ancorado não entra mais aqui: ele é copiado do original para o
        // parágrafo reescrito. Sobra o que continua sem volta — o VML antigo
        // (`w:pict`) e o desenho preso a um run que também traz texto, que
        // copiado traria a frase junto.
        if (original.Descendants<Picture>().Any() ||
            original.Elements<Run>().Any(run =>
                run.Descendants<WordDrawing.Anchor>().Any() && !ParagraphWriter.IsAnchoredOnly(run)))
        {
            inventory.NoteLoss("forma ou caixa de texto num parágrafo que você editou");
        }
    }

    private static void ApplyPageSetup(SectionProperties section, PageSetupDto page)
    {
        var landscape = string.Equals(page.Orientation, "landscape", StringComparison.Ordinal);
        var (shortSide, longSide) = string.Equals(page.Size, "Letter", StringComparison.Ordinal)
            ? (12240U, 15840U)
            : (11906U, 16838U);

        var size = section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();
        if (size is null)
        {
            size = new DocumentFormat.OpenXml.Wordprocessing.PageSize();
            section.PrependChild(size);
        }

        size.Width = landscape ? longSide : shortSide;
        size.Height = landscape ? shortSide : longSide;
        size.Orient = landscape ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;

        var margin = section.GetFirstChild<PageMargin>();
        if (margin is null)
        {
            margin = new PageMargin();
            section.InsertAfter(margin, size);
        }

        margin.Top = Twips(page.Margins.Top);
        margin.Bottom = Twips(page.Margins.Bottom);
        margin.Left = (uint)Math.Max(0, Twips(page.Margins.Left));
        margin.Right = (uint)Math.Max(0, Twips(page.Margins.Right));
    }

    private static int Twips(double millimeters) =>
        (int)Math.Round(millimeters * 1440 / 25.4, MidpointRounding.AwayFromZero);

    private static WordprocessingDocument OpenEditable(Stream stream)
    {
        try
        {
            return WordprocessingDocument.Open(stream, isEditable: true);
        }
        catch (Exception problem) when (problem is not DocxException)
        {
            throw new DocxException(
                "Não foi possível gravar sobre o arquivo original. Ele pode ter sido alterado ou danificado.",
                problem);
        }
    }
}
