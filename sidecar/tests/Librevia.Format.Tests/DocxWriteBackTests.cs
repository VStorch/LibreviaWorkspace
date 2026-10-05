using System.Reflection;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// What saving does to an edited block, where loss happens; what nobody touched belongs to <see
/// cref="DocxRoundTripTests"/>.
/// </summary>
public class DocxWriteBackTests
{
    private static Node BlockOf(DocumentModelDto model, int index) => model.Doc.Content![index];

    /// <summary>
    /// Every fixture is born inside the OOXML schema, so <see cref="Roundtrip.AssertSchema"/> never
    /// blames the writer for a test defect.
    /// </summary>
    [Fact]
    public void TodoFixtureNasceDentroDoEsquema()
    {
        var invalid = new List<string>();

        foreach (var factory in typeof(Fixtures)
                     .GetMethods(BindingFlags.Public | BindingFlags.Static)
                     .Where(method => method.ReturnType == typeof(byte[]) && method.GetParameters().Length == 0)
                     .OrderBy(method => method.Name, StringComparer.Ordinal))
        {
            var bytes = (byte[])factory.Invoke(null, null)!;

            // A fixture that is not a package is an image, like `SquarePng`.
            if (bytes.Length < 2 || bytes[0] != 'P' || bytes[1] != 'K') continue;

            try
            {
                Roundtrip.AssertSchema(bytes);
            }
            catch (Exception failure)
            {
                invalid.Add($"{factory.Name}: {failure.Message}");
            }
        }

        Assert.Empty(invalid);
    }

    [Fact]
    public void EditarParagrafoPreservaEstiloEspacamentoEFundo()
    {
        // Fixing a comma in the heading does not touch the rest of `w:pPr`.
        var original = Fixtures.WithFormattedParagraph();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Título formatado", "Título corrigido."));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        // The document's style, not a name invented from the level.
        Assert.Contains("<w:pStyle w:val=\"Ttulo1\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Heading1", xml, StringComparison.Ordinal);

        Assert.Contains("w:keepNext", xml, StringComparison.Ordinal);
        Assert.Contains("w:fill=\"C00000\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:before=\"360\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:after=\"180\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:line=\"271\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:sz w:val=\"20\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:ascii=\"Arial\"", xml, StringComparison.Ordinal);

        // What the editor does not even know passes through: the border and the tab.
        Assert.Contains("w:pBdr", xml, StringComparison.Ordinal);
        Assert.Contains("w:pos=\"4500\"", xml, StringComparison.Ordinal);

        Assert.Equal("Título corrigido.Parágrafo comum.", Roundtrip.TextOf(Roundtrip.Open(saved)));
        Assert.Equal(1, result.RewrittenBlocks);
    }

    [Fact]
    public void ParagrafoQueDeixouDeSerTituloPerdeOEstiloDeTitulo()
    {
        // Keeping the `styleId` does not leave a heading look on what became a paragraph.
        var original = Fixtures.WithFormattedParagraph();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var block = BlockOf(model, 0);
        Assert.Equal("heading", block.Type);

        var paragraph = Node.Of("paragraph");
        paragraph.Attrs = block.Attrs;
        paragraph.Content = block.Content;
        model.Doc.Content![0] = paragraph;

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.DoesNotContain("w:pStyle", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void OBlocoEditadoNaoRepeteOEstiloNosRunsNemNoPPr()
    {
        // The band run's marks come from the `Faixa` style, and writing them as direct would repeat
        // the style in every `w:r`.
        var original = Fixtures.WithStyles();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        Assert.True(Roundtrip.EditFirstTextContaining(model, "Informações", "Dados"));

        var saved = Roundtrip.Save(original, model);
        Assert.Equal(1, saved.Result.RewrittenBlocks);

        // Only the edited block; the others go back byte for byte.
        var all = Roundtrip.XmlOf(saved.Bytes);
        var start = all.IndexOf("<w:p>", StringComparison.Ordinal);
        var xml = all[start..(all.IndexOf("</w:p>", start, StringComparison.Ordinal) + 6)];
        Assert.Contains("<w:pStyle w:val=\"Faixa\" />", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:rPr>", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("w:shd", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:jc", xml, StringComparison.Ordinal);

        var banner = Roundtrip.OpenFlat(saved.Bytes).Doc.Content![0];
        Assert.Equal("#943634", banner.Attrs!["background"]!.GetValue<string>());
    }

    [Fact]
    public void TrechoSemNegritoNumEstiloNegritoSegueOQueATelaMostra()
    {
        // The style rule makes the block bold: `w:b w:val="0"` would make the file diverge from the
        // screen.
        var original = Fixtures.WithStyles();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var text = BlockOf(model, 0).Content![0];
        text.Marks = (text.Marks ?? []).Where(mark => mark.Type != "bold").ToList();
        text.Text = "Sem a marca.";

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);
        Assert.DoesNotMatch("<w:b w:val=", xml);
    }

    [Fact]
    public void ODesligadoSobreOEstiloVaiEVoltaDoArquivo()
    {
        // A mark with `off` is `w:b w:val="0"`, and reread it becomes the same mark.
        var original = Fixtures.WithStyles();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var text = BlockOf(model, 0).Content![0];
        text.Marks = [.. (text.Marks ?? []).Where(mark => mark.Type != "bold"), RunReader.Off("bold"), Mark.Of("charStyle", "styleId", "Destaque")];
        text.Text = "Sem negrito.";

        var saved = Roundtrip.Save(original, model).Bytes;
        var xml = Roundtrip.XmlOf(saved);
        Assert.Matches("<w:b w:val=\"(0|false)\" />", xml);
        Assert.Contains("<w:rStyle w:val=\"Destaque\" />", xml, StringComparison.Ordinal);

        var marks = Roundtrip.Open(saved).Doc.Content![0].Content![0].Marks!;
        Assert.Contains(marks, mark => mark.Type == "bold" && mark.Attrs?["off"]?.GetValue<bool>() == true);
        Assert.Contains(marks, mark => mark.Type == "charStyle");

        // An old draft knows neither.
        Assert.DoesNotContain(
            Roundtrip.OpenFlat(saved).Doc.Content![0].Content![0].Marks!,
            mark => mark.Type == "charStyle" || mark.Attrs?.ContainsKey("off") == true);
    }

    [Fact]
    public void FormatacaoLimpaSaiDoArquivo()
    {
        // Direct formatting that left the node was cleared by someone and goes; border and tab
        // stay.
        var original = Fixtures.WithFormattedParagraph();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var attrs = BlockOf(model, 0).Attrs!;
        foreach (var name in (string[])["background", "keepNext", "spaceBefore", "spaceAfter", "lineHeight", "textAlign"])
        {
            attrs[name] = null;
        }

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.DoesNotContain("w:shd", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("w:keepNext", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("w:spacing", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:jc", xml, StringComparison.Ordinal);
        Assert.Contains("w:pBdr", xml, StringComparison.Ordinal);
        Assert.Contains("w:tabs", xml, StringComparison.Ordinal);
        Assert.Contains("<w:pStyle w:val=\"Ttulo1\" />", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void NoWordEmPortuguesOTituloApontaTtulo1()
    {
        // The id is translated, and the name `heading 1` is not: `word/styles.xml` goes back
        // untouched.
        var original = DocxTemplateTests.WithLocalizedHeadingStyle();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var paragraph = BlockOf(model, 0);
        var heading = Node.Of("heading").With("level", 1).With("indent", 0);
        heading.Content = paragraph.Content;
        model.Doc.Content![0] = heading;

        var saved = Roundtrip.Save(original, model).Bytes;

        Assert.Contains("<w:pStyle w:val=\"Ttulo1\" />", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        Assert.Equal(
            Roundtrip.XmlOf(original, "word/styles.xml"),
            Roundtrip.XmlOf(saved, "word/styles.xml"));
    }

    [Fact]
    public void EstiloQueOPacoteNaoTemECopiadoDoModeloEmbutido()
    {
        // `ListParagraph` came from the new document, and this package does not define it.
        var original = DocxTemplateTests.WithLocalizedHeadingStyle();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        BlockOf(model, 0).With("styleId", "ListParagraph");
        BlockOf(model, 0).Content![0].Text = "Recuado pelo estilo.";

        var saved = Roundtrip.Save(original, model).Bytes;
        var styles = Roundtrip.XmlOf(saved, "word/styles.xml");

        Assert.Contains("w:styleId=\"ListParagraph\"", styles, StringComparison.Ordinal);
        Assert.Contains("<w:pStyle w:val=\"ListParagraph\" />", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        // And the unstyled paragraph nobody touched does not gain a `w:pStyle`.
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(Roundtrip.XmlOf(saved), "<w:pStyle"));
    }

    [Fact]
    public void DiminuirORecuoAteZeroChegaAoArquivo()
    {
        // A zeroed indent leaves `w:pPr`, and the paragraph goes back to the style's indent.
        var original = Fixtures.WithFormattedParagraph();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var block = BlockOf(model, 0);
        block.Attrs!["indentMm"] = null;
        block.Attrs["firstLineMm"] = null;

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.DoesNotContain("<w:ind", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("w:firstLine=", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void RecuoNegativoDoArquivoSobreviveAEdicao()
    {
        // The reader only emits positive indents: a negative one nobody touched stays.
        var original = Fixtures.WithNegativeIndent();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "para fora", "corrigido"));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("w:left=\"-284\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void RecuoAusenteNaoViraAtributoVazio()
    {
        // A `string`'s `null` becomes `StringValue(null)`, and the SDK writes `w:right=""`, which
        // Word refuses.
        var original = Fixtures.WithFormattedParagraph();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Título formatado", "Título corrigido."));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("w:left=\"720\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("=\"\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void MarcadorDoParagrafoSobreviveAEdicaoDoTexto()
    {
        // A bookmark is the target of cross-references, indexes and internal links, and the model
        // does not represent it.
        var original = Fixtures.WithBookmarkAroundParagraph();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Parágrafo marcado", "Parágrafo corrigido."));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Contains("w:name=\"CapituloUm\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:bookmarkEnd", xml, StringComparison.Ordinal);
        Assert.Equal("Parágrafo corrigido.Parágrafo comum.", Roundtrip.TextOf(Roundtrip.Open(saved)));
        Assert.Equal(1, result.RewrittenBlocks);
    }

    [Fact]
    public void BlocoColadoComOMesmoOidNaoDuplicaOMarcador()
    {
        // A repeated `oid` is a pasted block: from the second one on the bookmark is not copied, or
        // two bookmarks would share an id.
        var original = Fixtures.WithBookmarkAroundParagraph();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        model.Doc.Content!.Insert(1, BlockOf(model, 0));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, Roundtrip.Clone(model)).Bytes);

        Assert.Equal(1, Regex.Matches(xml, "<w:bookmarkStart").Count);
        Assert.Equal(1, Regex.Matches(xml, "<w:bookmarkEnd").Count);
    }

    [Fact]
    public void ItemDeListaEditadoContinuaApontandoAMesmaNumeracao()
    {
        // `w:numId w:val="0"` means "no numbering".
        var original = Fixtures.WithBulletList();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Primeiro item", "Item corrigido"));

        var (saved, _) = Roundtrip.Save(original, model);

        Assert.Contains("<w:numId w:val=\"1\"", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        Assert.DoesNotContain("<w:numId w:val=\"0\"", Roundtrip.XmlOf(saved), StringComparison.Ordinal);

        // And it is still a list on reopen.
        var reopened = Roundtrip.Open(saved);
        var list = Assert.Single(Roundtrip.Walk(reopened.Doc).Where(node => node.Type == "bulletList"));
        Assert.Equal(2, list.Content!.Count);
    }

    [Fact]
    public void ListaCriadaNoEditorGanhaNumeracaoNoArquivo()
    {
        // A document without lists has no `word/numbering.xml`: the part is created.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.DoesNotContain("word/numbering.xml", Roundtrip.PartsOf(original).Keys);

        var item = Node.Of("listItem", BlockOf(model, 1));
        model.Doc.Content![1] = Node.Of("bulletList", item);

        var (saved, _) = Roundtrip.Save(original, model);

        Assert.Contains("word/numbering.xml", Roundtrip.PartsOf(saved).Keys);

        var numbering = Roundtrip.XmlOf(saved, "word/numbering.xml");
        Assert.Contains("w:abstractNum", numbering, StringComparison.Ordinal);
        Assert.Contains("w:numFmt w:val=\"bullet\"", numbering, StringComparison.Ordinal);

        var reopened = Roundtrip.Open(saved);
        Assert.Contains(Roundtrip.Walk(reopened.Doc), node => node.Type == "bulletList");
    }

    [Fact]
    public void ListaOrdenadaNovaSaiNumerada()
    {
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        model.Doc.Content![1] = Node.Of("orderedList", Node.Of("listItem", BlockOf(model, 1)));

        var (saved, _) = Roundtrip.Save(original, model);

        Assert.Contains(
            "w:numFmt w:val=\"decimal\"",
            Roundtrip.XmlOf(saved, "word/numbering.xml"),
            StringComparison.Ordinal);
        Assert.Contains(Roundtrip.Walk(Roundtrip.Open(saved).Doc), node => node.Type == "orderedList");
    }

    [Fact]
    public void ListaDentroDeCelulaSobreviveAEdicaoDaCelula()
    {
        // Without list context, the writer would read that the paragraph stopped being an item.
        var original = Fixtures.WithListInsideTableCell();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Primeiro da célula", "Corrigido"));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Equal(2, Regex.Count(xml, "<w:numId w:val=\"1\""));
        Assert.Contains("Corrigido", xml, StringComparison.Ordinal);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void ListaColadaDeOutroDocumentoGanhaNumeracaoQueODestinoDefine()
    {
        // A pasted list carries the source document's `numId`, which this one does not define.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        model.Doc.Content![1] = Node.Of("bulletList", Node.Of("listItem", BlockOf(model, 1)))
            .With("numId", 7);

        var (saved, _) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.DoesNotContain("<w:numId w:val=\"7\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:numFmt w:val=\"bullet\"", Roundtrip.XmlOf(saved, "word/numbering.xml"), StringComparison.Ordinal);
        Assert.Contains(Roundtrip.Walk(Roundtrip.Open(saved).Doc), node => node.Type == "bulletList");
    }

    [Fact]
    public void SublistaOrdenadaDentroDeListaComMarcadorSaiNumerada()
    {
        // A numbered sublist inside a bulleted one does not inherit the outer `numId`.
        var original = Fixtures.WithBulletList();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var item = model.Doc.Content![1].Content![0];
        (item.Content ??= []).Add(Node.Of(
            "orderedList",
            Node.Of("listItem", Node.Of("paragraph", new Node { Type = "text", Text = "Sub-item" }))));

        var (saved, _) = Roundtrip.Save(original, model);

        Assert.Contains(
            "w:numFmt w:val=\"decimal\"",
            Roundtrip.XmlOf(saved, "word/numbering.xml"),
            StringComparison.Ordinal);

        var reopened = Roundtrip.Open(saved);
        Assert.Contains(Roundtrip.Walk(reopened.Doc), node => node.Type == "orderedList");
    }

    [Fact]
    public void ImagemInseridaPelaBarraSobreviveAoSalvar()
    {
        // A toolbar image is a block and arrives without a measure.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var source = "data:image/png;base64," + Convert.ToBase64String(Fixtures.SquarePng());
        model.Doc.Content!.Add(Node.Of("image").With("src", source));

        var (saved, result) = Roundtrip.Save(original, model);

        Assert.Empty(result.Inventory.Lost);
        Assert.Contains("w:drawing", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        Assert.Contains(Roundtrip.PartsOf(saved).Keys, name => name.Contains("image", StringComparison.Ordinal));

        // The measure comes from the PNG header, 4 × 4.
        var image = Assert.Single(Roundtrip.Walk(Roundtrip.Open(saved).Doc).Where(node => node.Type == "image"));
        Assert.Equal(4, image.Attrs!["width"]!.GetValue<int>());
        Assert.Equal(4, image.Attrs["height"]!.GetValue<int>());
    }

    [Fact]
    public void ImagemMaiorQueAColunaEncolheNaProporcao()
    {
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var source = "data:image/png;base64," + Convert.ToBase64String(Fixtures.SquarePng());
        model.Doc.Content!.Add(Node.Of("image")
            .With("src", source)
            .With("width", 4000)
            .With("height", 2000));

        var saved = Roundtrip.Save(original, model).Bytes;
        var image = Assert.Single(Roundtrip.Walk(Roundtrip.Open(saved).Doc).Where(node => node.Type == "image"));

        var width = image.Attrs!["width"]!.GetValue<int>();
        var height = image.Attrs["height"]!.GetValue<int>();

        // Fits the column of an A4 with 2.5 cm margins, at a 2:1 ratio.
        Assert.InRange(width, 500, 650);
        Assert.InRange((double)width / height, 1.9, 2.1);
    }

    [Fact]
    public void ImagemComDataUriSemVirgulaNaoDerrubaAGravacao()
    {
        // `data:image/png` without a comma: one image must not bring the save down.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        model.Doc.Content!.Add(Node.Of("image").With("src", "data:image/png"));

        var (saved, result) = Roundtrip.Save(original, model);

        Assert.Contains(result.Inventory.Lost, message => message.Contains("imagem", StringComparison.Ordinal));
        Assert.Contains("Primeiro parágrafo", Roundtrip.TextOf(Roundtrip.Open(saved)), StringComparison.Ordinal);
    }

    [Fact]
    public void IdDeDesenhoNaoDependeDeQuantasGravacoesJaAconteceram()
    {
        // The id depends neither on how many saves the process has done nor collides with the
        // file's, which Word shows as a damaged document.
        var original = Fixtures.WithInlineImage(1001);

        var first = DrawingIdsOf(original);
        var second = DrawingIdsOf(original);

        Assert.Equal(first, second);
        Assert.Equal(first.Distinct(), first);
        Assert.Contains("1001", first);
        Assert.All(first, id => Assert.True(int.Parse(id) >= 1001));
    }

    /// <summary>The `wp:docPr/@id`s of the document saved with a new image.</summary>
    private static List<string> DrawingIdsOf(byte[] original)
    {
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var source = "data:image/png;base64," + Convert.ToBase64String(Fixtures.SquarePng());
        model.Doc.Content!.Add(Node.Of("image").With("src", source));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);
        return [.. Regex.Matches(xml, "docPr id=\"(\\d+)\"").Select(match => match.Groups[1].Value)];
    }

    [Fact]
    public void TabelaEditadaPreservaEstiloLargurasEMesclagem()
    {
        var original = Fixtures.WithStyledTable();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Dado B", "Dado corrigido"));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("w:tblStyle w:val=\"GradeMedia3\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:tblGrid", xml, StringComparison.Ordinal);
        Assert.Contains("w:w=\"4000\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:tblHeader", xml, StringComparison.Ordinal);
        Assert.Contains("w:fill=\"D9D9D9\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:vMerge", xml, StringComparison.Ordinal);
        Assert.Contains("w:val=\"double\"", xml, StringComparison.Ordinal);
        Assert.Contains("Dado corrigido", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void TabelaAninhadaSobreviveAAberturaEAGravacao()
    {
        var original = Fixtures.WithNestedTable();
        var model = Roundtrip.Open(original);

        // First on screen: the inner table's text reaches the editor.
        Assert.Contains("Dentro da tabela de dentro", Roundtrip.TextOf(model), StringComparison.Ordinal);

        var edited = Roundtrip.Clone(model);
        Assert.True(Roundtrip.EditFirstTextContaining(edited, "Antes da aninhada", "Antes, corrigido"));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, edited).Bytes);

        Assert.Contains("GradeInterna", xml, StringComparison.Ordinal);
        Assert.Contains("Dentro da tabela de dentro", xml, StringComparison.Ordinal);
        Assert.Contains("Antes, corrigido", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void CelulaQueTerminaEmTabelaGanhaParagrafoNoFim()
    {
        // An editor cell may end in a nested table; Word's `w:tc` must end in a `w:p`.
        var original = Fixtures.WithNestedTable();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var cell = model.Doc.Content![0].Content![0].Content![0];
        cell.Content!.RemoveAt(cell.Content.Count - 1);

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.DoesNotContain("Depois da aninhada", xml, StringComparison.Ordinal);
        Assert.Contains("</w:tbl><w:p /></w:tc>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void LinhaInseridaNoMeioRegistraPerdaEDescartaAMesclagem()
    {
        // An inserted row misaligns positions and would take `w:vMerge` to the wrong row.
        var original = Fixtures.WithStyledTable();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var rows = model.Doc.Content![0].Content!;
        rows.Insert(1, Node.Of(
            "tableRow",
            Node.Of("tableCell", Node.Of("paragraph")),
            Node.Of("tableCell", Node.Of("paragraph"))));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Contains(result.Inventory.Lost, message => message.Contains("tabela", StringComparison.Ordinal));
        Assert.DoesNotContain("w:vMerge", xml, StringComparison.Ordinal);

        Assert.Equal(3, Roundtrip.Open(saved).Doc.Content![0].Content!.Count);
    }

    [Fact]
    public void LarguraDeColunaArrastadaChegaAoArquivo()
    {
        // The divider dragged in TableKit reaches `w:tblGrid`.
        var original = Fixtures.WithTable();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var table = model.Doc.Content!.First(block => block.Type == "table");
        foreach (var row in table.Content!)
        {
            row.Content![0].With("colwidth", new JsonArray(200));
            row.Content[1].With("colwidth", new JsonArray(400));
        }

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        // 15 twips per pixel.
        Assert.Contains("<w:gridCol w:w=\"3000\" /><w:gridCol w:w=\"6000\" />", xml, StringComparison.Ordinal);

        // The cell width follows, or it contradicts the grid and Word picks one.
        Assert.Contains("w:tcW w:w=\"3000\"", xml, StringComparison.Ordinal);

        // Word honors the grid like the screen with `table-layout: fixed`.
        Assert.Contains("w:tblLayout w:type=\"fixed\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void EditarTextoDaCelulaNaoMexeNaGradeDeColunas()
    {
        // Twip→pixel rounds (4675 twips is 311.67 px): an unresized grid stays.
        var original = Fixtures.WithTable();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "A1", "A1 corrigido"));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("w:w=\"4675\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("w:tblLayout", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void TabelaInseridaNaTelaNasceComGradeDeColunas()
    {
        // The schema requires `w:tblGrid`: equal columns across the text column, as Word does on
        // insert.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        model.Doc.Content!.Add(Node.Of(
            "table",
            Node.Of(
                "tableRow",
                Node.Of("tableHeader", Node.Of("paragraph")),
                Node.Of("tableHeader", Node.Of("paragraph")))));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("<w:tblGrid>", xml, StringComparison.Ordinal);

        // A row made entirely of `tableHeader` is the one that repeats at the top of each page.
        Assert.Contains("w:tblHeader", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void LinhaDeCabecalhoDesligadaSaiDoArquivo()
    {
        // `w:trPr` also carries height and revision: turning the flag off does not take it along.
        var original = Fixtures.WithStyledTable();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var first = model.Doc.Content!.First(block => block.Type == "table").Content![0];
        first.Content = [.. first.Content!.Select(cell => new Node
        {
            Type = "tableCell",
            Attrs = cell.Attrs,
            Content = cell.Content,
        })];

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Dado B", "Dado corrigido"));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.DoesNotContain("w:tblHeader", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void SombreamentoEBordaEscolhidosNaTelaChegamAoArquivo()
    {
        var original = Fixtures.WithTable();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var cell = model.Doc.Content!.First(block => block.Type == "table").Content![0].Content![0];
        cell.With("shading", "#d9d9d9");
        cell.With("borders", "top:double,1.5,#ff0000;bottom:none,0.5,#000000");

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("w:fill=\"D9D9D9\"", xml, StringComparison.Ordinal);

        // 1.5 pt is 12 eighths, the `w:sz` unit.
        Assert.Contains("<w:top w:val=\"double\" w:color=\"FF0000\" w:sz=\"12\" />", xml, StringComparison.Ordinal);

        // Only `w:nil` removes a border the table asked for.
        Assert.Contains("w:bottom w:val=\"nil\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void CelulaNaoFormatadaConservaTramaEEstiloExotico()
    {
        // As long as nobody formats the cell, the pattern and the exotic border come back from the
        // original XML.
        var original = Fixtures.WithPatternedCell();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Com trama", "Corrigido"));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Contains("w:val=\"pct25\"", xml, StringComparison.Ordinal);
        Assert.Contains("thickThinSmallGap", xml, StringComparison.Ordinal);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void LarguraArrastadaNumaTabelaNovaChegaAoArquivo()
    {
        // TableKit only puts `colwidth` on the dragged column.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var cell = (string text, int? width) =>
        {
            var node = Node.Of("tableCell").With("colspan", 1).With("rowspan", 1);
            if (width is { } px) node.With("colwidth", new JsonArray(px));
            node.Content = [Node.Of("paragraph")];
            node.Content[0].Content = [new Node { Type = "text", Text = text }];
            return node;
        };

        var row = Node.Of("tableRow");
        row.Content = [cell("A", 300), cell("B", null), cell("C", null)];
        var table = Node.Of("table");
        table.Content = [row];
        model.Doc.Content!.Add(table);

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        // 300 px is 4500 twips; the others share what is left of the text column.
        var grid = Regex.Matches(xml, "<w:gridCol w:w=\"(\\d+)\"").Select(match => match.Groups[1].Value).ToList();
        Assert.Equal(3, grid.Count);
        Assert.Equal("4500", grid[0]);
        Assert.Equal(grid[1], grid[2]);
    }

    [Fact]
    public void ArrastarUmaColunaNaoMexeNaMedidaDasOutras()
    {
        // A neighbour nobody touched does not go through `px × 15` (2000 twips is 133.33 px).
        var original = Fixtures.WithThreeColumns();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var table = model.Doc.Content!.First(block => block.Type == "table");
        foreach (var row in table.Content!) row.Content![1].With("colwidth", new JsonArray(250));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("<w:gridCol w:w=\"2000\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:gridCol w:w=\"3750\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:gridCol w:w=\"4000\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("1995", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void ColunaInseridaRecalculaAGrade()
    {
        // After inserting a column, TableKit leaves the new cell without a width.
        var original = Fixtures.WithThreeColumns();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var table = model.Doc.Content!.First(block => block.Type == "table");
        foreach (var row in table.Content!)
        {
            var added = Node.Of("tableCell").With("colspan", 1).With("rowspan", 1).With("colwidth", new JsonArray(0));
            added.Content = [Node.Of("paragraph")];
            row.Content!.Add(added);
        }

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Equal(4, Regex.Matches(xml, "<w:gridCol ").Count);
        Assert.Contains("<w:gridCol w:w=\"2000\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:gridCol w:w=\"4000\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void CabecalhoLigadoNumaLinhaComRevisaoFicaAntesDaRevisao()
    {
        // In `w:trPr` the revision (`w:ins`, `w:del`, `w:trPrChange`) closes the sequence.
        var original = Fixtures.WithInsertedRow();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var row = model.Doc.Content!.First(block => block.Type == "table").Content![0];
        row.Content = [.. row.Content!.Select(cell => new Node { Type = "tableHeader", Attrs = cell.Attrs, Content = cell.Content })];

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        var header = xml.IndexOf("<w:tblHeader", StringComparison.Ordinal);
        var inserted = xml.IndexOf("<w:ins ", StringComparison.Ordinal);
        Assert.True(header > 0 && inserted > header, "o w:tblHeader saiu depois do w:ins");
    }

    [Fact]
    public void LinhaQueNaoCobreAGradeNaoEncolheAGrade()
    {
        // With `w:gridBefore`, the model grid must have the file's columns.
        var original = Fixtures.WithGridBefore();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var table = model.Doc.Content!.First(block => block.Type == "table");
        var recuada = table.Content![0].Content![0];

        // The indented cell has the width of the grid's second column, 3000 twips.
        Assert.Equal(200, recuada.Attrs!["colwidth"]![0]!.GetValue<int>());

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Cheia A", "Cheia corrigida"));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Equal(3, Regex.Matches(xml, "<w:gridCol ").Count);
        Assert.DoesNotContain("w:tblLayout", xml, StringComparison.Ordinal);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void LarguraQueNaoCabeNaGradeDoArquivoEntraNoInventario()
    {
        // A grid the model does not fully describe: the file's stays, with a warning.
        var original = Fixtures.WithGridBefore();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var table = model.Doc.Content!.First(block => block.Type == "table");
        table.Content![0].Content![0].With("colwidth", new JsonArray(250));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Contains("<w:gridCol w:w=\"2000\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:gridCol w:w=\"3000\"", xml, StringComparison.Ordinal);
        Assert.Contains(result.Inventory.Lost, message => message.Contains("largura", StringComparison.Ordinal));
    }

    [Fact]
    public void FormatarUmLadoConservaOQueOModeloNaoRepresenta()
    {
        // Adding the bottom border does not erase the diagonal, inner border, `w:space` or theme
        // color.
        var original = Fixtures.WithRichCellBorders();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var cell = model.Doc.Content!.First(block => block.Type == "table").Content![0].Content![0];
        var borders = cell.Attrs!["borders"]!.GetValue<string>();
        cell.With("borders", borders.Replace("bottom:none,0.5,#000000", "bottom:single,1,#000000", StringComparison.Ordinal));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Contains("w:tl2br", xml, StringComparison.Ordinal);
        Assert.Contains("w:insideH", xml, StringComparison.Ordinal);
        Assert.Contains("w:themeColor=\"accent1\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:space=\"0\"", xml, StringComparison.Ordinal);
        Assert.Matches("<w:bottom w:val=\"single\"[^>]*w:sz=\"8\"", xml);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void TrocarACorDoLadoTiraACorDeTemaDele()
    {
        // A theme color beats `w:color` in Word, and goes with the color change; `w:space` stays.
        var original = Fixtures.WithRichCellBorders();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var cell = model.Doc.Content!.First(block => block.Type == "table").Content![0].Content![0];
        var borders = cell.Attrs!["borders"]!.GetValue<string>();
        cell.With("borders", borders.Replace("top:single,0.5,#4472c4", "top:single,0.5,#ff0000", StringComparison.Ordinal));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.DoesNotContain("w:themeColor", xml, StringComparison.Ordinal);
        Assert.Contains("w:color=\"FF0000\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:space=\"0\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:tl2br", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void FormatarACelulaDaTramaRegistraAPerda()
    {
        // Changing the cell look erases the pattern and the exotic style, with a warning.
        var original = Fixtures.WithPatternedCell();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var cell = model.Doc.Content!.First(block => block.Type == "table").Content![0].Content![0];
        cell.With("shading", "#00ff00");
        cell.With("borders", "top:single,0.5,#000000");

        var result = Roundtrip.Save(original, model).Result;

        Assert.Contains(result.Inventory.Lost, message => message.Contains("trama", StringComparison.Ordinal));
        Assert.Contains(
            result.Inventory.Lost,
            message => message.Contains("estilo de borda", StringComparison.Ordinal));
    }

    [Fact]
    public void MesclagemVerticalFeitaNaTelaEntraNoInventario()
    {
        // TableKit's `rowspan` does not become `w:vMerge` yet: the least is to say so.
        var original = Fixtures.WithTable();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var rows = model.Doc.Content!.First(block => block.Type == "table").Content!;
        rows[0].Content![0].With("rowspan", 2);
        rows[1].Content!.RemoveAt(0);

        var result = Roundtrip.Save(original, model).Result;

        Assert.Contains(result.Inventory.Lost, message => message.Contains("mesclagem vertical", StringComparison.Ordinal));
    }

    [Fact]
    public void TextoAlternativoDaImagemVaiEVolta()
    {
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var source = "data:image/png;base64," + Convert.ToBase64String(Fixtures.SquarePng());

        model.Doc.Content!.Add(Node.Of("image")
            .With("src", source)
            .With("alt", "Organograma da diretoria")
            .With("align", "center"));

        var (saved, _) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Contains("descr=\"Organograma da diretoria\"", xml, StringComparison.Ordinal);

        // In OOXML a centered image is a centered paragraph with the image inside.
        Assert.Contains("w:jc w:val=\"center\"", xml, StringComparison.Ordinal);

        var back = Roundtrip.Open(saved);
        var image = Roundtrip.Walk(back.Doc).First(node => node.Type == "image");
        Assert.Equal("Organograma da diretoria", image.Attrs!["alt"]!.GetValue<string>());
    }

    [Fact]
    public void RedimensionarImagemDoArquivoMantemODesenhoOriginal()
    {
        // Resizing changes only the size: `wp:docPr`, the image part and what the writer does not
        // generate (effect, crop, border) stay.
        var original = Fixtures.WithInlineImage(7);
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var image = Roundtrip.Walk(model.Doc).Single(node => node.Type == "image");
        image.With("width", 300).With("height", 150).With("alt", "Gráfico de vendas");

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);
        var before = Roundtrip.XmlOf(original);
        var embed = Regex.Match(before, "r:embed=\"([^\"]+)\"").Groups[1].Value;

        Assert.Contains("docPr id=\"7\" name=\"Imagem 7\"", xml, StringComparison.Ordinal);
        Assert.Contains("descr=\"Gráfico de vendas\"", xml, StringComparison.Ordinal);
        Assert.Single(Regex.Matches(xml, "r:embed=\"([^\"]+)\""));
        Assert.Contains($"r:embed=\"{embed}\"", xml, StringComparison.Ordinal);
        Assert.Single(Roundtrip.PartsOf(saved).Keys, name => name.EndsWith(".png", StringComparison.Ordinal));

        // 300 × 150 px is 2857500 × 1428750 EMU, in `wp:extent` and in `a:ext`.
        Assert.Equal(2, Regex.Matches(xml, "cx=\"2857500\" cy=\"1428750\"").Count);

        Assert.Equal(
            Regex.Matches(before, "<w:p[ >]").Count,
            Regex.Matches(xml, "<w:p[ >]").Count);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void ImagemAncoradaNoFluxoNaoSeDuplicaAoEditarOParagrafo()
    {
        // An anchored image in the flow arrives as a paragraph image: the original `w:r` does not
        // come along.
        var original = Fixtures.WithAnchoredImageInTheFlow();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var image = Roundtrip.Walk(model.Doc).Single(node => node.Type == "image");
        image.With("width", 100).With("height", 50);

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Single(Regex.Matches(xml, "<wp:docPr "));
        Assert.Contains("<wp:anchor", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void MarcadorEntreLinhasContinuaEntreAsLinhas()
    {
        // Children that are not rows stay in place.
        var original = Fixtures.WithBookmarkBetweenRows();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Linha de cima", "Linha corrigida"));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        var firstRowEnd = xml.IndexOf("</w:tr>", StringComparison.Ordinal);
        var bookmark = xml.IndexOf("bookmarkStart", StringComparison.Ordinal);

        Assert.True(bookmark > firstRowEnd, "o marcador voltou para antes da primeira linha");
        Assert.True(
            xml.IndexOf("bookmarkEnd", StringComparison.Ordinal) > xml.LastIndexOf("</w:tr>", StringComparison.Ordinal),
            "o fim do marcador saiu de depois da última linha");
    }

    [Fact]
    public void CorDoCssViraHexadecimalDeSeisDigitos()
    {
        // Word reports `rgb(255, 0, 0)` as damaged, and draws color names as black.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var colored = new Node
        {
            Type = "text",
            Text = "colorido",
            Marks =
            [
                Mark.Of("textStyle", "color", "rgb(255, 0, 0)"),
                Mark.Of("highlight", "color", "yellow"),
            ],
        };

        BlockOf(model, 1).Content = [colored];
        BlockOf(model, 1).With("background", "#abc");

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Contains("w:color w:val=\"FF0000\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:fill=\"FFFF00\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:fill=\"AABBCC\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("rgb(", xml, StringComparison.Ordinal);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void EditarParagrafoMantemSobrescritoESubscrito()
    {
        var original = Fixtures.WithVerticalAlignment();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "O e m", "O e n"));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Contains("w:vertAlign w:val=\"superscript\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:vertAlign w:val=\"subscript\"", xml, StringComparison.Ordinal);
        Assert.Empty(result.Inventory.Lost);

        // And the reread document carries both marks.
        var reread = Roundtrip.Walk(Roundtrip.Open(saved).Doc)
            .SelectMany(node => node.Marks ?? [])
            .Select(mark => mark.Type)
            .ToList();

        Assert.Contains("superscript", reread);
        Assert.Contains("subscript", reread);
    }

    [Fact]
    public void CorQueNaoDaParaConverterEntraNoInventario()
    {
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        BlockOf(model, 1).Content =
        [
            new Node
            {
                Type = "text",
                Text = "exótico",
                Marks = [Mark.Of("textStyle", "color", "hsl(120, 50%, 50%)")],
            },
        ];

        var (saved, result) = Roundtrip.Save(original, model);

        Assert.Contains(result.Inventory.Lost, message => message.Contains("cor", StringComparison.Ordinal));
        Assert.DoesNotContain("hsl(", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
    }

    [Fact]
    public void NomeDeCorForaDaTabelaNaoViraHexadecimal()
    {
        // `fade` is valid hex, but not a color.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        BlockOf(model, 1).Content =
        [
            new Node
            {
                Type = "text",
                Text = "desbotado",
                Marks = [Mark.Of("textStyle", "color", "fade")],
            },
        ];

        var (saved, result) = Roundtrip.Save(original, model);

        Assert.Contains(result.Inventory.Lost, message => message.Contains("cor", StringComparison.Ordinal));
        Assert.DoesNotContain("FFAADD", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
    }

    [Fact]
    public void TamanhoDeFonteEmPixelsViraPontos()
    {
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        BlockOf(model, 1).Content =
        [
            new Node
            {
                Type = "text",
                Text = "dezesseis pixels",
                Marks = [Mark.Of("textStyle", "fontSize", "16px")],
            },
        ];

        // 16 px is 12 pt, and `w:sz` is in half-points.
        Assert.Contains(
            "w:sz w:val=\"24\"",
            Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes),
            StringComparison.Ordinal);
    }

    [Fact]
    public void CaractereDeControleNaoDerrubaAGravacao()
    {
        // `\u0001` does not exist in XML 1.0, not even escaped; `\u000B` and `\n` are breaks, and
        // the tab is `w:tab`.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        BlockOf(model, 1).Content =
        [
            new Node { Type = "text", Text = "a\u0001b\u000Bc\td" },
        ];

        var (saved, _) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Contains("<w:br />", xml, StringComparison.Ordinal);
        Assert.Contains("<w:tab />", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("\u0001", xml, StringComparison.Ordinal);

        Assert.Equal("Primeiro parágrafo.abc\tdTerceiro parágrafo.", Roundtrip.TextOf(Roundtrip.Open(saved)));
    }

    [Fact]
    public void PapelForaDeA4NaoEArredondadoAoSalvar()
    {
        // The model only names A4 and Letter: `w:sectPr` only changes when the page changes.
        var original = Fixtures.WithCustomPaper();
        var (saved, _) = Roundtrip.Save(original, Roundtrip.Clone(Roundtrip.Open(original)));

        var xml = Roundtrip.XmlOf(saved);
        Assert.Contains("w:w=\"8391\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:h=\"11907\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void PapelQueOModeloNaoNomeiaEAvisadoComoInvisivel()
    {
        // The panel shows A4 for an A5: the warning says the screen approximates.
        var inventory = DocxReader.Read(Fixtures.WithCustomPaper()).Inventory;

        Assert.Contains(inventory.Invisible, message => message.Contains("papel", StringComparison.Ordinal));
        Assert.Empty(inventory.Lost);
    }

    [Fact]
    public void GirarAFolhaMantemOPapelDoArquivo()
    {
        var original = Fixtures.WithCustomPaper();
        var opened = Roundtrip.Clone(Roundtrip.Open(original));
        var model = opened with { Page = opened.Page with { Orientation = "landscape" } };

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        // Continua A5: o lado curto virou a altura.
        Assert.Contains("w:w=\"11907\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:h=\"8391\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:orient=\"landscape\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void MudarAMargemRegravaAMargemESomenteEla()
    {
        var original = Fixtures.WithCustomPaper();
        var opened = Roundtrip.Clone(Roundtrip.Open(original));
        var model = opened with { Page = opened.Page with { Margins = opened.Page.Margins with { Top = 40 } } };

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("w:top=\"2268\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:w=\"8391\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void KeepsKeepLinesWhenTheParagraphIsEdited()
    {
        var original = Fixtures.WithKeepLines();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Linhas juntas", "Linhas ainda juntas."));

        var (saved, _) = Roundtrip.Save(original, model);
        Assert.Contains("<w:keepLines />", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
    }

    [Fact]
    public void KeepsWidowControlOffWhenTheParagraphIsEdited()
    {
        var original = Fixtures.WithKeepLines();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Viúva permitida", "Viúva ainda permitida."));

        var (saved, _) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);
        Assert.Matches("<w:widowControl w:val=\"(0|false)\"", xml);
        // A regular paragraph stays silent, on by default.
        Assert.Equal(1, xml.Split("w:widowControl").Length - 1);
    }

    [Fact]
    public void KeepsTheTableCellMarginsWhenTheTableIsEdited()
    {
        // `cellMargins` is read-only: the file's `w:tblCellMar` goes back as it was.
        var original = Fixtures.WithCellMargins();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Margem própria", "Margem própria, editada."));

        var (saved, _) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);
        Assert.Contains("<w:tblCellMar><w:top w:w=\"100\" w:type=\"dxa\" /></w:tblCellMar>", xml, StringComparison.Ordinal);
        Assert.Contains("Margem própria, editada.", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void TopAnchoredImageKeepsItsTextWhenTheParagraphIsEdited()
    {
        // The frame moved to the start of the paragraph in the editor, and once saved it is still a
        // single anchored one.
        var original = Fixtures.WithTextAroundTopAnchoredImage();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "registros OK", " registros conferidos"));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);
        Assert.Single(Regex.Matches(xml, "<wp:docPr "));
        Assert.Contains("<wp:anchor", xml, StringComparison.Ordinal);
        Assert.Contains("Múltiplos", xml, StringComparison.Ordinal);
        Assert.Contains("registros conferidos", xml, StringComparison.Ordinal);
    }
}
