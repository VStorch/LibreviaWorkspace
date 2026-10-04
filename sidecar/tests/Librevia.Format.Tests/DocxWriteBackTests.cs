using System.Reflection;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// O que a gravação faz com o bloco que foi editado, onde a perda acontece; o que
/// ninguém tocou é de <see cref="DocxRoundTripTests"/>.
/// </summary>
public class DocxWriteBackTests
{
    private static Node BlockOf(DocumentModelDto model, int index) => model.Doc.Content![index];

    /// <summary>
    /// Todo fixture nasce dentro do esquema OOXML, para que
    /// <see cref="Roundtrip.AssertSchema"/> nunca acuse o escritor por um defeito do teste.
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

            // Fixture que não é pacote é imagem, como `SquarePng`.
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
        // Corrigir uma vírgula no título não toca o resto do `w:pPr`.
        var original = Fixtures.WithFormattedParagraph();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Título formatado", "Título corrigido."));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        // O estilo do documento, e não um nome inventado a partir do nível.
        Assert.Contains("<w:pStyle w:val=\"Ttulo1\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Heading1", xml, StringComparison.Ordinal);

        Assert.Contains("w:keepNext", xml, StringComparison.Ordinal);
        Assert.Contains("w:fill=\"C00000\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:before=\"360\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:after=\"180\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:line=\"271\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:sz w:val=\"20\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:ascii=\"Arial\"", xml, StringComparison.Ordinal);

        // O que o editor nem conhece atravessa: a borda e a tabulação.
        Assert.Contains("w:pBdr", xml, StringComparison.Ordinal);
        Assert.Contains("w:pos=\"4500\"", xml, StringComparison.Ordinal);

        Assert.Equal("Título corrigido.Parágrafo comum.", Roundtrip.TextOf(Roundtrip.Open(saved)));
        Assert.Equal(1, result.RewrittenBlocks);
    }

    [Fact]
    public void ParagrafoQueDeixouDeSerTituloPerdeOEstiloDeTitulo()
    {
        // Preservar o `styleId` não deixa cara de título no que virou parágrafo.
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
        // As marcas do trecho da faixa vêm do estilo `Faixa`, e gravá-las como direto
        // repetiria o estilo em cada `w:r`.
        var original = Fixtures.WithStyles();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        Assert.True(Roundtrip.EditFirstTextContaining(model, "Informações", "Dados"));

        var saved = Roundtrip.Save(original, model);
        Assert.Equal(1, saved.Result.RewrittenBlocks);

        // Só o bloco editado; os outros voltam byte a byte.
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
        // A regra do estilo dá negrito ao bloco: `w:b w:val="0"` faria o arquivo divergir da tela.
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
        // A marca com `off` é `w:b w:val="0"`, e relida vira a mesma marca.
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

        // O rascunho antigo não conhece nem uma nem outra.
        Assert.DoesNotContain(
            Roundtrip.OpenFlat(saved).Doc.Content![0].Content![0].Marks!,
            mark => mark.Type == "charStyle" || mark.Attrs?.ContainsKey("off") == true);
    }

    [Fact]
    public void FormatacaoLimpaSaiDoArquivo()
    {
        // O direto que sumiu do nó foi limpo por alguém e sai; borda e tabulação ficam.
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
        // O id é traduzido, e o nome `heading 1` não: `word/styles.xml` volta intocado.
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
        // `ListParagraph` veio do documento novo, e este pacote não o define.
        var original = DocxTemplateTests.WithLocalizedHeadingStyle();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        BlockOf(model, 0).With("styleId", "ListParagraph");
        BlockOf(model, 0).Content![0].Text = "Recuado pelo estilo.";

        var saved = Roundtrip.Save(original, model).Bytes;
        var styles = Roundtrip.XmlOf(saved, "word/styles.xml");

        Assert.Contains("w:styleId=\"ListParagraph\"", styles, StringComparison.Ordinal);
        Assert.Contains("<w:pStyle w:val=\"ListParagraph\" />", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        // E o parágrafo sem estilo, que ninguém tocou, não ganha `w:pStyle`.
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(Roundtrip.XmlOf(saved), "<w:pStyle"));
    }

    [Fact]
    public void DiminuirORecuoAteZeroChegaAoArquivo()
    {
        // O recuo zerado sai do `w:pPr`, e o parágrafo volta ao recuo do estilo.
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
        // O leitor só emite recuo positivo: o negativo que ninguém tocou fica.
        var original = Fixtures.WithNegativeIndent();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "para fora", "corrigido"));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("w:left=\"-284\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void RecuoAusenteNaoViraAtributoVazio()
    {
        // O `null` de um `string` vira `StringValue(null)`, e o SDK grava `w:right=""`,
        // que o Word recusa.
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
        // O marcador é destino de referência cruzada, índice e link interno, e o modelo não o representa.
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
        // `oid` repetido é bloco colado: do segundo em diante não se copia o marcador,
        // senão dois marcadores teriam o mesmo id.
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
        // `w:numId w:val="0"` quer dizer "sem numeração".
        var original = Fixtures.WithBulletList();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Primeiro item", "Item corrigido"));

        var (saved, _) = Roundtrip.Save(original, model);

        Assert.Contains("<w:numId w:val=\"1\"", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        Assert.DoesNotContain("<w:numId w:val=\"0\"", Roundtrip.XmlOf(saved), StringComparison.Ordinal);

        // E continua lista ao reabrir.
        var reopened = Roundtrip.Open(saved);
        var list = Assert.Single(Roundtrip.Walk(reopened.Doc).Where(node => node.Type == "bulletList"));
        Assert.Equal(2, list.Content!.Count);
    }

    [Fact]
    public void ListaCriadaNoEditorGanhaNumeracaoNoArquivo()
    {
        // Documento sem lista não tem `word/numbering.xml`: a parte é criada.
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
        // Sem contexto de lista, o escritor leria que o parágrafo deixou de ser item.
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
        // A lista colada traz o `numId` do documento de origem, que este não define.
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
        // A sublista numerada dentro de uma com marcador não herda o `numId` de fora.
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
        // A imagem da barra de ferramentas é um bloco e chega sem medida.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var source = "data:image/png;base64," + Convert.ToBase64String(Fixtures.SquarePng());
        model.Doc.Content!.Add(Node.Of("image").With("src", source));

        var (saved, result) = Roundtrip.Save(original, model);

        Assert.Empty(result.Inventory.Lost);
        Assert.Contains("w:drawing", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        Assert.Contains(Roundtrip.PartsOf(saved).Keys, name => name.Contains("image", StringComparison.Ordinal));

        // A medida sai do cabeçalho do PNG, 4 × 4.
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

        // Cabe na coluna de uma A4 com margens de 2,5 cm, com a proporção de 2 para 1.
        Assert.InRange(width, 500, 650);
        Assert.InRange((double)width / height, 1.9, 2.1);
    }

    [Fact]
    public void ImagemComDataUriSemVirgulaNaoDerrubaAGravacao()
    {
        // `data:image/png` sem vírgula: uma imagem não pode derrubar a gravação.
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
        // O id não depende de quantas gravações o processo já fez nem colide com o do
        // arquivo, o que o Word mostra como documento danificado.
        var original = Fixtures.WithInlineImage(1001);

        var first = DrawingIdsOf(original);
        var second = DrawingIdsOf(original);

        Assert.Equal(first, second);
        Assert.Equal(first.Distinct(), first);
        Assert.Contains("1001", first);
        Assert.All(first, id => Assert.True(int.Parse(id) >= 1001));
    }

    /// <summary>Os `wp:docPr/@id` do documento gravado com uma imagem nova.</summary>
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

        // Primeiro na tela: o texto da tabela de dentro chega ao editor.
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
        // A célula do editor pode terminar na tabela aninhada; o `w:tc` do Word tem de terminar em `w:p`.
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
        // Linha inserida desalinha as posições e levaria o `w:vMerge` à linha errada.
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
        // A divisória arrastada no TableKit chega ao `w:tblGrid`.
        var original = Fixtures.WithTable();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var table = model.Doc.Content!.First(block => block.Type == "table");
        foreach (var row in table.Content!)
        {
            row.Content![0].With("colwidth", new JsonArray(200));
            row.Content[1].With("colwidth", new JsonArray(400));
        }

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        // 15 twips por pixel.
        Assert.Contains("<w:gridCol w:w=\"3000\" /><w:gridCol w:w=\"6000\" />", xml, StringComparison.Ordinal);

        // A largura da célula acompanha, senão contradiz a grade e o Word escolhe uma.
        Assert.Contains("w:tcW w:w=\"3000\"", xml, StringComparison.Ordinal);

        // O Word honra a grade como a tela com `table-layout: fixed`.
        Assert.Contains("w:tblLayout w:type=\"fixed\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void EditarTextoDaCelulaNaoMexeNaGradeDeColunas()
    {
        // Twip→pixel arredonda (4675 twips são 311,67 px): a grade não redimensionada fica.
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
        // O esquema exige `w:tblGrid`: colunas iguais na coluna de texto, como o Word ao inserir.
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

        // Linha toda de `tableHeader` é a que se repete no alto de cada página.
        Assert.Contains("w:tblHeader", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void LinhaDeCabecalhoDesligadaSaiDoArquivo()
    {
        // O `w:trPr` traz altura e revisão também: desligar a bandeira não o leva.
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

        // 1,5 pt são 12 oitavos, a unidade do `w:sz`.
        Assert.Contains("<w:top w:val=\"double\" w:color=\"FF0000\" w:sz=\"12\" />", xml, StringComparison.Ordinal);

        // Só `w:nil` remove a borda que a tabela pediu.
        Assert.Contains("w:bottom w:val=\"nil\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void CelulaNaoFormatadaConservaTramaEEstiloExotico()
    {
        // Enquanto ninguém formata a célula, a trama e a borda exótica voltam do XML original.
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
        // O TableKit põe `colwidth` só na coluna arrastada.
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

        // 300 px são 4500 twips; as outras dividem o que sobra da coluna de texto.
        var grid = Regex.Matches(xml, "<w:gridCol w:w=\"(\\d+)\"").Select(match => match.Groups[1].Value).ToList();
        Assert.Equal(3, grid.Count);
        Assert.Equal("4500", grid[0]);
        Assert.Equal(grid[1], grid[2]);
    }

    [Fact]
    public void ArrastarUmaColunaNaoMexeNaMedidaDasOutras()
    {
        // A vizinha que ninguém tocou não passa por `px × 15` (2000 twips são 133,33 px).
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
        // Depois de inserir coluna, o TableKit deixa a célula nova sem largura.
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
        // No `w:trPr` a revisão (`w:ins`, `w:del`, `w:trPrChange`) fecha a sequência.
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
        // Com `w:gridBefore`, a grade do modelo tem de ter as colunas do arquivo.
        var original = Fixtures.WithGridBefore();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var table = model.Doc.Content!.First(block => block.Type == "table");
        var recuada = table.Content![0].Content![0];

        // A célula recuada tem a largura da segunda coluna da grade, 3000 twips.
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
        // Grade que o modelo não descreve inteira: a do arquivo fica, com aviso.
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
        // Pôr a borda de baixo não apaga diagonal, borda interna, `w:space` nem cor de tema.
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
        // A cor de tema vence o `w:color` no Word, e sai com a troca de cor; o `w:space` fica.
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
        // Trocar a aparência da célula apaga a trama e o estilo exótico, com aviso.
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
        // O `rowspan` do TableKit ainda não vira `w:vMerge`: o mínimo é dizer.
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

        // No OOXML a imagem centralizada é um parágrafo centralizado com ela dentro.
        Assert.Contains("w:jc w:val=\"center\"", xml, StringComparison.Ordinal);

        var back = Roundtrip.Open(saved);
        var image = Roundtrip.Walk(back.Doc).First(node => node.Type == "image");
        Assert.Equal("Organograma da diretoria", image.Attrs!["alt"]!.GetValue<string>());
    }

    [Fact]
    public void RedimensionarImagemDoArquivoMantemODesenhoOriginal()
    {
        // Redimensionar muda só o tamanho: o `wp:docPr`, a parte de imagem e o que o
        // escritor não gera (efeito, recorte, borda) ficam.
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

        // 300 × 150 px são 2857500 × 1428750 EMU, no `wp:extent` e no `a:ext`.
        Assert.Equal(2, Regex.Matches(xml, "cx=\"2857500\" cy=\"1428750\"").Count);

        Assert.Equal(
            Regex.Matches(before, "<w:p[ >]").Count,
            Regex.Matches(xml, "<w:p[ >]").Count);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void ImagemAncoradaNoFluxoNaoSeDuplicaAoEditarOParagrafo()
    {
        // A imagem ancorada no fluxo chega como imagem do parágrafo: o `w:r` original não vem junto.
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
        // Os filhos que não são linha ficam no lugar.
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
        // O Word dá por danificado `rgb(255, 0, 0)`, e desenha nome de cor como preto.
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

        // E o documento relido traz as duas marcas.
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
        // `fade` é hexadecimal válido, mas não cor.
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

        // 16 px são 12 pt, e `w:sz` é em meios-pontos.
        Assert.Contains(
            "w:sz w:val=\"24\"",
            Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes),
            StringComparison.Ordinal);
    }

    [Fact]
    public void CaractereDeControleNaoDerrubaAGravacao()
    {
        // `\u0001` não existe no XML 1.0, nem escapado; `\u000B` e `\n` são quebra, e a tabulação, `w:tab`.
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
        // O modelo só nomeia A4 e Carta: o `w:sectPr` só muda quando a página muda.
        var original = Fixtures.WithCustomPaper();
        var (saved, _) = Roundtrip.Save(original, Roundtrip.Clone(Roundtrip.Open(original)));

        var xml = Roundtrip.XmlOf(saved);
        Assert.Contains("w:w=\"8391\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:h=\"11907\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void PapelQueOModeloNaoNomeiaEAvisadoComoInvisivel()
    {
        // O painel mostra A4 para um A5: o aviso diz que a tela aproxima.
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
        // O parágrafo comum continua calado, ligado pelo padrão.
        Assert.Equal(1, xml.Split("w:widowControl").Length - 1);
    }

    [Fact]
    public void KeepsTheTableCellMarginsWhenTheTableIsEdited()
    {
        // `cellMargins` é só leitura: o `w:tblCellMar` do arquivo volta como estava.
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
        // O quadro foi para o começo do parágrafo no editor, e gravado continua um só, ancorado.
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
