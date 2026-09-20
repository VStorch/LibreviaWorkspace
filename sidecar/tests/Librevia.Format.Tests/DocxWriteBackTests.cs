using System.Reflection;
using System.Text.RegularExpressions;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// O que a gravação faz com o bloco que **foi** editado.
/// </summary>
/// <remarks>
/// A edição cirúrgica protege o que ninguém tocou, e é disso que trata
/// <see cref="DocxRoundTripTests"/>. Estes testes olham o outro lado, que é onde
/// a perda de verdade acontece: o bloco reescrito. Cada um deles nasceu de um
/// defeito encontrado no documento real — formatação de parágrafo apagada, lista
/// que volta como parágrafo comum, imagem que desaparece, tabela desmontada, cor
/// que o Word recusa, caractere que derruba a gravação, papel arredondado.
/// </remarks>
public class DocxWriteBackTests
{
    private static Node BlockOf(DocumentModelDto model, int index) => model.Doc.Content![index];

    /// <summary>
    /// Todo fixture nasce **dentro** do esquema OOXML.
    /// </summary>
    /// <remarks>
    /// É o contrapeso de <see cref="Roundtrip.AssertSchema"/>, que confere o
    /// documento gravado: com um fixture inválido, a conferência acusaria o
    /// escritor por um defeito do teste — e um dia alguém desligaria a conferência
    /// para fazer o teste passar. Aqui o dedo aponta para o lugar certo.
    ///
    /// Nove fixtures nasceram fora do esquema, e nenhum deles por capricho do
    /// validador: tabela sem `w:tblGrid`, `w:headerReference` depois do papel,
    /// `w:vMerge` depois do sombreamento, marca de parágrafo antes do
    /// espaçamento — cada uma dessas ordens é obrigatória no OOXML, e o Word
    /// recusa o arquivo que as troca.
    /// </remarks>
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

            // Fixture que não é pacote é imagem — `SquarePng` e companhia.
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

    // --- formatação de parágrafo --------------------------------------------

    [Fact]
    public void EditarParagrafoPreservaEstiloEspacamentoEFundo()
    {
        // O `w:pPr` era montado do zero: corrigir uma vírgula no título trocava
        // `Ttulo1` por `Heading1` — que o documento não define —, apagava o fundo
        // vermelho, o espaçamento, a entrelinha e a borda de baixo.
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

        // E o que o editor nem conhece atravessa a edição porque ninguém o
        // reescreve: a borda e a parada de tabulação.
        Assert.Contains("w:pBdr", xml, StringComparison.Ordinal);
        Assert.Contains("w:pos=\"4500\"", xml, StringComparison.Ordinal);

        Assert.Equal("Título corrigido.Parágrafo comum.", Roundtrip.TextOf(Roundtrip.Open(saved)));
        Assert.Equal(1, result.RewrittenBlocks);
    }

    [Fact]
    public void ParagrafoQueDeixouDeSerTituloPerdeOEstiloDeTitulo()
    {
        // O contrapeso do teste acima: preservar o `styleId` não pode significar
        // que um título continue com cara de título depois de virar parágrafo.
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

    // --- recuo ---------------------------------------------------------------

    [Fact]
    public void DiminuirORecuoAteZeroChegaAoArquivo()
    {
        // O escritor saía calado quando o modelo dizia zero, e o `w:ind` do
        // arquivo ficava: quem apertava Ctrl+[ até o fim via o recuo voltar ao
        // reabrir o documento.
        var original = Fixtures.WithFormattedParagraph();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var block = BlockOf(model, 0);
        block.Attrs!["indentMm"] = null;
        block.Attrs["firstLineMm"] = null;

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("w:left=\"0\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("w:firstLine=", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void RecuoNegativoDoArquivoSobreviveAEdicao()
    {
        // O contrapeso: o leitor só emite recuo positivo, então o modelo diz zero
        // sobre um recuo negativo que ninguém tocou. Zerá-lo traria a linha de
        // volta para dentro da margem sem que ninguém pedisse.
        var original = Fixtures.WithNegativeIndent();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "para fora", "corrigido"));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Contains("w:left=\"-284\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void RecuoAusenteNaoViraAtributoVazio()
    {
        // `indentation.Right = right > 0 ? Invariant(right) : null` tem tipo
        // `string`, e o `null` de um `string` vira `StringValue(null)`: o SDK
        // gravava `w:right=""` e `w:hanging=""` em vez de omitir os atributos.
        // O LibreOffice engole; o Word recusa o documento. Editar um parágrafo
        // que tivesse `w:ind` bastava para produzir o arquivo inválido.
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
        // O marcador é o destino da referência cruzada, da entrada de índice e do
        // link interno. Reescrevendo o parágrafo só a partir do modelo — que não
        // o representa —, ele desaparecia do arquivo, e o inventário nada dizia:
        // quem citava o marcador passava a apontar para o vazio.
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
        // `oid` repetido é bloco colado, e o XML original é de **um** deles. A
        // partir da segunda ocorrência o bloco já era gerado do zero; o que ainda
        // vinha do original era o que o escritor copia de lá — objeto ancorado e,
        // agora, marcador. Copiado duas vezes, o documento fica com dois
        // marcadores de mesmo id, que é âncora ambígua para quem os cita.
        var original = Fixtures.WithBookmarkAroundParagraph();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        model.Doc.Content!.Insert(1, BlockOf(model, 0));

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, Roundtrip.Clone(model)).Bytes);

        Assert.Equal(1, Regex.Matches(xml, "<w:bookmarkStart").Count);
        Assert.Equal(1, Regex.Matches(xml, "<w:bookmarkEnd").Count);
    }

    // --- listas -------------------------------------------------------------

    [Fact]
    public void ItemDeListaEditadoContinuaApontandoAMesmaNumeracao()
    {
        // O leitor não punha o `numId` no nó, e o escritor gravava
        // `w:numId w:val="0"` — que no formato quer dizer "sem numeração". A
        // pessoa corrigia uma palavra e o item virava parágrafo comum, sem marca
        // e sem recuo.
        var original = Fixtures.WithBulletList();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "Primeiro item", "Item corrigido"));

        var (saved, _) = Roundtrip.Save(original, model);

        Assert.Contains("<w:numId w:val=\"1\"", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        Assert.DoesNotContain("<w:numId w:val=\"0\"", Roundtrip.XmlOf(saved), StringComparison.Ordinal);

        // E continua sendo lista ao reabrir, que é o que o usuário vê.
        var reopened = Roundtrip.Open(saved);
        var list = Assert.Single(Roundtrip.Walk(reopened.Doc).Where(node => node.Type == "bulletList"));
        Assert.Equal(2, list.Content!.Count);
    }

    [Fact]
    public void ListaCriadaNoEditorGanhaNumeracaoNoArquivo()
    {
        // Documento que nunca teve lista não tem `word/numbering.xml`. A parte é
        // criada, com uma definição de marcador — senão a lista nova volta como
        // parágrafo comum, e isso acontecia sem aviso nenhum.
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
        // O leitor só junta parágrafos numerados numa lista no laço do corpo, e
        // a tabela chamava o escritor sempre sem contexto de lista — o que ele
        // lia como "este parágrafo deixou de ser item" e apagava o `w:numPr`.
        // Corrigir uma palavra na célula tirava os marcadores, sem nada no
        // inventário.
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
        // O `numId` viaja no modelo, e uma lista colada de outro documento traz o
        // do documento de origem. Gravá-lo cru aponta para uma definição que este
        // arquivo não tem: a lista perde o marcador, que é o mesmo sintoma que o
        // `numId` veio curar.
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
        // A sublista herdava o `numId` da lista de fora sem olhar o tipo: uma
        // lista numerada dentro de uma com marcador saía com marcador.
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

    // --- imagens ------------------------------------------------------------

    [Fact]
    public void ImagemInseridaPelaBarraSobreviveAoSalvar()
    {
        // A imagem da barra de ferramentas é um **bloco** e chega sem medida. O
        // escritor não tratava o tipo, então ela virava parágrafo vazio: a pessoa
        // inseria a captura de tela, salvava e o documento voltava sem ela.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var source = "data:image/png;base64," + Convert.ToBase64String(Fixtures.SquarePng());
        model.Doc.Content!.Add(Node.Of("image").With("src", source));

        var (saved, result) = Roundtrip.Save(original, model);

        Assert.Empty(result.Inventory.Lost);
        Assert.Contains("w:drawing", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        Assert.Contains(Roundtrip.PartsOf(saved).Keys, name => name.Contains("image", StringComparison.Ordinal));

        // A medida sai do cabeçalho do PNG — 4 × 4 —, e não do chute de 600 × 450
        // que deformava toda imagem sem tamanho declarado.
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

        // Cabe na coluna de uma A4 com margens de 2,5 cm, e continua com o dobro
        // de largura que de altura.
        Assert.InRange(width, 500, 650);
        Assert.InRange((double)width / height, 1.9, 2.1);
    }

    [Fact]
    public void ImagemComDataUriSemVirgulaNaoDerrubaAGravacao()
    {
        // `data:image/png` sem vírgula fazia o recorte do cabeçalho estourar o fim
        // da string, e a exceção subia até o usuário como falha ao salvar: o
        // documento inteiro por causa de uma imagem.
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
        // O contador era `static`: o servidor atende várias requisições no mesmo
        // processo, e o id da imagem dependia de quantas gravações já tinham
        // acontecido — quando não colidia com o desenho que o arquivo já traz,
        // que é o que o Word mostra como documento danificado.
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

    // --- tabelas ------------------------------------------------------------

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

        // Primeiro na tela: o texto da tabela de dentro precisa chegar ao editor.
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
        // O `tableCell` do editor aceita qualquer bloco, e desde que a tabela
        // aninhada é lida a célula pode terminar nela — basta apagar o parágrafo
        // de depois. Um `w:tc` que não termina em `w:p` é inválido para o Word.
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
        // Só a **remoção** entrava no inventário. Uma linha inserida desalinha as
        // posições e joga o `w:trPr` e o `w:tcPr` de uma linha na de baixo: o
        // `w:vMerge` fora de lugar é o caso em que o Word acusa tabela corrompida.
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

        // E a tabela continua abrindo com as três linhas que a pessoa vê.
        Assert.Equal(3, Roundtrip.Open(saved).Doc.Content![0].Content!.Count);
    }

    [Fact]
    public void MarcadorEntreLinhasContinuaEntreAsLinhas()
    {
        // Todos os filhos que não são linha iam para antes da primeira: um
        // `w:bookmarkStart` que abraçava a segunda linha virava marcador vazio no
        // alto da tabela.
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

    // --- cores e texto ------------------------------------------------------

    [Fact]
    public void CorDoCssViraHexadecimalDeSeisDigitos()
    {
        // `w:color w:val="rgb(255, 0, 0)"` faz o Word declarar o documento
        // danificado, e um nome de cor é desenhado como preto. O editor grava cor
        // como o CSS a escreve, e é aqui que ela vira o que o formato aceita.
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
        // O parágrafo reescrito é o único lugar onde a perda acontece de verdade.
        // Antes destas duas linhas no escritor, o expoente sobrevivia à leitura e
        // morria na gravação: o trecho voltava para a linha do texto, e o aviso
        // saía como "formatação superscript" no inventário.
        var original = Fixtures.WithVerticalAlignment();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        Assert.True(Roundtrip.EditFirstTextContaining(model, "O e m", "O e n"));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Contains("w:vertAlign w:val=\"superscript\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:vertAlign w:val=\"subscript\"", xml, StringComparison.Ordinal);
        Assert.Empty(result.Inventory.Lost);

        // E o documento relido traz as duas marcas de volta: a ida e a volta
        // fecham, que é o que o editor vai ver na próxima abertura.
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
        // `fade` não é cor nomeada, mas é hexadecimal válido: a conversão caía no
        // último caso e devolvia FFAADD — uma cor errada no lugar de um aviso.
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
        // `\u0001` não existe no XML 1.0, nem escapado: o serializador levantava
        // exceção e a gravação inteira falhava. `\u000B` e `\n` são quebra de
        // linha, e a tabulação é `w:tab`.
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

        // O texto continua legível, sem o caractere que ninguém vê.
        Assert.Equal("Primeiro parágrafo.abc\tdTerceiro parágrafo.", Roundtrip.TextOf(Roundtrip.Open(saved)));
    }

    // --- página -------------------------------------------------------------

    [Fact]
    public void PapelForaDeA4NaoEArredondadoAoSalvar()
    {
        // O modelo só nomeia A4 e Carta, e a gravação regravava `w:pgSz` em todo
        // save: um documento em A5 virava A4 por ter recebido uma correção de
        // vírgula. Agora o `w:sectPr` só é tocado quando a página mudou.
        var original = Fixtures.WithCustomPaper();
        var (saved, _) = Roundtrip.Save(original, Roundtrip.Clone(Roundtrip.Open(original)));

        var xml = Roundtrip.XmlOf(saved);
        Assert.Contains("w:w=\"8391\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:h=\"11907\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void PapelQueOModeloNaoNomeiaEAvisadoComoInvisivel()
    {
        // O painel mostra A4 para um documento em A5. O arquivo mantém o papel,
        // mas o usuário precisa saber que o que ele lê na tela é aproximação.
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

        // Continua A5: o lado curto virou a altura, e não 11906 de A4.
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
}
