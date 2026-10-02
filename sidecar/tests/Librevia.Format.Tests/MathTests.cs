using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Librevia.Format.Docx;
using static Librevia.Format.Tests.Roundtrip;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;

namespace Librevia.Format.Tests;

/// <summary>
/// Equações (M11, fase 1): lidas como nó `math`, desenhadas em MathML, devolvidas
/// ao arquivo com o OMML como veio.
/// </summary>
/// <remarks>
/// Antes desta fase o `m:oMath` caía no `default:` do leitor e sumia sem aviso ao
/// editar o parágrafo. O OMML dos fixtures é escrito como o Word e o LibreOffice o
/// gravam — o corpus não tem equação nenhuma.
/// </remarks>
public class MathTests
{
    internal const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    internal const string Mns = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>O `m:r` do Word: a fonte de matemática no `w:rPr`, o estilo no `m:rPr`.</summary>
    internal static string R(string text, string? sty = null) =>
        "<m:r>" + (sty is null ? string.Empty : $"""<m:rPr><m:sty m:val="{sty}"/></m:rPr>""") +
        """<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>""" +
        $"<m:t>{text}</m:t></m:r>";

    internal const string Ctrl = """<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr>""";

    /// <summary>πr² no meio da frase, como o Word grava a equação em linha.</summary>
    internal static readonly string Inline =
        """<w:p><w:r><w:t xml:space="preserve">A área é </w:t></w:r>""" +
        $"<m:oMath>{R("π")}<m:sSup><m:sSupPr>{Ctrl}</m:sSupPr><m:e>{R("r")}</m:e><m:sup>{R("2")}</m:sup></m:sSup></m:oMath>" +
        """<w:r><w:t xml:space="preserve"> e acabou.</w:t></w:r></w:p>""";

    /// <summary>A fórmula de Bhaskara em exibição, centrada, como o Word a grava.</summary>
    internal static readonly string Display =
        """<w:p><m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath>""" +
        R("x") + R("=") +
        $"<m:f><m:fPr>{Ctrl}</m:fPr><m:num>{R("-b±")}<m:rad><m:radPr><m:degHide m:val=\"1\"/>{Ctrl}</m:radPr><m:deg/>" +
        $"<m:e>{R("Δ")}</m:e></m:rad></m:num><m:den>{R("2a")}</m:den></m:f>" +
        "</m:oMath></m:oMathPara></w:p>";

    /// <summary>
    /// Como o LibreOffice grava: `m:sty` sempre explícito, o somatório com os
    /// liga-desliga declarados, o delimitador com o separador vazio.
    /// </summary>
    internal static readonly string LibreOffice =
        """<w:p><w:r><w:t xml:space="preserve">Soma </w:t></w:r><m:oMath>""" +
        """<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/><m:subHide m:val="0"/><m:supHide m:val="0"/></m:naryPr>""" +
        $"<m:sub>{R("i", "i")}{R("=", "p")}{R("1", "p")}</m:sub><m:sup>{R("n", "i")}</m:sup>" +
        $"<m:e><m:sSub><m:e>{R("x", "i")}</m:e><m:sub>{R("i", "i")}</m:sub></m:sSub></m:e></m:nary>" +
        """<m:d><m:dPr><m:begChr m:val="["/><m:sepChr m:val=""/><m:endChr m:val="]"/></m:dPr>""" +
        $"<m:e>{R("y", "i")}</m:e></m:d>" +
        "</m:oMath></w:p>";

    /// <summary>A caixa sem o lado de cima: o CSS de um `mrow` não a desenha.</summary>
    internal static readonly string Lossy =
        """<w:p><w:r><w:t xml:space="preserve">Caixa </w:t></w:r><m:oMath>""" +
        $"""<m:borderBox><m:borderBoxPr><m:hideTop m:val="1"/></m:borderBoxPr><m:e>{R("z")}</m:e></m:borderBox>""" +
        "</m:oMath></w:p>";

    /// <summary>Duas linhas numa equação de exibição (Shift+Enter no Word).</summary>
    internal static readonly string TwoLines =
        """<w:p><m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>""" +
        $"<m:oMath>{R("a")}{R("=")}{R("1")}</m:oMath><m:oMath>{R("b")}{R("=")}{R("2")}</m:oMath>" +
        "</m:oMathPara></w:p>";

    private const string Plain = """<w:p><w:r><w:t>Intocado.</w:t></w:r></w:p>""";

    internal static byte[] WithMath(string? body = null)
    {
        var ns = $"xmlns:w=\"{W}\" xmlns:m=\"{Mns}\"";
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var part = document.AddMainDocumentPart();
            part.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                $"<w:document {ns}><w:body>{body ?? Plain + Inline + Display + LibreOffice + Lossy + TwoLines}" +
                "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>" +
                "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>");
            part.Document.Save();
        }

        return buffer.ToArray();
    }

    internal static List<Node> Equations(DocumentModelDto model) =>
        Walk(model.Doc).Where(node => node.Type == "math").ToList();

    private static string? Attr(Node node, string name) =>
        node.Attrs?.GetValueOrDefault(name)?.GetValue<string>();

    private static bool Flag(Node node, string name) =>
        node.Attrs?.GetValueOrDefault(name)?.GetValue<bool>() == true;

    internal static List<string> OmmlOf(byte[] docx)
    {
        using var stream = new MemoryStream(docx);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.Document!.Descendants<OfficeMath>().Select(math => math.OuterXml).ToList();
    }

    private static string Convert(string inner) =>
        OmmlMath.Convert($"<m:oMath xmlns:m=\"{Mns}\" xmlns:w=\"{W}\">{inner}</m:oMath>").MathMl;

    // --- a conversão, construção a construção --------------------------------

    [Theory]
    [InlineData("<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>",
        "<mfrac><mrow><mi>a</mi></mrow><mrow><mi>b</mi></mrow></mfrac>")]
    [InlineData("""<m:f><m:fPr><m:type m:val="lin"/></m:fPr><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>""",
        "<mrow><mrow><mi>a</mi></mrow><mo>/</mo><mrow><mi>b</mi></mrow></mrow>")]
    [InlineData("""<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num/><m:den/></m:f>""",
        """<mfrac linethickness="0">""")]
    [InlineData("""<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>""",
        "<msqrt><mrow><mi>x</mi></mrow></msqrt>")]
    [InlineData("<m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>",
        "<mroot><mrow><mi>x</mi></mrow><mrow><mn>3</mn></mrow></mroot>")]
    [InlineData("<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>",
        "<msub><mrow><mi>x</mi></mrow><mrow><mi>i</mi></mrow></msub>")]
    [InlineData("<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>",
        "<msup><mrow><mi>x</mi></mrow><mrow><mn>2</mn></mrow></msup>")]
    [InlineData("<m:sSubSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>",
        "<msubsup><mrow><mi>x</mi></mrow><mrow><mn>1</mn></mrow><mrow><mn>2</mn></mrow></msubsup>")]
    [InlineData("<m:sPre><m:sub><m:r><m:t>1</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup><m:e><m:r><m:t>C</m:t></m:r></m:e></m:sPre>",
        "<mmultiscripts><mrow><mi>C</mi></mrow><mprescripts /><mrow><mn>1</mn></mrow><mrow><mn>2</mn></mrow></mmultiscripts>")]
    [InlineData("<m:nary><m:sub><m:r><m:t>0</m:t></m:r></m:sub><m:sup><m:r><m:t>1</m:t></m:r></m:sup><m:e><m:r><m:t>f</m:t></m:r></m:e></m:nary>",
        """<msubsup><mo largeop="true" movablelimits="false">∫</mo>""")]
    [InlineData("""<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub/><m:sup/><m:e/></m:nary>""",
        """<munderover><mo largeop="true">∑</mo>""")]
    [InlineData("""<m:nary><m:naryPr><m:chr m:val="∏"/><m:supHide m:val="1"/></m:naryPr><m:sub/><m:sup/><m:e/></m:nary>""",
        """<munder><mo largeop="true">∏</mo>""")]
    [InlineData("<m:d><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:d>",
        """<mrow><mo fence="true" stretchy="true">(</mo><mrow><mi>a</mi></mrow><mo separator="true">|</mo><mrow><mi>b</mi></mrow><mo fence="true" stretchy="true">)</mo></mrow>""")]
    [InlineData("""<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val=""/></m:dPr><m:e/></m:d>""",
        """<mrow><mo fence="true" stretchy="true">{</mo><mrow /></mrow>""")]
    [InlineData("<m:m><m:mr><m:e><m:r><m:t>1</m:t></m:r></m:e><m:e><m:r><m:t>0</m:t></m:r></m:e></m:mr><m:mr><m:e/><m:e/></m:mr></m:m>",
        "<mtable><mtr><mtd><mrow><mn>1</mn></mrow></mtd><mtd><mrow><mn>0</mn></mrow></mtd></mtr><mtr><mtd><mrow /></mtd><mtd><mrow /></mtd></mtr></mtable>")]
    [InlineData("<m:eqArr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:eqArr>",
        "<mtable><mtr><mtd><mrow><mi>a</mi></mrow></mtd></mtr><mtr><mtd><mrow><mi>b</mi></mrow></mtd></mtr></mtable>")]
    [InlineData("""<m:acc><m:accPr><m:chr m:val="̇"/></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>""",
        """<mover accent="true"><mrow><mi>x</mi></mrow><mo stretchy="false">˙</mo></mover>""")]
    [InlineData("<m:acc><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>",
        """<mo stretchy="false">^</mo>""")]
    [InlineData("""<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:bar>""",
        """<mover accent="true"><mrow><mi>x</mi></mrow><mo stretchy="true">‾</mo></mover>""")]
    [InlineData("<m:bar><m:e><m:r><m:t>x</m:t></m:r></m:e></m:bar>",
        """<munder accentunder="true">""")]
    [InlineData("<m:groupChr><m:e><m:r><m:t>abc</m:t></m:r></m:e></m:groupChr>",
        """<munder><mrow><mi>a</mi><mi>b</mi><mi>c</mi></mrow><mo stretchy="true">⏟</mo></munder>""")]
    [InlineData("<m:limLow><m:e><m:r><m:t>lim</m:t></m:r></m:e><m:lim><m:r><m:t>n→∞</m:t></m:r></m:lim></m:limLow>",
        "<munder><mrow><mi>l</mi><mi>i</mi><mi>m</mi></mrow><mrow><mi>n</mi><mo>→</mo><mi>∞</mi></mrow></munder>")]
    [InlineData("<m:limUpp><m:e><m:r><m:t>x</m:t></m:r></m:e><m:lim><m:r><m:t>y</m:t></m:r></m:lim></m:limUpp>",
        "<mover><mrow><mi>x</mi></mrow><mrow><mi>y</mi></mrow></mover>")]
    [InlineData("""<m:func><m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func>""",
        """<mrow><mrow><mi mathvariant="normal">sin</mi></mrow><mo>⁡</mo><mrow><mi>x</mi></mrow></mrow>""")]
    [InlineData("<m:box><m:e><m:r><m:t>x</m:t></m:r></m:e></m:box>",
        "<mrow><mrow><mi>x</mi></mrow></mrow>")]
    [InlineData("<m:borderBox><m:e><m:r><m:t>x</m:t></m:r></m:e></m:borderBox>",
        """<mrow class="omml-caixa"><mrow><mi>x</mi></mrow></mrow>""")]
    [InlineData("""<m:phant><m:phantPr><m:show m:val="0"/></m:phantPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:phant>""",
        "<mphantom><mrow><mi>x</mi></mrow></mphantom>")]
    [InlineData("""<m:phant><m:phantPr><m:zeroWid/></m:phantPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:phant>""",
        """<mpadded width="0"><mrow><mi>x</mi></mrow></mpadded>""")]
    [InlineData("""<m:r><m:rPr><m:nor/></m:rPr><m:t>se e só se</m:t></m:r>""",
        "<mtext>se e só se</mtext>")]
    [InlineData("""<m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>v1</m:t></m:r>""",
        "<mi>𝐯</mi><mn>𝟏</mn>")]
    [InlineData("""<m:r><m:rPr><m:scr m:val="double-struck"/><m:sty m:val="p"/></m:rPr><m:t>R</m:t></m:r>""",
        "<mi>ℝ</mi>")]
    [InlineData("""<m:r><m:rPr><m:scr m:val="script"/></m:rPr><m:t>L</m:t></m:r>""",
        "<mi>ℒ</mi>")]
    [InlineData("""<m:r><m:rPr><m:scr m:val="fraktur"/></m:rPr><m:t>g</m:t></m:r>""",
        "<mi>𝔤</mi>")]
    [InlineData("<m:r><m:t>3,14+x</m:t></m:r>",
        "<mn>3,14</mn><mo>+</mo><mi>x</mi>")]
    public void ConverteCadaConstrucao(string omml, string expected)
    {
        var mathml = Convert(omml);
        Assert.Contains(expected, mathml);
        Assert.StartsWith("""<math display="inline" xmlns="http://www.w3.org/1998/Math/MathML">""", mathml);
    }

    [Theory]
    [InlineData("""<m:borderBox><m:borderBoxPr><m:hideTop m:val="1"/></m:borderBoxPr><m:e/></m:borderBox>""", "m:borderBox")]
    [InlineData("""<m:r><m:t>x</m:t></m:r><w:sdt/>""", "w:sdt")]
    public void ConstrucaoQueATelaNaoDesenhaVaiParaALista(string omml, string label)
    {
        var converted = OmmlMath.Convert($"<m:oMath xmlns:m=\"{Mns}\" xmlns:w=\"{W}\">{omml}</m:oMath>");
        Assert.Equal([label], converted.Lossy);
    }

    [Fact]
    public void TudoOQueSeDesenhaNaoVaiParaALista()
    {
        var model = Open(WithMath());
        var drawn = Equations(model).Where(node => !(Attr(node, "omml") ?? "").Contains("borderBox", StringComparison.Ordinal));
        Assert.All(drawn, node => Assert.True(Flag(node, "editable")));
    }

    // --- a leitura -----------------------------------------------------------

    [Fact]
    public void LeAEquacaoEmLinhaEADeExibicao()
    {
        var result = DocxReader.Read(WithMath());
        var equations = Equations(result.Model);
        Assert.Equal(5, equations.Count);

        var inline = equations[0];
        Assert.False(Flag(inline, "display"));
        Assert.StartsWith("<m:oMath", Attr(inline, "omml"));
        Assert.Contains("<msup><mrow><mi>r</mi></mrow><mrow><mn>2</mn></mrow></msup>", Attr(inline, "mathml"));
        Assert.Equal(string.Empty, Attr(inline, "latex"));

        var display = equations[1];
        Assert.True(Flag(display, "display"));
        Assert.Equal("center", Attr(display, "jc"));
        Assert.StartsWith("<m:oMathPara", Attr(display, "omml"));
        Assert.Contains("<msqrt><mrow><mi>Δ</mi></mrow></msqrt>", Attr(display, "mathml"));

        // A de duas linhas é um nó só, com uma linha por `m:oMath`.
        var lines = equations[4];
        Assert.Contains("""<mtable columnalign="left"><mtr>""", Attr(lines, "mathml"));
        Assert.Equal(2, Attr(lines, "mathml")!.Split("<mtr>").Length - 1);

        // O texto em volta continua sendo texto.
        Assert.Contains("A área é ", TextOf(result.Model));
        Assert.Contains(" e acabou.", TextOf(result.Model));
    }

    [Fact]
    public void ConstrucaoNaoDesenhadaTravaAEquacaoEAvisaSemSerEstrutural()
    {
        var result = DocxReader.Read(WithMath());
        var locked = Equations(result.Model).Single(node => !Flag(node, "editable"));
        Assert.Equal(["m:borderBox"], locked.Attrs!["lossy"]!.AsArray().Select(item => item!.GetValue<string>()));
        Assert.Contains(Inventory.Equations, result.Inventory.Invisible);
        Assert.DoesNotContain(Inventory.Equations, result.Inventory.Structural);
    }

    [Fact]
    public void DocumentoSoComEquacoesDesenhadasNaoAvisaNada()
    {
        var result = DocxReader.Read(WithMath(Plain + Inline + Display));
        Assert.DoesNotContain(Inventory.Equations, result.Inventory.Invisible);
    }

    // --- a gravação ----------------------------------------------------------

    [Fact]
    public void AbrirESalvarDevolveCadaEquacaoComoVeio()
    {
        var original = WithMath();
        var (bytes, result) = Save(original, Clone(Open(original)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Equal(OmmlOf(original), OmmlOf(bytes));
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void EditarOTextoAoLadoMantemOOmmlIntacto()
    {
        var original = WithMath();
        var model = Clone(Open(original));
        Assert.True(EditFirstTextContaining(model, "A área é ", "A área do círculo é "));
        Assert.True(EditFirstTextContaining(model, "Soma ", "Somatório "));

        var (bytes, result) = Save(original, model);

        Assert.Equal(2, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);
        Assert.Equal(OmmlOf(original), OmmlOf(bytes));

        var xml = XmlOf(bytes);
        Assert.Contains("A área do círculo é ", xml);
        // Escrita como entrou: sem o `xmlns:w` que o OuterXml declara em cada `w:rPr`.
        Assert.DoesNotContain("<w:rPr xmlns:w=", xml);
        // A equação volta no mesmo lugar: entre os dois trechos de texto.
        var reread = Open(bytes);
        var paragraph = reread.Doc.Content!.Single(block => Walk(block).Any(node => node.Text == "A área do círculo é "));
        Assert.Equal(["text", "math", "text"], paragraph.Content!.Select(node => node.Type));
    }

    [Fact]
    public void OMathMLDerivadoNaoMudaAImpressaoDigital()
    {
        // O MathML, o LaTeX e a lista saem do OMML: uma conversão melhor numa versão
        // futura não pode fazer o parágrafo parecer editado.
        var original = WithMath();
        var model = Clone(Open(original));
        foreach (var equation in Equations(model))
        {
            equation.With("mathml", "<math/>").With("latex", "x").With("lossy", new JsonArray()).With("editable", false);
            equation.Attrs!.Remove("jc");
        }

        Assert.Equal(0, Save(original, model).Result.RewrittenBlocks);
    }

    [Fact]
    public void EquacaoSemOmmlNemMathMlEDeclaradaComoPerda()
    {
        var original = WithMath();
        var model = Clone(Open(original));
        var equation = Equations(model)[0];
        equation.With("omml", null).With("mathml", "");

        var (bytes, result) = Save(original, model);
        Assert.Contains("equação que não pôde ser gravada", result.Inventory.Lost);
        Assert.Equal(OmmlOf(original).Count - 1, OmmlOf(bytes).Count);
    }

    [Fact]
    public void RascunhoDeAntesDasEquacoesDeclaraAPerdaAoEditar()
    {
        var original = WithMath();
        // O modelo como o leitor de antes das equações o dava: sem `math`.
        List<Node> content;
        using (var stream = new MemoryStream(original))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var part = document.MainDocumentPart!;
            (content, _) = new BodyReader(part, new Inventory(), math: false).Read(part.Document!.Body!);
        }

        var doc = Node.Of("doc");
        doc.Content = content;
        var model = Clone(new DocumentModelDto(Open(original).Page, doc));
        Assert.DoesNotContain(Walk(model.Doc), node => node.Type == "math");

        var untouched = Save(original, model with { BeforeMath = true });
        Assert.Equal(0, untouched.Result.RewrittenBlocks);
        Assert.Equal(OmmlOf(original), OmmlOf(untouched.Bytes));

        Assert.True(EditFirstTextContaining(model, "A área é ", "A área editada é "));
        var (_, result) = Save(original, model with { BeforeMath = true });
        Assert.Contains("equação num parágrafo que você editou", result.Inventory.Lost);
    }
}
