using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Librevia.Format.Docx;
using static Librevia.Format.Tests.MathTests;
using static Librevia.Format.Tests.Roundtrip;

namespace Librevia.Format.Tests;

/// <summary>
/// Equações: a nova ou editada chega sem OMML, só com o MathML, e o OMML sai
/// dele (OmmlMath.ToOmml).
/// </summary>
public class MathEditTests
{
    private static readonly XNamespace M = OmmlMath.M;

    /// <summary>
    /// A forma de um OMML, para comparar estrutura: o nome de cada construção com
    /// os argumentos entre parênteses, o texto dos runs vizinhos fundido (o Word e
    /// este escritor partem os runs em lugares diferentes) e as propriedades que
    /// mudam o desenho — o caractere do n-ário, do acento e dos delimitadores.
    /// </summary>
    private static string Shape(XElement element)
    {
        var builder = new StringBuilder();
        var pendingText = new StringBuilder();

        void Flush()
        {
            if (pendingText.Length == 0) return;
            builder.Append('"').Append(pendingText).Append('"');
            pendingText.Clear();
        }

        foreach (var child in element.Elements())
        {
            if (child.Name.Namespace != M || child.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal)) continue;
            if (child.Name.LocalName == "r")
            {
                pendingText.Append(string.Concat(child.Elements(M + "t").Select(t => t.Value)));
                continue;
            }

            Flush();
            builder.Append(child.Name.LocalName).Append(Detail(child)).Append('(').Append(Shape(child)).Append(')');
        }

        Flush();
        return builder.ToString();
    }

    private static string Detail(XElement element)
    {
        var properties = element.Element(M + element.Name.LocalName + "Pr");
        string? Val(string name) => properties?.Element(M + name)?.Attribute(M + "val")?.Value;
        return element.Name.LocalName switch
        {
            "nary" => $"[{Val("chr") ?? "∫"}]",
            "acc" => $"[{Val("chr") ?? "̂"}]",
            "d" => $"[{Val("begChr") ?? "("}{Val("endChr") ?? ")"}]",
            "bar" => $"[{Val("pos") ?? "bot"}]",
            "groupChr" => $"[{Val("chr") ?? "⏟"}{Val("pos") ?? "bot"}]",
            _ => string.Empty,
        };
    }

    private static string ShapeOf(string mathMl, bool display = false) =>
        Shape(XElement.Parse(OmmlMath.ToOmml(mathMl, display, null)!));

    private const string Ns = "xmlns=\"http://www.w3.org/1998/Math/MathML\"";

    // --- construção a construção --------------------------------------------

    [Theory]
    // O Temml: fração e raiz (com a escora `mspace` dentro).
    [InlineData("<mfrac><mi>a</mi><mi>b</mi></mfrac><mo>+</mo><msqrt><mrow><mi>x</mi><mspace width=\"0pt\" height=\"0.5em\"/></mrow></msqrt>",
        "f(num(\"a\")den(\"b\"))\"+\"rad(deg()e(\"x\"))")]
    [InlineData("<mroot><mrow><mi>x</mi></mrow><mn>3</mn></mroot>", "rad(deg(\"3\")e(\"x\"))")]
    [InlineData("<mfrac linethickness=\"0px\"><mi>n</mi><mi>k</mi></mfrac>", "f(num(\"n\")den(\"k\"))")]
    [InlineData("<msub><mi>x</mi><mi>i</mi></msub>", "sSub(e(\"x\")sub(\"i\"))")]
    [InlineData("<msup><mi>x</mi><mn>2</mn></msup>", "sSup(e(\"x\")sup(\"2\"))")]
    [InlineData("<msubsup><mi>x</mi><mi>i</mi><mn>2</mn></msubsup>", "sSubSup(e(\"x\")sub(\"i\")sup(\"2\"))")]
    // O pré-índice: o nosso `mmultiscripts` e o `{}_a^b X` do Temml.
    [InlineData("<mmultiscripts><mi>X</mi><mprescripts/><mi>a</mi><mi>b</mi></mmultiscripts>", "sPre(sub(\"a\")sup(\"b\")e(\"X\"))")]
    [InlineData("<mrow><msubsup><mrow></mrow><mi>a</mi><mi>b</mi></msubsup><mi>X</mi></mrow>", "sPre(sub(\"a\")sup(\"b\")e(\"X\"))")]
    // O n-ário do Temml, com o corpo como irmão até o próximo operador.
    [InlineData("<mrow><msubsup><mo movablelimits=\"false\">∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></msubsup><msub><mi>x</mi><mi>i</mi></msub><mo>=</mo><mi>y</mi></mrow>",
        "nary[∑](sub(\"i=1\")sup(\"n\")e(sSub(e(\"x\")sub(\"i\"))))\"=y\"")]
    [InlineData("<mrow><mrow><munderover><mo movablelimits=\"false\">∏</mo><mi>k</mi><mi>m</mi></munderover></mrow><mi>a</mi></mrow>",
        "nary[∏](sub(\"k\")sup(\"m\")e(\"a\"))")]
    [InlineData("<mrow><msubsup><mo movablelimits=\"false\">∫</mo><mn>0</mn><mn>1</mn></msubsup><mi>f</mi><mspace width=\"0.1667em\"/><mi>d</mi><mi>x</mi></mrow>",
        "nary[∫](sub(\"0\")sup(\"1\")e(\"fdx\"))")]
    // Os delimitadores, com separador.
    [InlineData("<mrow><mo fence=\"true\" stretchy=\"true\">(</mo><mi>x</mi><mo stretchy=\"true\" form=\"infix\">|</mo><mi>y</mi><mo fence=\"true\" stretchy=\"true\">)</mo></mrow>",
        "d[()](e(\"x\")e(\"y\"))")]
    [InlineData("<mrow><mo fence=\"true\">{</mo><mi>x</mi><mo fence=\"true\"></mo></mrow>", "d[{](e(\"x\"))")]
    [InlineData("<mrow><mo fence=\"true\">|</mo><mi>x</mi><mo fence=\"true\">|</mo></mrow>", "d[||](e(\"x\"))")]
    // A matriz e as linhas.
    [InlineData("<mrow><mo fence=\"true\">(</mo><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable><mo fence=\"true\">)</mo></mrow>",
        "d[()](e(m(mr(e(\"a\")e(\"b\"))mr(e(\"c\")e(\"d\")))))")]
    [InlineData("<mtable><mtr><mtd><mi>a</mi></mtd></mtr><mtr><mtd><mi>b</mi></mtd></mtr></mtable>", "eqArr(e(\"a\")e(\"b\"))")]
    // Acentos, barra, chave por cima e por baixo, limites.
    [InlineData("<mover><mi>x</mi><mo stretchy=\"false\">ˆ</mo></mover>", "acc[̂](e(\"x\"))")]
    [InlineData("<mover><mi>v</mi><mo stretchy=\"false\">→</mo></mover>", "acc[⃗](e(\"v\"))")]
    [InlineData("<mover><mi>a</mi><mo stretchy=\"false\">˙</mo></mover>", "acc[̇](e(\"a\"))")]
    [InlineData("<mover><mi>y</mi><mo stretchy=\"false\">‾</mo></mover>", "acc[̅](e(\"y\"))")]
    [InlineData("<mover accent=\"true\"><mi>y</mi><mo stretchy=\"true\">‾</mo></mover>", "bar[top](e(\"y\"))")]
    [InlineData("<munder accentunder=\"true\"><mi>y</mi><mo stretchy=\"true\">‾</mo></munder>", "bar[bot](e(\"y\"))")]
    [InlineData("<munder><munder><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mo stretchy=\"true\">⏟</mo></munder><mi>n</mi></munder>",
        "limLow(e(groupChr[⏟bot](e(\"a+b\")))lim(\"n\"))")]
    [InlineData("<mover><mrow><mi>a</mi></mrow><mo stretchy=\"true\">⏞</mo></mover>", "groupChr[⏞top](e(\"a\"))")]
    [InlineData("<munder><mi>lim</mi><mrow><mi>n</mi><mo>→</mo><mi>∞</mi></mrow></munder>", "limLow(e(\"lim\")lim(\"n→∞\"))")]
    [InlineData("<mover><mi>b</mi><mi>a</mi></mover>", "limUpp(e(\"b\")lim(\"a\"))")]
    // A função: a do Temml (o nome com o U+2061 num `mrow`) e a nossa.
    [InlineData("<mrow><mrow><mi>sin</mi><mo>⁡</mo><mspace width=\"0.1667em\"/></mrow><mi>x</mi></mrow>", "func(fName(\"sin\")e(\"x\"))")]
    [InlineData("<mrow><mrow><mi mathvariant=\"normal\">log</mi></mrow><mo>⁡</mo><mrow><mi>z</mi></mrow></mrow>", "func(fName(\"log\")e(\"z\"))")]
    [InlineData("<mrow><munder><mi>lim</mi><mi>x</mi></munder><mo>⁡</mo><mspace width=\"0.1667em\"/><mi>f</mi></mrow>", "func(fName(limLow(e(\"lim\")lim(\"x\")))e(\"f\"))")]
    // A caixa, e o texto.
    [InlineData("<mrow class=\"omml-caixa\"><mi>z</mi></mrow>", "borderBox(e(\"z\"))")]
    [InlineData("<mi>a</mi><mtext> se </mtext><mi>b</mi>", "\"a se b\"")]
    public void ConverteCadaConstrucao(string inner, string expected)
    {
        Assert.Equal(expected, ShapeOf($"<math {Ns}>{inner}</math>"));
    }

    [Fact]
    public void AsFichasLevamAFonteDeMatematicaEOEstilo()
    {
        var omml = XElement.Parse(OmmlMath.ToOmml(
            $"<math {Ns}><mi>x</mi><mi>sin</mi><mi mathvariant=\"normal\">d</mi><mi>𝐯</mi><mi>ℝ</mi><mi>ℒ</mi><mtext>se</mtext></math>",
            false,
            null)!);
        var runs = omml.Elements(M + "r").ToList();
        Assert.All(runs, run => Assert.Equal(
            "Cambria Math",
            run.Element(OmmlW + "rPr")!.Element(OmmlW + "rFonts")!.Attribute(OmmlW + "ascii")!.Value));

        string Describe(XElement run)
        {
            var properties = run.Element(M + "rPr");
            var nor = properties?.Element(M + "nor") is null ? string.Empty : "nor";
            var scr = properties?.Element(M + "scr")?.Attribute(M + "val")?.Value ?? string.Empty;
            var sty = properties?.Element(M + "sty")?.Attribute(M + "val")?.Value ?? string.Empty;
            return $"{run.Element(M + "t")!.Value}|{nor}{scr}|{sty}";
        }

        Assert.Equal(
            ["x||", "sind||p", "v||b", "R|double-struck|", "L|script|", "se|nor|"],
            runs.Select(Describe));
    }

    private static readonly XNamespace OmmlW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData(null, "center")]
    [InlineData("right", "right")]
    [InlineData("bogus", "center")]
    public void ADeExibicaoViraOMathParaComAlinhamento(string? jc, string expected)
    {
        var omml = XElement.Parse(OmmlMath.ToOmml($"<math {Ns} display=\"block\"><mi>x</mi></math>", true, jc)!);
        Assert.Equal(M + "oMathPara", omml.Name);
        Assert.Equal(expected, omml.Element(M + "oMathParaPr")!.Element(M + "jc")!.Attribute(M + "val")!.Value);
        Assert.Single(omml.Elements(M + "oMath"));
    }

    [Fact]
    public void MathMlMalFormadoNaoViraOmml()
    {
        Assert.Null(OmmlMath.ToOmml("<math><mi>x</math>", false, null));
        Assert.Null(OmmlMath.ToOmml("<p>x</p>", false, null));
    }

    // --- ida e volta: OMML → MathML → OMML -----------------------------------

    public static TheoryData<string> Fixtures() =>
    [
        Inline,
        Display,
        LibreOffice,
        TwoLines,
        // As construções que a tela desenha e os fixtures não trazem.
        $"<w:p><m:oMath><m:sPre><m:sub>{R("a")}</m:sub><m:sup>{R("b")}</m:sup><m:e>{R("X")}</m:e></m:sPre>" +
        $"""<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr><m:e>{R("v")}</m:e></m:acc>""" +
        $"""<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>{R("AB")}</m:e></m:bar>""" +
        $"""<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/></m:groupChrPr><m:e>{R("a+b")}</m:e></m:groupChr>""" +
        $"<m:limLow><m:e>{R("lim", "p")}</m:e><m:lim>{R("n→∞")}</m:lim></m:limLow>" +
        $"<m:limUpp><m:e>{R("x")}</m:e><m:lim>{R("2")}</m:lim></m:limUpp>" +
        $"<m:func><m:fName>{R("sin", "p")}</m:fName><m:e>{R("x")}</m:e></m:func>" +
        $"<m:borderBox><m:e>{R("z")}</m:e></m:borderBox>" +
        $"<m:m><m:mr><m:e>{R("1")}</m:e><m:e>{R("0")}</m:e></m:mr><m:mr><m:e>{R("0")}</m:e><m:e>{R("1")}</m:e></m:mr></m:m>" +
        $"<m:eqArr><m:e>{R("a=1")}</m:e><m:e>{R("b=2")}</m:e></m:eqArr>" +
        $"<m:rad><m:deg>{R("3")}</m:deg><m:e>{R("x")}</m:e></m:rad>" +
        $"""<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>{R("n")}</m:num><m:den>{R("k")}</m:den></m:f>""" +
        $"""<m:nary><m:naryPr><m:chr m:val="∫"/><m:limLoc m:val="subSup"/></m:naryPr><m:sub>{R("0")}</m:sub><m:sup>{R("1")}</m:sup><m:e>{R("f")}</m:e></m:nary>""" +
        "<m:d><m:dPr><m:begChr m:val=\"{\"/><m:sepChr m:val=\",\"/><m:endChr m:val=\"}\"/></m:dPr>" +
        $"<m:e>{R("a")}</m:e><m:e>{R("b")}</m:e></m:d>" +
        $"<m:sSubSup><m:e>{R("x")}</m:e><m:sub>{R("i")}</m:sub><m:sup>{R("2")}</m:sup></m:sSubSup>" +
        "</m:oMath></w:p>",
    ];

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void OmmlMathMlOmmlMantemAEstrutura(string paragraph)
    {
        foreach (var omml in OmmlOf(WithMath(paragraph)).Concat(ParasOf(WithMath(paragraph))))
        {
            var original = XElement.Parse(omml);
            var converted = OmmlMath.Convert(omml);
            Assert.Empty(converted.Lossy);
            var back = XElement.Parse(OmmlMath.ToOmml(converted.MathMl, converted.Display, converted.Jc)!);
            Assert.Equal(Shape(new XElement("x", original)), Shape(new XElement("x", back)));
        }
    }

    /// <summary>Os `m:oMathPara` (a de exibição é comparada inteira, linhas incluídas).</summary>
    private static List<string> ParasOf(byte[] docx)
    {
        using var stream = new MemoryStream(docx);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.Document!.Descendants<DocumentFormat.OpenXml.Math.Paragraph>()
            .Select(math => math.OuterXml).ToList();
    }

    // --- gravação ------------------------------------------------------------

    /// <summary>Um MathML com todas as construções do mapa, como o Temml as escreve.</summary>
    private static readonly string Everything =
        $"<math {Ns}><mrow>" +
        "<mfrac><mi>a</mi><mi>b</mi></mfrac><msqrt><mi>x</mi></msqrt><mroot><mi>x</mi><mn>3</mn></mroot>" +
        "<msub><mi>x</mi><mi>i</mi></msub><msup><mi>x</mi><mn>2</mn></msup><msubsup><mi>x</mi><mi>i</mi><mn>2</mn></msubsup>" +
        "<msubsup><mrow></mrow><mi>a</mi><mi>b</mi></msubsup><mi>X</mi>" +
        "<msubsup><mo movablelimits=\"false\">∑</mo><mi>i</mi><mi>n</mi></msubsup><mi>x</mi><mo>+</mo>" +
        "<mrow><munderover><mo>∫</mo><mn>0</mn><mn>1</mn></munderover></mrow><mi>f</mi><mo>+</mo>" +
        "<mrow><mo fence=\"true\">(</mo><mi>x</mi><mo stretchy=\"true\">|</mo><mi>y</mi><mo fence=\"true\">)</mo></mrow>" +
        "<mrow><mo fence=\"true\">[</mo><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr></mtable><mo fence=\"true\">]</mo></mrow>" +
        "<mtable><mtr><mtd><mi>a</mi></mtd></mtr><mtr><mtd><mi>b</mi></mtd></mtr></mtable>" +
        "<mover><mi>x</mi><mo>ˆ</mo></mover><mover><mi>y</mi><mo stretchy=\"true\">‾</mo></mover>" +
        "<munder><mi>y</mi><mo stretchy=\"true\">⏟</mo></munder><munder><mi>lim</mi><mi>x</mi></munder>" +
        "<mover><mi>b</mi><mi>a</mi></mover><munderover><mi>A</mi><mi>a</mi><mi>b</mi></munderover>" +
        "<mrow><mi>sin</mi><mo>⁡</mo></mrow><mi>x</mi><mrow class=\"omml-caixa\"><mi>z</mi></mrow>" +
        "<mi>𝐯</mi><mi>ℝ</mi><mtext> e </mtext><mphantom><mi>q</mi></mphantom>" +
        "</mrow></math>";

    private static DocumentModelDto WithNewEquation(byte[] original, string mathMl, bool display, string? jc = null)
    {
        var model = Clone(Open(original));
        var math = Node.Of("math").With("omml", null).With("mathml", mathMl).With("latex", "x").With("display", display);
        if (jc is not null) math.With("jc", jc);
        model.Doc.Content!.Insert(0, Node.Of("paragraph", math));
        return model;
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AEquacaoNovaGravaOmmlValido(bool display)
    {
        var original = WithMath(MathTestsPlain);
        // Save confere o documento com o OpenXmlValidator.
        var (bytes, result) = Save(original, WithNewEquation(original, Everything, display, display ? "left" : null));

        Assert.Empty(result.Inventory.Lost);
        var xml = XmlOf(bytes);
        Assert.Contains("<m:f>", xml);
        Assert.Contains("<m:rad>", xml);
        Assert.Contains("<m:nary>", xml);
        Assert.Contains("Cambria Math", xml);
        if (display) Assert.Matches("<m:oMathPara[^>]*><m:oMathParaPr><m:jc m:val=\"left\" ?/></m:oMathParaPr>", xml);
        else Assert.DoesNotContain("oMathPara", xml);

        // Reaberta, a equação desenha o que foi gravado.
        var reread = Equations(Open(bytes)).Single();
        Assert.Equal("true", reread.Attrs!["editable"]!.ToJsonString());
        Assert.Contains("<mfrac>", reread.Attrs!["mathml"]!.GetValue<string>());
    }

    private const string MathTestsPlain = """<w:p><w:r><w:t>Intocado.</w:t></w:r></w:p>""";

    [Fact]
    public void EditarAEquacaoTrocaOOmmlEReescreveOParagrafo()
    {
        var original = WithMath();
        var model = Clone(Open(original));
        var first = Equations(model)[0];
        first.With("omml", null).With("mathml", $"<math {Ns}><mfrac><mi>a</mi><mi>b</mi></mfrac></math>");

        var (bytes, result) = Save(original, model);
        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);
        var after = OmmlOf(bytes);
        Assert.Contains("<m:f>", after[0]);
        Assert.Equal(OmmlOf(original).Skip(1), after.Skip(1));
    }

    [Fact]
    public void EditarEquacaoFormatadaPorDentroAvisaUmaVez()
    {
        var colored =
            "<w:p><m:oMath><m:r><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/><w:color w:val=\"FF0000\"/></w:rPr><m:t>x</m:t></m:r></m:oMath>" +
            "<m:oMath><m:r><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr><m:t>y</m:t></m:r></m:oMath></w:p>";
        var original = WithMath(colored);

        // Sem editar a equação, nada se perde, mesmo reescrevendo o parágrafo.
        var untouched = Clone(Open(original));
        untouched.Doc.Content![0].Content!.Add(new Node { Type = "text", Text = " fim" });
        Assert.Empty(Save(original, untouched).Result.Inventory.Lost);

        var model = Clone(Open(original));
        foreach (var equation in Equations(model)) equation.With("omml", null);
        var (_, result) = Save(original, model);
        Assert.Equal(["formatação dentro da equação editada"], result.Inventory.Lost);
    }

    [Fact]
    public void AImpressaoDigitalDaEquacaoSemOmmlVeOMathMl()
    {
        var original = WithMath(MathTestsPlain);
        var model = WithNewEquation(original, $"<math {Ns}><mi>x</mi></math>", false);
        var other = WithNewEquation(original, $"<math {Ns}><mi>y</mi></math>", false);
        Assert.NotEqual(
            model.Doc.Content![0].Fingerprint(),
            other.Doc.Content![0].Fingerprint());
    }

    [Fact]
    public void TrocarAEquacaoComOControleLigadoGravaExclusaoEInsercao()
    {
        // O editor troca a equação como exclusão da antiga e inserção da nova (ver
        // `track-input.ts`): as duas viajam no mesmo parágrafo, com as marcas.
        var original = WithMath();
        var model = Clone(Open(original));
        var old = Equations(model)[0];
        var paragraph = model.Doc.Content!.First(block => block.Content?.Contains(old) == true);
        Mark Revision(string type) => new()
        {
            Type = type,
            Attrs = new() { ["author"] = "Ana", ["date"] = "2026-10-01T12:34:00Z", ["rid"] = type == "deletion" ? "901" : "902" },
        };
        old.Marks = [Revision("deletion")];
        var replacement = Node.Of("math").With("omml", null).With("display", false)
            .With("mathml", $"<math {Ns}><mfrac><mi>a</mi><mi>b</mi></mfrac></math>");
        replacement.Marks = [Revision("insertion")];
        paragraph.Content!.Insert(paragraph.Content.IndexOf(old) + 1, replacement);

        var (bytes, result) = Save(original, model);
        Assert.Empty(result.Inventory.Lost);
        var xml = XmlOf(bytes);
        Assert.Matches("<w:del [^>]*><m:oMath", xml);
        Assert.Matches("<w:ins [^>]*><m:oMath[^>]*>.*?<m:f>", xml);
    }

    [Fact]
    public void ConstrucaoDesconhecidaListaSoElaENaoOsArgumentos()
    {
        var converted = OmmlMath.Convert(
            $"<m:oMath xmlns:m=\"{Mns}\" xmlns:w=\"{W}\"><m:foo><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub/></m:foo></m:oMath>");
        Assert.Equal(["m:foo"], converted.Lossy);
        Assert.Contains("<mi>x</mi>", converted.MathMl);
    }
}
