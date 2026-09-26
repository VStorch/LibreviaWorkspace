using System.Text.RegularExpressions;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// Seções (M9): todas lidas, todas gravadas, e a que ninguém tocou volta byte a byte.
/// </summary>
/// <remarks>
/// O documento tinha uma configuração de página só, lida da primeira seção, e
/// uma divergência entre seções virava perda declarada. Agora cada `w:sectPr` é
/// uma seção do modelo: o de parágrafo com id (o `sectionBreak` do parágrafo que
/// a encerra), o do corpo como `page`.
/// </remarks>
public class SectionsTests
{
    /// <summary>
    /// Os `w:sectPr` do corpo, em ordem, como texto.
    /// </summary>
    /// <remarks>
    /// Sem as declarações de namespace: o fixture não declara `r:` na raiz, e o
    /// SDK o declara no próprio elemento ao gravar. O documento real (Word,
    /// LibreOffice) declara tudo na raiz, e ali o texto sai igual.
    /// </remarks>
    private static List<string> SectionXmlOf(byte[] docx) =>
        [.. Regex.Matches(Roundtrip.XmlOf(docx), "<w:sectPr[ >].*?</w:sectPr>", RegexOptions.Singleline)
            .Select(match => Regex.Replace(match.Value, " xmlns:\\w+=\"[^\"]*\"", string.Empty))];

    private static Node Holder(DocumentModelDto model, string id) =>
        Roundtrip.Walk(model.Doc).First(node =>
            node.Attrs?.GetValueOrDefault("sectionBreak")?.GetValue<string>() == id);

    [Fact]
    public void LeTodasAsSecoesComGeometriaInicioENumeracao()
    {
        var result = DocxReader.Read(Fixtures.WithThreeSections());
        var model = result.Model;

        Assert.NotNull(model.Sections);
        Assert.Equal(["s1", "s2"], model.Sections!.Select(section => section.Id));
        Assert.Equal("portrait", model.Sections[0].Orientation);
        Assert.Equal("lowerRoman", model.Sections[0].PageNumberFormat);
        Assert.Equal("nextPage", model.Sections[0].Start);
        Assert.Equal("continuous", model.Sections[1].Start);

        // A última é a do corpo: paisagem, página ímpar, numeração reiniciada.
        Assert.Equal("landscape", model.Page.Orientation);
        Assert.Equal("oddPage", model.Page.Start);
        Assert.True(PageReader.TryStartOf(model.Page, out var start));
        Assert.Equal(1, start);

        // Divergir já não é perda: a seção existe no modelo.
        Assert.DoesNotContain(result.Inventory.Lost, message => message.Contains("seções", StringComparison.Ordinal));
    }

    [Fact]
    public void FaixaQueASecaoNaoDeclaraFicaHerdadaENaoRepetida()
    {
        var model = Roundtrip.Open(Fixtures.WithThreeSections());

        Assert.NotNull(model.Sections![0].Header);
        // A do meio e a última não declaram cabeçalho: herdam — nulo, e não cópia.
        Assert.Null(model.Sections[1].Header);
        Assert.Null(model.Page.Header);
    }

    [Fact]
    public void ParagrafoQueEncerraSecaoLevaOId()
    {
        var model = Roundtrip.Open(Fixtures.WithThreeSections());

        // A marca vazia continua marca (sem altura), e o parágrafo com texto
        // também encerra seção.
        Assert.True(Holder(model, "s1").Attrs!.ContainsKey("sectionMark"));
        Assert.Equal("Segunda seção, que termina aqui.",
            string.Concat(Roundtrip.Walk(Holder(model, "s2")).Select(node => node.Text)));
    }

    [Fact]
    public void AbrirEGravarSemEditarNaoReescreveNadaEDevolveCadaSecao()
    {
        var original = Fixtures.WithThreeSections();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var (saved, result) = Roundtrip.Save(original, model);

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Equal(SectionXmlOf(original), SectionXmlOf(saved));
    }

    [Fact]
    public void MudarAOrientacaoDeUmaSecaoSoTocaOWSectPrDela()
    {
        var original = Fixtures.WithThreeSections();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        model.Sections![1] = model.Sections[1] with { Orientation = "landscape" };

        var (saved, result) = Roundtrip.Save(original, model);
        var before = SectionXmlOf(original);
        var after = SectionXmlOf(saved);

        // O parágrafo da marca é preservado; só a configuração muda.
        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Equal(before[0], after[0]);
        Assert.Equal(before[2], after[2]);
        Assert.Contains("w:orient=\"landscape\"", after[1], StringComparison.Ordinal);

        var reread = Roundtrip.Open(saved);
        Assert.Equal("landscape", reread.Sections![1].Orientation);
        Assert.Equal("continuous", reread.Sections[1].Start);
    }

    [Fact]
    public void MudarInicioENumeracaoDeUmaSecao()
    {
        var original = Fixtures.WithThreeSections();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        model.Sections![1] = model.Sections[1] with
        {
            Start = "evenPage",
            PageNumberStart = PageReader.StartElement(5),
            TitlePage = true,
        };
        model = model with { Page = model.Page with { Start = "nextPage" } };

        var (saved, _) = Roundtrip.Save(original, model);
        var reread = Roundtrip.Open(saved);

        Assert.Equal("evenPage", reread.Sections![1].Start);
        Assert.True(PageReader.TryStartOf(reread.Sections[1], out var start));
        Assert.Equal(5, start);
        Assert.True(reread.Sections[1].TitlePage);
        // "Próxima página" é o padrão: sem `w:type`, como o Word grava.
        Assert.Equal("nextPage", reread.Page.Start);
        Assert.DoesNotContain("w:type", SectionXmlOf(saved)[2], StringComparison.Ordinal);
    }

    [Fact]
    public void ApagarAMarcaJuntaOTrechoASecaoDeBaixo()
    {
        var original = Fixtures.WithThreeSections();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        // Como o Word: excluída a quebra, o trecho de cima assume a seção de baixo.
        Holder(model, "s2").Attrs!.Remove("sectionBreak");
        model = model with { Sections = [model.Sections![0]] };

        var (saved, result) = Roundtrip.Save(original, model);
        var after = SectionXmlOf(saved);

        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Equal(2, after.Count);
        Assert.Equal(SectionXmlOf(original)[0], after[0]);
        Assert.Contains("Segunda seção", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
    }

    [Fact]
    public void MarcaNovaCopiaASecaoQuePartiuEAplicaOModelo()
    {
        var original = Fixtures.WithThreeSections();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        // Inserir → Quebra de seção no último parágrafo: a seção nova é a de cima,
        // cópia da que foi partida (a do corpo), agora em retrato.
        var last = model.Doc.Content!.Last();
        last.With("sectionBreak", "n1");
        model = model with
        {
            Sections = [.. model.Sections!, model.Page with { Id = "n1", Orientation = "portrait", Start = "continuous" }],
        };
        model.Doc.Content!.Add(Node.Of("paragraph", new Node { Type = "text", Text = "Depois da quebra." }));

        var (saved, _) = Roundtrip.Save(original, model);
        var reread = Roundtrip.Open(saved);

        Assert.Equal(["s1", "s2", "s3"], reread.Sections!.Select(section => section.Id));
        Assert.Equal("portrait", reread.Sections[2].Orientation);
        Assert.Equal("continuous", reread.Sections[2].Start);
        // Cópia da do corpo: o reinício da numeração veio junto.
        Assert.True(PageReader.TryStartOf(reread.Sections[2], out var start));
        Assert.Equal(1, start);
        Assert.Equal("landscape", reread.Page.Orientation);
    }

    [Fact]
    public void MarcaRepetidaPorParagrafoPartidoNaoDuplicaASecao()
    {
        var original = Fixtures.WithThreeSections();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var holder = Holder(model, "s2");
        var index = model.Doc.Content!.IndexOf(holder);
        model.Doc.Content.Insert(index, Roundtrip.Clone(model).Doc.Content![index]);

        var (saved, _) = Roundtrip.Save(original, model);

        Assert.Equal(3, SectionXmlOf(saved).Count);
    }

    [Fact]
    public void RascunhoDeAntesDasSecoesGravaComoAntes()
    {
        // O `.sdoc` < 6 não traz `sectionBreak` nem `sections`: a leitura de
        // referência também não, e nenhum bloco parece mudado.
        var original = Fixtures.WithThreeSections();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        foreach (var node in Roundtrip.Walk(model.Doc)) node.Attrs?.Remove("sectionBreak");
        model = model with { Sections = null, BeforeSections = true };

        var (saved, result) = Roundtrip.Save(original, model);

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Equal(SectionXmlOf(original), SectionXmlOf(saved));
    }
}
