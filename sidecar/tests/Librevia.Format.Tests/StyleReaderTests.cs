using DocumentFormat.OpenXml.Packaging;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// Os estilos como dado: o que o leitor extrai, e nas unidades de quem vai usar.
/// </summary>
/// <remarks>
/// A tela ainda não desenha a partir disto — o leitor continua achatando a
/// formatação em cada bloco. O que estes testes protegem é o **contrato**: as
/// unidades (pontos, milímetros, `12pt`), a entrelinha em forma bruta e o fato de
/// a herança **não** ser resolvida aqui. Uma unidade trocada só apareceria na
/// entrega em que a tela passa a nascer dos estilos, e aí como "o documento abriu
/// diferente" — longe da causa.
/// </remarks>
public class StyleReaderTests
{
    private static StyleSheetDto Read(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        return StyleReader.Read(document.MainDocumentPart!);
    }

    [Fact]
    public void PadroesDoDocumentoVemComFonteETamanho()
    {
        var sheet = Read(Fixtures.WithStyles());

        // O `w:docDefaults` é a base de toda a cascata, e é dele que a fonte do
        // documento vem — não do `Normal`.
        Assert.Equal("Calibri", sheet.Defaults.Character.FontFamily);
        Assert.Equal("11pt", sheet.Defaults.Character.FontSize);
    }

    [Fact]
    public void EstiloTrazParagrafoECaractereNasUnidadesDoEditor()
    {
        var faixa = Read(Fixtures.WithStyles()).Styles["Faixa"];

        Assert.Equal("Faixa", faixa.Name);
        Assert.Equal("paragraph", faixa.Type);
        Assert.Equal("center", faixa.Paragraph!.TextAlign);
        // Hexadecimal com `#`, como o atributo do bloco — e minúsculo, para que
        // duas escritas da mesma cor não pareçam duas cores.
        Assert.Equal("#943634", faixa.Paragraph.Background);
        Assert.Equal("Arial", faixa.Character!.FontFamily);
        Assert.Equal("10pt", faixa.Character.FontSize);
        Assert.True(faixa.Character.Bold);
        Assert.Equal("#ffffff", faixa.Character.Color);
    }

    [Fact]
    public void HerancaNaoEResolvidaAqui()
    {
        // `Corpo` herda de `Base`, que é Arial 10 pt. O estilo tem de sair daqui
        // dizendo **só o que ele diz**: resolver a cascata no sidecar é o que a
        // entrega seguinte faz no TS, onde a altura da linha de cada fonte já é
        // conhecida. Resolvido aqui, o `basedOn` perderia o sentido e mudar o
        // estilo pai deixaria de mudar os filhos.
        var sheet = Read(Fixtures.WithStyles());
        var body = sheet.Styles["Corpo"];

        Assert.Equal("Base", body.BasedOn);
        Assert.Equal("justify", body.Paragraph!.TextAlign);
        Assert.Null(body.Character);
    }

    [Fact]
    public void EspacamentoRecuoEEntrelinhaVemNaFormaCerta()
    {
        var normal = Read(Fixtures.WithStyleSpacingAndDirectMargins()).Styles["Normal"];
        var paragraph = normal.Paragraph!;

        // Zero explícito é preservado: "sem espaço antes" é uma instrução do
        // documento, e descartá-la deixaria a margem do editor reaparecer.
        Assert.Equal(0, paragraph.SpaceBefore);
        Assert.Equal(7, paragraph.SpaceAfter);
        Assert.Equal(12.7, paragraph.IndentMm);
        Assert.Equal(1.06, paragraph.IndentRightMm);
        // Deslocamento é recuo negativo de primeira linha, como no `text-indent`.
        Assert.Equal(-6.35, paragraph.FirstLineMm);

        // **Forma bruta**: 276/240 = 1,15 vez a altura natural da linha, e não o
        // número do CSS. A multiplicação depende da fonte, e a fonte pode vir do
        // estilo — quem multiplica é `line-metrics.ts`.
        Assert.Equal("multiple", paragraph.LineSpacing!.Kind);
        Assert.Equal(1.15, paragraph.LineSpacing.Factor);
        Assert.Null(paragraph.LineSpacing.Points);
    }

    [Fact]
    public void EstiloDeTabelaEDeNumeracaoFicamDeFora()
    {
        // O modelo do documento só representa estilo de parágrafo e de caractere.
        // Nada se perde: `word/styles.xml` volta ao arquivo byte a byte pela
        // gravação cirúrgica — estes estilos só não aparecem no painel.
        var sheet = Read(DocxTemplate.Create(Page()));

        Assert.DoesNotContain("TableNormal", sheet.Styles.Keys);
        Assert.DoesNotContain("TableGrid", sheet.Styles.Keys);
        Assert.DoesNotContain("NoList", sheet.Styles.Keys);
        Assert.Contains("Normal", sheet.Styles.Keys);
        Assert.Contains("Hyperlink", sheet.Styles.Keys);
    }

    [Fact]
    public void PacoteDoDocumentoNovoVoltaComATabelaDeDados()
    {
        // O outro lado do teste de contrato de `src/main/sidecar/builtin-styles.test.ts`:
        // lá o `BUILTIN_STYLES` do TS é comparado com a tabela do C#; aqui a
        // tabela do C# é comparada com o **arquivo** que ela gera. Com os dois, uma
        // medida só existe em um lugar — e o documento novo abre como foi gravado.
        var sheet = Read(DocxTemplate.Create(Page()));

        Assert.Equal(BuiltinStyles.All.Length, sheet.Styles.Count);
        Assert.Equal("Normal", sheet.Defaults.ParagraphStyleId);
        Assert.Equal("DefaultParagraphFont", sheet.Defaults.CharacterStyleId);
        // A fonte do padrão volta com a substituta genérica atrás, porque o
        // pacote declara em `word/fontTable.xml` que a Times New Roman é serifada.
        Assert.StartsWith(BuiltinStyles.BodyFont, sheet.Defaults.Character.FontFamily!, StringComparison.Ordinal);
        Assert.Equal("12pt", sheet.Defaults.Character.FontSize);

        foreach (var expected in BuiltinStyles.All)
        {
            var actual = sheet.Styles[expected.Id];

            Assert.Equal(expected.Name, actual.Name);
            Assert.Equal(expected.Character ? "character" : "paragraph", actual.Type);
            Assert.Equal(expected.BasedOn, actual.BasedOn);
            Assert.Equal(expected.Next, actual.Next);
            Assert.Equal(expected.UiPriority, actual.UiPriority);
            Assert.Equal(expected.QFormat, actual.QFormat);
            Assert.Equal(expected.Hidden || expected.SemiHidden, actual.Hidden);
            Assert.False(actual.Custom);

            Assert.Equal(expected.BeforePt, actual.Paragraph?.SpaceBefore);
            Assert.Equal(expected.AfterPt, actual.Paragraph?.SpaceAfter);
            Assert.Equal(expected.LineFactor, actual.Paragraph?.LineSpacing?.Factor);
            Assert.Equal(expected.IndentMm, actual.Paragraph?.IndentMm);
            Assert.Equal(expected.OutlineLevel, actual.Paragraph?.OutlineLevel);
            Assert.Equal(expected.KeepNext ? true : null, actual.Paragraph?.KeepNext);
            Assert.Equal(expected.ContextualSpacing ? true : null, actual.Paragraph?.ContextualSpacing);

            Assert.Equal(expected.SizePt is null ? null : PointsText(expected.SizePt.Value), actual.Character?.FontSize);
            Assert.Equal(expected.Bold ? true : null, actual.Character?.Bold);
            Assert.Equal(expected.Color, actual.Character?.Color);
            Assert.Equal(expected.Underline ? true : null, actual.Character?.Underline);
        }
    }

    [Fact]
    public void DocumentoSemEstiloNenhumNaoQuebraALeitura()
    {
        // Um `.docx` sem `word/styles.xml` é raro, mas existe — e a resposta certa
        // é uma folha de estilos vazia, não uma exceção.
        var sheet = Read(Fixtures.Simple());

        Assert.Empty(sheet.Styles);
        Assert.Null(sheet.Defaults.ParagraphStyleId);
        Assert.Null(sheet.Defaults.Character.FontFamily);
    }

    private static string PointsText(double points) =>
        points == Math.Floor(points)
            ? $"{(int)points}pt"
            : points.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "pt";

    private static PageSetupDto Page() =>
        new("A4", "portrait", new MarginsDto(25, 25, 25, 25), null, null);
}
