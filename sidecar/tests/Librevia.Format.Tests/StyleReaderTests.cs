using DocumentFormat.OpenXml.Packaging;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// Os estilos como dado, nas unidades de quem os usa (pontos, milímetros, <c>12pt</c>),
/// com a entrelinha em forma bruta e a herança sem resolver.
/// </summary>
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

        // A fonte do documento vem do `w:docDefaults`, e não do `Normal`.
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
        // Com `#` e minúsculo, para duas escritas da mesma cor não parecerem duas.
        Assert.Equal("#943634", faixa.Paragraph.Background);
        Assert.Equal("Arial", faixa.Character!.FontFamily);
        Assert.Equal("10pt", faixa.Character.FontSize);
        Assert.True(faixa.Character.Bold);
        Assert.Equal("#ffffff", faixa.Character.Color);
    }

    [Fact]
    public void HerancaNaoEResolvidaAqui()
    {
        // `Corpo` diz só o que ele diz: a cascata se resolve no TS, onde se conhece a
        // altura da linha de cada fonte, e mudar o pai continua mudando os filhos.
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

        // Zero explícito é instrução do documento, e não some.
        Assert.Equal(0, paragraph.SpaceBefore);
        Assert.Equal(7, paragraph.SpaceAfter);
        Assert.Equal(12.7, paragraph.IndentMm);
        Assert.Equal(1.06, paragraph.IndentRightMm);
        // Deslocamento é recuo negativo de primeira linha, como no `text-indent`.
        Assert.Equal(-6.35, paragraph.FirstLineMm);

        // Forma bruta: a multiplicação depende da fonte, e quem a faz é `line-metrics.ts`.
        Assert.Equal("multiple", paragraph.LineSpacing!.Kind);
        Assert.Equal(1.15, paragraph.LineSpacing.Factor);
        Assert.Null(paragraph.LineSpacing.Points);
    }

    [Fact]
    public void EstiloDeTabelaEDeNumeracaoFicamDeFora()
    {
        // Só estilo de parágrafo e de caractere; os outros voltam em `word/styles.xml` intocado.
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
        // O outro lado de `builtin-styles.test.ts`, que compara `LEGACY_STYLES` com a
        // tabela do C#: aqui a tabela é comparada com o arquivo que ela gera.
        var sheet = Read(DocxTemplate.Create(Page()));

        Assert.Equal(BuiltinStyles.All.Length, sheet.Styles.Count);
        Assert.Equal("Normal", sheet.Defaults.ParagraphStyleId);
        Assert.Equal("DefaultParagraphFont", sheet.Defaults.CharacterStyleId);
        // A substituta genérica vem atrás, porque `word/fontTable.xml` declara a Times serifada.
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
        // Sem `word/styles.xml`, uma folha vazia, e não exceção.
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
