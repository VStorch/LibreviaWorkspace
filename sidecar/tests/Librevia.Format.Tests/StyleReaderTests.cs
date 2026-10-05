using DocumentFormat.OpenXml.Packaging;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// Styles as data, in their users' units (points, millimetres, <c>12pt</c>), with line spacing in
/// raw form and inheritance unresolved.
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

        // The document font comes from `w:docDefaults`, not from `Normal`.
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
        // With `#` and lowercase, so two spellings of the same color do not look like two.
        Assert.Equal("#943634", faixa.Paragraph.Background);
        Assert.Equal("Arial", faixa.Character!.FontFamily);
        Assert.Equal("10pt", faixa.Character.FontSize);
        Assert.True(faixa.Character.Bold);
        Assert.Equal("#ffffff", faixa.Character.Color);
    }

    [Fact]
    public void HerancaNaoEResolvidaAqui()
    {
        // `Corpo` says only what it says: the cascade is resolved in TS, where each font's line
        // height is known, and changing the parent still changes the children.
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

        // An explicit zero is a document instruction, and does not vanish.
        Assert.Equal(0, paragraph.SpaceBefore);
        Assert.Equal(7, paragraph.SpaceAfter);
        Assert.Equal(12.7, paragraph.IndentMm);
        Assert.Equal(1.06, paragraph.IndentRightMm);
        // Hanging is a negative first-line indent, as in `text-indent`.
        Assert.Equal(-6.35, paragraph.FirstLineMm);

        // Raw form: the multiplication depends on the font, and `line-metrics.ts` does it.
        Assert.Equal("multiple", paragraph.LineSpacing!.Kind);
        Assert.Equal(1.15, paragraph.LineSpacing.Factor);
        Assert.Null(paragraph.LineSpacing.Points);
    }

    [Fact]
    public void EstiloDeTabelaEDeNumeracaoFicamDeFora()
    {
        // Only paragraph and character styles; the others go back in an untouched
        // `word/styles.xml`.
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
        // The other side of `builtin-styles.test.ts`, which compares `LEGACY_STYLES` with the C#
        // table: here the table is compared with the file it generates.
        var sheet = Read(DocxTemplate.Create(Page()));

        Assert.Equal(BuiltinStyles.All.Length, sheet.Styles.Count);
        Assert.Equal("Normal", sheet.Defaults.ParagraphStyleId);
        Assert.Equal("DefaultParagraphFont", sheet.Defaults.CharacterStyleId);
        // The generic substitute comes behind, because `word/fontTable.xml` declares Times as
        // serif.
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
        // Without `word/styles.xml`, an empty sheet, not an exception.
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
