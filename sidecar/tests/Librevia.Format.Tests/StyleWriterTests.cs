using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// Saving styles: only the `w:style` that changed changes, and the rest of the package stays what
/// it was.
/// </summary>
public class StyleWriterTests
{
    private static DocumentModelDto Changed(byte[] original, Func<StyleSheetDto, StyleSheetDto> change)
    {
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        return model with { Styles = change(model.Styles!) };
    }

    private static StyleSheetDto WithStyle(StyleSheetDto sheet, StyleDefinitionDto style) =>
        sheet with { Styles = new Dictionary<string, StyleDefinitionDto>(sheet.Styles) { [style.Id] = style } };

    /// <summary>Each `w:style` in the XML, by id.</summary>
    private static Dictionary<string, string> StylesOf(string xml) =>
        Regex.Matches(xml, "<w:style [^>]*w:styleId=\"([^\"]+)\"[^>]*>.*?</w:style>", RegexOptions.Singleline)
            .ToDictionary(match => match.Groups[1].Value, match => match.Value);

    [Fact]
    public void SemMudancaOsEstilosVoltamByteAByte()
    {
        var original = Fixtures.WithStyles();
        var saved = Roundtrip.Save(original, Roundtrip.Clone(Roundtrip.Open(original)));

        Assert.Equal(Roundtrip.PartsOf(original)["word/styles.xml"], Roundtrip.PartsOf(saved.Bytes)["word/styles.xml"]);
    }

    [Fact]
    public void MudarUmEstiloMexeSoNele()
    {
        var original = Fixtures.WithStyles();
        var model = Changed(original, sheet =>
        {
            var faixa = sheet.Styles["Faixa"];
            return WithStyle(sheet, faixa with
            {
                Paragraph = faixa.Paragraph! with { Background = "#1f4e79", SpaceAfter = 6 },
                Character = faixa.Character! with { FontSize = "12pt", Italic = true },
            });
        });

        var saved = Roundtrip.Save(original, model);

        // The body did not change: no block rewritten.
        Assert.Equal(0, saved.Result.RewrittenBlocks);
        Assert.Equal(Roundtrip.PartsOf(original)["word/document.xml"], Roundtrip.PartsOf(saved.Bytes)["word/document.xml"]);

        // In `styles.xml`, only the band's `w:style` differs.
        var before = StylesOf(Roundtrip.XmlOf(original, "word/styles.xml"));
        var after = StylesOf(Roundtrip.XmlOf(saved.Bytes, "word/styles.xml"));
        Assert.Equal(before.Keys.Order(), after.Keys.Order());
        Assert.Equal(["Faixa"], before.Keys.Where(id => before[id] != after[id]));

        // And the round trip through the reader returns what the model asked for.
        Assert.Equal(model.Styles!.Styles["Faixa"], Roundtrip.Open(saved.Bytes).Styles!.Styles["Faixa"]);
    }

    [Fact]
    public void TirarUmaPropriedadeRemoveOElemento()
    {
        var original = Fixtures.WithStyles();
        var model = Changed(original, sheet =>
        {
            var faixa = sheet.Styles["Faixa"];
            return WithStyle(sheet, faixa with
            {
                Paragraph = faixa.Paragraph! with { Background = null },
                Character = faixa.Character! with { Bold = null },
            });
        });

        var saved = Roundtrip.Save(original, model);
        var faixa = StylesOf(Roundtrip.XmlOf(saved.Bytes, "word/styles.xml"))["Faixa"];

        Assert.DoesNotContain("w:shd", faixa, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:b />", faixa, StringComparison.Ordinal);
        Assert.Equal(model.Styles!.Styles["Faixa"], Roundtrip.Open(saved.Bytes).Styles!.Styles["Faixa"]);
    }

    [Fact]
    public void EstiloNovoEntraComoPersonalizado()
    {
        var original = Fixtures.WithStyles();
        var novo = new StyleDefinitionDto(
            "Destaque", "Destaque", "paragraph", QFormat: true, Hidden: false, Custom: true,
            BasedOn: "Corpo",
            Paragraph: new StyleParagraphDto(TextAlign: "center", SpaceBefore: 12, KeepNext: true),
            Character: new StyleCharacterDto(Bold: true, Color: "#c00000"));
        var model = Changed(original, sheet => WithStyle(sheet, novo));

        var saved = Roundtrip.Save(original, model);
        var xml = StylesOf(Roundtrip.XmlOf(saved.Bytes, "word/styles.xml"))["Destaque"];

        Assert.Contains("w:customStyle=\"1\"", xml, StringComparison.Ordinal);
        Assert.Equal(novo, Roundtrip.Open(saved.Bytes).Styles!.Styles["Destaque"]);
    }

    [Fact]
    public void OsPadroesDoDocumentoTambemVoltam()
    {
        var original = Fixtures.WithStyles();
        var model = Changed(original, sheet => sheet with
        {
            Defaults = sheet.Defaults with
            {
                Character = sheet.Defaults.Character with { FontSize = "10pt" },
                Paragraph = sheet.Defaults.Paragraph with { SpaceAfter = 4 },
            },
        });

        var reopened = Roundtrip.Open(Roundtrip.Save(original, model).Bytes).Styles!;

        Assert.Equal(model.Styles!.Defaults, reopened.Defaults);
    }

    [Fact]
    public void RascunhoAntigoSoAcrescentaEstilos()
    {
        // Version 2 draft styles were invented in the migration: the new style goes in, and a
        // change to an existing one stays out of the file, with a warning.
        var original = Fixtures.WithStyles();
        var novo = new StyleDefinitionDto(
            "Destaque", "Destaque", "paragraph", QFormat: true, Hidden: false, Custom: true,
            Character: new StyleCharacterDto(Bold: true));
        var model = Changed(original, sheet => WithStyle(
            WithStyle(sheet, sheet.Styles["Faixa"] with { Character = new StyleCharacterDto(Bold: false) }),
            novo));

        var saved = Roundtrip.Save(original, Roundtrip.Clone(Roundtrip.OpenFlat(original)) with
        {
            Styles = model.Styles,
            Flatten = true,
        });

        var reopened = Roundtrip.Open(saved.Bytes).Styles!;
        Assert.Equal(0, saved.Result.RewrittenBlocks);
        Assert.Equal(novo, reopened.Styles["Destaque"]);
        Assert.Equal(Roundtrip.Open(original).Styles!.Styles["Faixa"], reopened.Styles["Faixa"]);
        Assert.Contains(saved.Result.Inventory.Lost, item => item.Contains("rascunho", StringComparison.Ordinal));
    }

    [Fact]
    public void MeioPontoArredondaParaLongeDoZeroEMedidaEmLinhasSai()
    {
        // `w:beforeLines` beats `w:before` in Word: without removing it, the new space would show
        // nowhere.
        byte[] original;
        using (var buffer = new MemoryStream())
        {
            buffer.Write(Fixtures.WithStyles());
            using (var document = WordprocessingDocument.Open(buffer, true))
            {
                var faixa = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>()
                    .First(style => style.StyleId == "Faixa");
                faixa.StyleParagraphProperties!.SpacingBetweenLines = new SpacingBetweenLines { Before = "120", BeforeLines = 50 };
            }

            original = buffer.ToArray();
        }

        var model = Changed(original, sheet =>
        {
            var faixa = sheet.Styles["Faixa"];
            return WithStyle(sheet, faixa with
            {
                Paragraph = faixa.Paragraph! with { SpaceBefore = 12 },
                Character = faixa.Character! with { FontSize = "10.25pt" },
            });
        });

        var xml = StylesOf(Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes, "word/styles.xml"))["Faixa"];
        Assert.Contains("<w:sz w:val=\"21\" />", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("beforeLines", xml, StringComparison.Ordinal);
        Assert.Contains("w:before=\"240\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void NomeDeEmbutidoNaoMudaEOAvisoSai()
    {
        var original = Fixtures.WithStyleSpacingAndDirectMargins();
        var model = Changed(original, sheet => WithStyle(sheet, sheet.Styles["Normal"] with { Name = "Meu Normal" }));
        var saved = Roundtrip.Save(original, model);

        Assert.Contains("w:val=\"Normal\"", Roundtrip.XmlOf(saved.Bytes, "word/styles.xml"), StringComparison.Ordinal);
        Assert.Contains(saved.Result.Inventory.Lost, item => item.Contains("Normal", StringComparison.Ordinal));
    }

    [Fact]
    public void ParagrafoQueApontaEstiloIndefinidoEntraNoInventario()
    {
        var original = Fixtures.WithStyles();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        model.Doc.Content![0].With("styleId", "NaoExiste");
        model.Doc.Content[0].Content![0].Text = "Editado.";

        var saved = Roundtrip.Save(original, model);
        Assert.Contains(saved.Result.Inventory.Lost, item => item.Contains("NaoExiste", StringComparison.Ordinal));
    }
}
