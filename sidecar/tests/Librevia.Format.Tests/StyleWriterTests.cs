using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// A gravação dos estilos: só o `w:style` que mudou muda, e o resto do pacote
/// continua sendo o que era.
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

    /// <summary>Cada `w:style` do XML, pelo id.</summary>
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

        // O corpo não mudou: nenhum bloco foi reescrito, e o texto do documento
        // é o mesmo de antes, byte a byte.
        Assert.Equal(0, saved.Result.RewrittenBlocks);
        Assert.Equal(Roundtrip.PartsOf(original)["word/document.xml"], Roundtrip.PartsOf(saved.Bytes)["word/document.xml"]);

        // Em `styles.xml`, só o `w:style` da faixa difere.
        var before = StylesOf(Roundtrip.XmlOf(original, "word/styles.xml"));
        var after = StylesOf(Roundtrip.XmlOf(saved.Bytes, "word/styles.xml"));
        Assert.Equal(before.Keys.Order(), after.Keys.Order());
        Assert.Equal(["Faixa"], before.Keys.Where(id => before[id] != after[id]));

        // E a ida e volta pelo leitor devolve o que o modelo pediu.
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
        // Os estilos de um rascunho da versão 2 foram inventados na migração: a
        // diferença deles para os do arquivo não é obra da pessoa, e gravá-la
        // reescreveria os estilos verdadeiros. O estilo novo entra; a mudança num
        // existente fica fora do arquivo e é declarada.
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
        // `w:beforeLines` vence `w:before` no Word: sem tirá-lo, o espaço novo não
        // apareceria em lugar nenhum.
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
