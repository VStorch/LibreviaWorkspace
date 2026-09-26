using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// Numeração de página: `w:pgNumType`, `w:titlePg`, `w:evenAndOddHeaders` e o
/// campo inserido numa faixa vinda do Word.
/// </summary>
public class PageNumberingTests
{
    private static byte[] WithNumbering(string? format, int? start, bool titlePage = false) =>
        Fixtures.WithSection(section =>
        {
            var type = new PageNumberType();
            if (format == "lowerRoman") type.Format = NumberFormatValues.LowerRoman;
            if (start is { } value) type.Start = value;
            if (type.HasAttributes) section.AddChild(type, throwOnError: false);
            if (titlePage) section.AddChild(new TitlePage(), throwOnError: false);
        });

    [Fact]
    public void LeFormatoInicioEInterruptores()
    {
        var page = Roundtrip.Open(WithNumbering("lowerRoman", 3, titlePage: true)).Page;

        Assert.Equal("lowerRoman", page.PageNumberFormat);
        Assert.Equal(3, page.PageNumberStart);
        Assert.True(page.TitlePage);
        Assert.False(page.EvenAndOddHeaders);
    }

    [Fact]
    public void AbrirEGravarSemMexerNaoTocaNaSecaoNemNasConfiguracoes()
    {
        var original = WithNumbering("lowerRoman", 3, titlePage: true);
        var (saved, result) = Roundtrip.Save(original, Roundtrip.Clone(Roundtrip.Open(original)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Contains("<w:pgNumType w:fmt=\"lowerRoman\" w:start=\"3\" />", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        Assert.Equal(Roundtrip.PartsOf(original)["word/settings.xml"], Roundtrip.PartsOf(saved)["word/settings.xml"]);
    }

    [Fact]
    public void OPainelMudaFormatoInicioEInterruptoresNoArquivo()
    {
        var original = WithNumbering(null, null);
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        model = model with
        {
            Page = model.Page with
            {
                PageNumberFormat = "upperRoman",
                PageNumberStart = 5,
                TitlePage = true,
                EvenAndOddHeaders = true,
            },
        };

        var (saved, _) = Roundtrip.Save(original, model);
        var page = Roundtrip.Open(saved).Page;

        Assert.Equal("upperRoman", page.PageNumberFormat);
        Assert.Equal(5, page.PageNumberStart);
        Assert.True(page.TitlePage);
        Assert.True(page.EvenAndOddHeaders);
        Assert.Contains("w:evenAndOddHeaders", Roundtrip.XmlOf(saved, "word/settings.xml"), StringComparison.Ordinal);

        // E desligar volta a tirar, sem deixar `w:pgNumType` vazio para trás.
        var back = Roundtrip.Clone(Roundtrip.Open(saved));
        back = back with
        {
            Page = back.Page with { PageNumberFormat = "decimal", PageNumberStart = null, TitlePage = false, EvenAndOddHeaders = false },
        };
        var (again, _) = Roundtrip.Save(saved, back);
        Assert.DoesNotContain("w:pgNumType", Roundtrip.XmlOf(again), StringComparison.Ordinal);
        Assert.DoesNotContain("w:titlePg", Roundtrip.XmlOf(again), StringComparison.Ordinal);
        Assert.DoesNotContain("w:evenAndOddHeaders", Roundtrip.XmlOf(again, "word/settings.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public void CampoInseridoNaFaixaDoWordViraCampoDeVerdade()
    {
        var original = Fixtures.WithFooterOfThreeLines();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var footer = model.Page.Footer!;
        footer.Center[2] = footer.Center[2] with { Text = "Folha {n} de {total} " };

        var (saved, _) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved, Roundtrip.PartsOf(saved).Keys.First(key => key.StartsWith("word/footer", StringComparison.Ordinal)));

        Assert.Matches("Folha .*w:fldSimple w:instr=\" PAGE \".* de .*w:fldSimple w:instr=\" NUMPAGES \"", xml);
        Assert.DoesNotContain("{n}", xml, StringComparison.Ordinal);

        // Relido, o campo é número de página, e não o "1" que ficou em cache.
        var reread = Roundtrip.Open(saved).Page.Footer!;
        var kinds = reread.Left.Concat(reread.Center).Concat(reread.Right).Select(piece => piece.Kind).ToList();
        Assert.Single(kinds, kind => kind == PieceDto.KindPageNumber);
        Assert.Contains(PieceDto.KindTotalPages, kinds);
        Assert.Empty(Regex.Matches(xml, "\\{total\\}"));
    }
}
