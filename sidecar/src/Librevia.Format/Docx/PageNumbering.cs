using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A numeração de página e os interruptores das faixas, levados ao arquivo.
/// </summary>
/// <remarks>
/// Três lugares diferentes no OOXML para uma caixa só do painel: o formato e o
/// início moram em `w:sectPr/w:pgNumType`, a capa distinta em `w:sectPr/w:titlePg`
/// e as páginas pares em `w:settings/w:evenAndOddHeaders` — esta é do documento
/// inteiro, e não da seção, e por isso é a única que toca outra parte.
///
/// Cada um só é escrito quando o modelo diz algo **diferente** do que o arquivo
/// já diz. Campo nulo (o `.sdoc` de antes desta fase) não diz nada: o arquivo
/// fica como está. E o `w:pgNumType` só perde os atributos que o painel conhece —
/// `w:chapStyle` e companhia, que o editor não mostra, continuam lá.
/// </remarks>
internal static class PageNumbering
{
    public static void Apply(
        MainDocumentPart part,
        SectionProperties section,
        PageSetupDto page,
        HashSet<string> touched,
        Inventory inventory)
    {
        ApplyNumberType(section, page, inventory);

        if (page.TitlePage is { } title && title != PageReader.TitlePageOf(section))
        {
            section.RemoveAllChildren<TitlePage>();
            if (title && !section.AddChild(new TitlePage(), throwOnError: false))
            {
                inventory.NoteLoss("\"Primeira página diferente\" (o arquivo não aceitou o interruptor)");
            }
        }

        if (page.EvenAndOddHeaders is { } even && even != PageReader.EvenAndOddOf(part))
        {
            var settingsPart = part.DocumentSettingsPart ?? part.AddNewPart<DocumentSettingsPart>();
            var settings = settingsPart.Settings ??= new Settings();
            settings.RemoveAllChildren<EvenAndOddHeaders>();

            // Sem lugar no `w:settings` — a ordem dele é sequência rígida —, a
            // parte não é gravada e o aviso fica: marcar a parte como mudada
            // gravaria um `settings.xml` que não diz o que a tela mostra.
            if (even && !settings.AddChild(new EvenAndOddHeaders(), throwOnError: false))
            {
                inventory.NoteLoss("\"Pares e ímpares diferentes\" (o arquivo não aceitou o interruptor)");
                return;
            }

            settings.Save();
            touched.Add(settingsPart.Uri.ToString().TrimStart('/'));
        }
    }

    private static void ApplyNumberType(SectionProperties section, PageSetupDto page, Inventory inventory)
    {
        if (page.PageNumberFormat is null) return;

        var existing = section.GetFirstChild<PageNumberType>();
        var format = PageReader.PageNumberFormats.Contains(page.PageNumberFormat) ? page.PageNumberFormat : "decimal";
        var sameFormat = format == PageReader.PageNumberFormatOf(section);

        // Início ausente no modelo (rascunho de antes) é "não mexa": escolher só
        // o formato não pode apagar o `w:start` que o arquivo já tinha.
        var knowsStart = PageReader.TryStartOf(page, out var start);
        var sameStart = !knowsStart || start == existing?.Start?.Value;
        if (sameFormat && sameStart) return;

        var element = existing ?? new PageNumberType();
        if (!sameFormat)
        {
            element.Format = format == "decimal" ? null : new EnumValue<NumberFormatValues>(FormatOf(format));
        }

        if (!sameStart) element.Start = start;

        if (!element.HasAttributes)
        {
            element.Remove();
            return;
        }

        if (existing is null && !section.AddChild(element, throwOnError: false))
        {
            inventory.NoteLoss("formato e início da numeração de página (o arquivo não aceitou a mudança)");
        }
    }

    private static NumberFormatValues FormatOf(string name) => name switch
    {
        "lowerRoman" => NumberFormatValues.LowerRoman,
        "upperRoman" => NumberFormatValues.UpperRoman,
        "lowerLetter" => NumberFormatValues.LowerLetter,
        "upperLetter" => NumberFormatValues.UpperLetter,
        _ => NumberFormatValues.Decimal,
    };
}
