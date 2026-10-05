using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Page numbering and band switches: <c>w:sectPr/w:pgNumType</c>, <c>w:sectPr/w:titlePg</c> and
/// <c>w:settings/w:evenAndOddHeaders</c>, the latter document-wide. Only what differs from the file
/// is written; null says nothing, and <c>w:pgNumType</c> keeps the attributes the panel does not
/// know.
/// </summary>
internal static class PageNumbering
{
    public static void Apply(
        MainDocumentPart part,
        SectionProperties section,
        PageSetupDto page,
        HashSet<string> touched,
        Inventory inventory,
        bool documentWide = true)
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

        // Odd and even is document-wide: only the last section carries it.
        if (documentWide && page.EvenAndOddHeaders is { } even && even != PageReader.EvenAndOddOf(part))
        {
            var settingsPart = part.DocumentSettingsPart ?? part.AddNewPart<DocumentSettingsPart>();
            var settings = settingsPart.Settings ??= new Settings();
            settings.RemoveAllChildren<EvenAndOddHeaders>();

            // `w:settings` is a strict sequence: without a slot, the part is not written and the
            // warning stays.
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

        // A missing start means "leave alone": changing the format does not erase `w:start`.
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
