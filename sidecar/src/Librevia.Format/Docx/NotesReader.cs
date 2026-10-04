using System.Globalization;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Como o documento numera as notas — `w:footnotePr`/`w:endnotePr`.
/// </summary>
/// <param name="NumFmt">`w:numFmt`: `decimal`, `lowerRoman`, `upperLetter`, `chicago`…</param>
/// <param name="Start">`w:numStart`: o número da primeira nota.</param>
/// <param name="Restart">`w:numRestart`: `continuous`, `eachSect`, `eachPage`.</param>
/// <param name="Pos">`w:pos`: `pageBottom`, `beneathText`, `sectEnd`, `docEnd`.</param>
public sealed record NotePrDto(
    [property: JsonPropertyName("numFmt")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? NumFmt = null,
    [property: JsonPropertyName("start")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    int? Start = null,
    [property: JsonPropertyName("restart")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Restart = null,
    [property: JsonPropertyName("pos")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Pos = null);

/// <summary>A numeração das notas de rodapé e das de fim, fora dos nós — como os estilos.</summary>
public sealed record NotesDto(
    [property: JsonPropertyName("footnotePr")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    NotePrDto? FootnotePr = null,
    [property: JsonPropertyName("endnotePr")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    NotePrDto? EndnotePr = null);

/// <summary>
/// A numeração das notas: a do <c>w:sectPr</c> do corpo vence a do <c>settings.xml</c>.
/// </summary>
internal static class NotesReader
{
    public static NotesDto? Read(MainDocumentPart part, Body body)
    {
        var section = body.Elements<SectionProperties>().LastOrDefault();
        var settings = part.DocumentSettingsPart?.Settings;

        var footnote = Merge(
            Describe(section?.GetFirstChild<FootnoteProperties>()),
            Describe(settings?.GetFirstChild<FootnoteDocumentWideProperties>()));
        var endnote = Merge(
            Describe(section?.GetFirstChild<EndnoteProperties>()),
            Describe(settings?.GetFirstChild<EndnoteDocumentWideProperties>()));

        return footnote is null && endnote is null ? null : new NotesDto(footnote, endnote);
    }

    private static NotePrDto? Merge(NotePrDto? section, NotePrDto? document)
    {
        if (section is null) return document;
        if (document is null) return section;
        return new NotePrDto(
            section.NumFmt ?? document.NumFmt,
            section.Start ?? document.Start,
            section.Restart ?? document.Restart,
            section.Pos ?? document.Pos);
    }

    /// <summary>Pelo nome local: os dois tipos de nota têm os mesmos filhos.</summary>
    private static NotePrDto? Describe(OpenXmlElement? properties)
    {
        if (properties is null) return null;

        string? ValueOf(string name) =>
            properties.ChildElements.FirstOrDefault(child => child.LocalName == name)?
                .GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;

        var start = int.TryParse(ValueOf("numStart"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var number)
            ? (int?)number
            : null;
        var dto = new NotePrDto(ValueOf("numFmt"), start, ValueOf("numRestart"), ValueOf("pos"));
        return dto == new NotePrDto() ? null : dto;
    }
}
