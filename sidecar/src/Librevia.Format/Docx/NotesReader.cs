using System.Globalization;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>How the document numbers notes: `w:footnotePr`/`w:endnotePr`.</summary>
/// <param name="NumFmt">`w:numFmt`: `decimal`, `lowerRoman`, `upperLetter`, `chicago`…</param>
/// <param name="Start">`w:numStart`: the first note's number.</param>
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

/// <summary>Footnote and endnote numbering, outside the nodes, like styles.</summary>
public sealed record NotesDto(
    [property: JsonPropertyName("footnotePr")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    NotePrDto? FootnotePr = null,
    [property: JsonPropertyName("endnotePr")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    NotePrDto? EndnotePr = null);

/// <summary>The body's <c>w:sectPr</c> numbering beats <c>settings.xml</c>'s.</summary>
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

    /// <summary>By local name: both note kinds have the same children.</summary>
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
