using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>
/// **One way**: the original OOXML part goes back intact, and only what screen and PDF draw comes
/// out of here. Three columns and an optional rule, like a corporate header; anchored drawings go
/// in <c>Floats</c>, with the body's position math.
/// </summary>
public sealed record BandDto(
    [property: JsonPropertyName("left")] List<PieceDto> Left,
    [property: JsonPropertyName("center")] List<PieceDto> Center,
    [property: JsonPropertyName("right")] List<PieceDto> Right,
    [property: JsonPropertyName("rule")] bool Rule,
    [property: JsonPropertyName("floats")] List<FloatDto>? Floats = null,
    [property: JsonPropertyName("rows")] List<BandRowDto>? Rows = null)
{
    public static BandDto Empty() => new([], [], [], false, [], []);

    [JsonIgnore]
    public bool IsEmpty =>
        Left.Count == 0
        && Center.Count == 0
        && Right.Count == 0
        && !Rule
        && (Floats is null || Floats.Count == 0)
        && (Rows is null || Rows.Count == 0);
}

public sealed record BandRowDto(
    [property: JsonPropertyName("cells")] List<BandCellDto> Cells);

/// <summary>
/// The logo in a merged cell, the title beside it. Borders come as side initials (<c>t</c>,
/// <c>l</c>, <c>b</c>, <c>r</c>), already resolved: screen and paper do not redo the math.
/// </summary>
public sealed record BandCellDto(
    [property: JsonPropertyName("pieces")] List<PieceDto> Pieces,
    /// <summary>From 0 to 1.</summary>
    [property: JsonPropertyName("width")] double Width,
    [property: JsonPropertyName("span")] int Span,
    [property: JsonPropertyName("rowSpan")] int RowSpan,
    [property: JsonPropertyName("align")] string? Align,
    [property: JsonPropertyName("borders")] string Borders);

public sealed record PieceDto(
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("text")] string? Text = null,
    [property: JsonPropertyName("src")] string? Src = null,
    [property: JsonPropertyName("width")] int? Width = null,
    [property: JsonPropertyName("height")] int? Height = null,
    [property: JsonPropertyName("bold")] bool Bold = false,
    [property: JsonPropertyName("italic")] bool Italic = false,
    [property: JsonPropertyName("color")] string? Color = null,
    [property: JsonPropertyName("fontSize")] string? FontSize = null,
    /// <summary>Already as a CSS stack.</summary>
    [property: JsonPropertyName("fontFamily")] string? FontFamily = null,
    /// <summary>Each header paragraph is a line.</summary>
    [property: JsonPropertyName("line")] bool Line = false,
    /// <summary>
    /// The relationship, the paragraph and the piece: saving writes **only into its <c>w:t</c>**.
    /// Page numbers, images and field caches have no <c>w:t</c> and are not editable.
    /// </summary>
    [property: JsonPropertyName("pid")] string? Pid = null,
    /// <summary>
    /// The text has <c>{n}</c> or <c>{total}</c> written: it is text, not a field.
    /// </summary>
    [property: JsonPropertyName("literal")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool Literal = false)
{
    public const string KindText = "text";
    public const string KindImage = "image";
    public const string KindPageNumber = "pageNumber";
    public const string KindTotalPages = "totalPages";

    public static PieceDto Image(string src, int width, int height) =>
        new(KindImage, Src: src, Width: width, Height: height);
}
