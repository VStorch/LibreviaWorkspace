using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>
/// **De mão única**: a parte OOXML original volta intacta, e daqui sai só o que a
/// tela e o PDF desenham. Três colunas e filete opcional, como o cabeçalho
/// corporativo; o desenho ancorado vai em <c>Floats</c>, com a conta de posição do corpo.
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
/// O logotipo numa célula mesclada, o título ao lado. As bordas vêm como as iniciais
/// dos lados (<c>t</c>, <c>l</c>, <c>b</c>, <c>r</c>), já resolvidas: tela e papel não refazem a conta.
/// </summary>
public sealed record BandCellDto(
    [property: JsonPropertyName("pieces")] List<PieceDto> Pieces,
    /// <summary>De 0 a 1.</summary>
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
    /// <summary>Já como pilha de CSS.</summary>
    [property: JsonPropertyName("fontFamily")] string? FontFamily = null,
    /// <summary>Cada parágrafo do cabeçalho é uma linha.</summary>
    [property: JsonPropertyName("line")] bool Line = false,
    /// <summary>
    /// A relação, o parágrafo e a peça: a gravação escreve **só no <c>w:t</c> dela**. Número
    /// de página, imagem e cache de campo não têm <c>w:t</c> e não são editáveis.
    /// </summary>
    [property: JsonPropertyName("pid")] string? Pid = null,
    /// <summary>O texto traz <c>{n}</c> ou <c>{total}</c> escritos: é texto, e não campo.</summary>
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
