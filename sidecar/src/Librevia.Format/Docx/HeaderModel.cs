using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>
/// Cabeçalho ou rodapé em forma exibível.
/// </summary>
/// <remarks>
/// **De mão única.** O que vai para o arquivo é a parte OOXML original, copiada
/// intacta pela gravação cirúrgica; daqui sai só o que a tela e o PDF precisam
/// desenhar. É isso que torna esta extração barata: um erro aqui é cosmético,
/// não perda de dados.
///
/// O modelo é uma **faixa de três colunas com filete opcional** — esquerda,
/// centro, direita — que é como o Word sempre pensou cabeçalho e como quase todo
/// cabeçalho corporativo é montado. É o bastante para o texto.
///
/// O que **não** cabe em três colunas é o desenho ancorado: ele traz posição de
/// verdade e pode vir girado, e a marca lateral do corpus é uma faixa de 28,6 mm
/// em pé — não entra numa banda de 10 mm de altura. Esses saem em `Floats`, com
/// a mesma descrição dos objetos ancorados do corpo, e são desenhados pela mesma
/// conta de posição.
/// </remarks>
public sealed record BandDto(
    [property: JsonPropertyName("left")] List<PieceDto> Left,
    [property: JsonPropertyName("center")] List<PieceDto> Center,
    [property: JsonPropertyName("right")] List<PieceDto> Right,
    [property: JsonPropertyName("rule")] bool Rule,
    [property: JsonPropertyName("floats")] List<FloatDto>? Floats = null)
{
    public static BandDto Empty() => new([], [], [], false, []);

    [JsonIgnore]
    public bool IsEmpty =>
        Left.Count == 0
        && Center.Count == 0
        && Right.Count == 0
        && !Rule
        && (Floats is null || Floats.Count == 0);
}

/// <summary>Um pedaço do cabeçalho: texto, imagem ou número de página.</summary>
public sealed record PieceDto(
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("text")] string? Text = null,
    [property: JsonPropertyName("src")] string? Src = null,
    [property: JsonPropertyName("width")] int? Width = null,
    [property: JsonPropertyName("height")] int? Height = null,
    [property: JsonPropertyName("bold")] bool Bold = false,
    [property: JsonPropertyName("italic")] bool Italic = false,
    [property: JsonPropertyName("color")] string? Color = null,
    [property: JsonPropertyName("fontSize")] string? FontSize = null)
{
    public const string KindText = "text";
    public const string KindImage = "image";
    public const string KindPageNumber = "pageNumber";
    public const string KindTotalPages = "totalPages";

    public static PieceDto Image(string src, int width, int height) =>
        new(KindImage, Src: src, Width: width, Height: height);
}
