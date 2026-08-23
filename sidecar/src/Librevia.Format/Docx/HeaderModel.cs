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

/// <summary>Uma linha da grade do cabeçalho.</summary>
public sealed record BandRowDto(
    [property: JsonPropertyName("cells")] List<BandCellDto> Cells);

/// <summary>
/// Uma célula da grade: o que está escrito nela e o retângulo que ela ocupa.
/// </summary>
/// <remarks>
/// A grade é a outra metade do cabeçalho corporativo, e a que não cabia em três
/// colunas. No corpus real ela traz o logotipo numa célula mesclada por quatro
/// linhas, o título ao lado e a numeração à direita — e achatada em esquerda,
/// centro e direita virava uma sopa de palavras por cima do texto.
///
/// As bordas vêm como as iniciais dos lados que existem — `t`, `l`, `b`, `r`.
/// Não é economia de bytes: é que o OOXML resolve cada lado por três caminhos
/// (a borda da célula, a da tabela, a interna), e o resultado dessa conta é um
/// sim ou não por lado. Guardar a conta feita evita refazê-la em dois
/// desenhistas diferentes, que é como tela e papel divergem.
/// </remarks>
public sealed record BandCellDto(
    [property: JsonPropertyName("pieces")] List<PieceDto> Pieces,
    /// <summary>Fração da largura da grade, de 0 a 1.</summary>
    [property: JsonPropertyName("width")] double Width,
    [property: JsonPropertyName("span")] int Span,
    [property: JsonPropertyName("rowSpan")] int RowSpan,
    [property: JsonPropertyName("align")] string? Align,
    [property: JsonPropertyName("borders")] string Borders);

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
