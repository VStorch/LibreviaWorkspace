using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A aparência da tabela nos dois sentidos: `w:tblGrid`, `w:tcBorders` e `w:shd`
/// ↔ atributos do nó do editor.
/// </summary>
/// <remarks>
/// ## Por que existe
///
/// O editor **já** deixava a pessoa arrastar a divisória das colunas — a opção
/// `resizable` do TableKit —, e a largura resultante morria no caminho: o
/// gravador devolvia o `w:tblGrid` do arquivo por posição e o número novo não
/// chegava a lugar nenhum. Era perda silenciosa, o defeito mais grave do
/// projeto, e é o motivo desta classe.
///
/// ## A regra que evita estrago
///
/// Toda propriedade daqui é **projetada** do XML para o modelo na leitura e
/// comparada de volta na gravação. Quando a projeção do original e o atributo do
/// modelo são iguais, o gravador **não toca** no XML: o `w:tcBorders` com
/// `thickThinSmallGap`, o `w:shd` com trama de 25% e o `w:tblGrid` com medida
/// que não fecha em pixel inteiro voltam byte a byte. Só a célula que a pessoa
/// formatou é reescrita — e é só nela que a aproximação de um estilo exótico
/// aparece, com a linha correspondente no inventário.
/// </remarks>
internal static class TableLook
{
    /// <summary>Margem de célula do Word quando nada a declara: 0 em cima e embaixo, 0,075" dos lados.</summary>
    private static readonly int[] WordCellMargins = [0, 108, 0, 108];

    /// <summary>
    /// A margem de célula que vale na tabela, em twips — cima, direita, baixo,
    /// esquerda.
    /// </summary>
    /// <remarks>
    /// Lado a lado, na ordem do Word: o `w:tblCellMar` da própria tabela, o do
    /// estilo dela e dos estilos em que ele se baseia, o do estilo de tabela
    /// padrão e, calando todos, 0/108. Enquanto a tela usava a medida fixa do
    /// editor (4 × 8 px), cada linha de tabela importada saía 8 px mais alta que
    /// no papel do Word e do LibreOffice — numa tabela de oitenta linhas, uma
    /// folha a mais.
    /// </remarks>
    public static int[] CellMargins(Table table, DocumentFormat.OpenXml.Packaging.MainDocumentPart part)
    {
        var sides = new int?[4];
        Fill(sides, table.GetFirstChild<TableProperties>()?.GetFirstChild<TableCellMarginDefault>());

        var styles = part.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .Where(style => style.Type?.Value == StyleValues.Table)
            .ToList() ?? [];
        var styleId = table.GetFirstChild<TableProperties>()?.GetFirstChild<TableStyle>()?.Val?.Value;
        FillFromChain(sides, styles, styleId);
        FillFromChain(sides, styles, styles.FirstOrDefault(style => style.Default?.Value == true)?.StyleId?.Value);

        return [.. sides.Select((side, index) => side ?? WordCellMargins[index])];
    }

    private static void FillFromChain(int?[] sides, List<Style> styles, string? styleId)
    {
        // Um teto de voltas contra `w:basedOn` circular, que arquivo quebrado tem.
        for (var hops = 0; styleId is not null && hops < 16; hops++)
        {
            var style = styles.FirstOrDefault(candidate => candidate.StyleId?.Value == styleId);
            if (style is null) return;
            Fill(sides, style.StyleTableProperties?.GetFirstChild<TableCellMarginDefault>());
            styleId = style.BasedOn?.Val?.Value;
        }
    }

    private static void Fill(int?[] sides, TableCellMarginDefault? margins)
    {
        if (margins is null) return;
        sides[0] ??= Twips(margins.TopMargin?.Width?.Value, margins.TopMargin?.Type?.Value);
        sides[1] ??= Twips(margins.TableCellRightMargin?.Width?.Value.ToString(CultureInfo.InvariantCulture), margins.TableCellRightMargin?.Type?.Value)
                     ?? Twips(margins.EndMargin?.Width?.Value, margins.EndMargin?.Type?.Value);
        sides[2] ??= Twips(margins.BottomMargin?.Width?.Value, margins.BottomMargin?.Type?.Value);
        sides[3] ??= Twips(margins.TableCellLeftMargin?.Width?.Value.ToString(CultureInfo.InvariantCulture), margins.TableCellLeftMargin?.Type?.Value)
                     ?? Twips(margins.StartMargin?.Width?.Value, margins.StartMargin?.Type?.Value);
    }

    /// <summary>Só `dxa` (ou sem tipo) é medida; `nil` é zero; porcentagem não se aplica a margem.</summary>
    private static int? Twips(string? width, TableWidthUnitValues? type)
    {
        if (type == TableWidthUnitValues.Nil) return 0;
        if (type is not null && type != TableWidthUnitValues.Dxa) return null;
        return int.TryParse(width, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? Math.Max(0, value)
            : null;
    }

    private static int? Twips(string? width, TableWidthValues? type)
    {
        if (type == TableWidthValues.Nil) return 0;
        if (type is not null && type != TableWidthValues.Dxa) return null;
        return int.TryParse(width, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? Math.Max(0, value)
            : null;
    }

    /// <summary>1 pixel do CSS = 15 twips (1440/96).</summary>
    private const int TwipsPerPixel = 15;

    /// <summary>`w:sz` mede em oitavos de ponto.</summary>
    private const double EighthsPerPoint = 8;

    /// <summary>Espessura do `w:val` sem `w:sz`: meio ponto, como no Word.</summary>
    private const int DefaultBorderEighths = 4;

    private static readonly string[] Sides = ["top", "right", "bottom", "left"];

    public static int ToPixels(long twips) => (int)Math.Round((double)twips / TwipsPerPixel);

    public static int ToTwips(int pixels) => pixels * TwipsPerPixel;

    // --- leitura ------------------------------------------------------------

    /// <summary>
    /// A largura de cada coluna da grade, em pixels do CSS.
    /// </summary>
    /// <remarks>
    /// Do `w:tblGrid`, e não do `w:tcW` de cada célula: é a grade que o Word usa
    /// para desenhar a tabela de largura fixa, e é a grade que o editor espelha
    /// no `colgroup`. Tabela sem grade devolve lista vazia — e aí o modelo não
    /// declara largura nenhuma, que é o mesmo que dizer "deixe o navegador
    /// decidir".
    /// </remarks>
    public static List<int> GridWidths(Table table)
    {
        var widths = new List<int>();

        foreach (var column in table.GetFirstChild<TableGrid>()?.Elements<GridColumn>() ?? [])
        {
            if (!long.TryParse(column.Width?.Value, out var twips) || twips <= 0) return [];
            widths.Add(Math.Max(1, ToPixels(twips)));
        }

        return widths;
    }

    /// <summary>O `w:shd` da célula como `#rrggbb`, ou `null` quando não há sombreamento.</summary>
    public static string? Shading(TableCellProperties? properties)
    {
        var fill = properties?.Shading?.Fill?.Value;
        if (fill is null || fill.Length != 6) return null;
        if (!int.TryParse(fill, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)) return null;

        return "#" + fill.ToLowerInvariant();
    }

    /// <summary>
    /// O `w:tcBorders` no texto canônico do atributo — o mesmo que
    /// <c>table-format.ts</c> escreve.
    /// </summary>
    public static string? Borders(TableCellProperties? properties)
    {
        var borders = properties?.TableCellBorders;
        if (borders is null) return null;

        var text = new StringBuilder();
        foreach (var side in Sides)
        {
            if (BorderOf(borders, side) is not { } border) continue;
            if (text.Length > 0) text.Append(';');
            text.Append(side).Append(':').Append(Describe(border));
        }

        return text.Length == 0 ? null : text.ToString();
    }

    private static BorderType? BorderOf(TableCellBorders borders, string side) => side switch
    {
        "top" => borders.TopBorder,
        "right" => borders.RightBorder,
        "bottom" => borders.BottomBorder,
        "left" => borders.LeftBorder,
        _ => null,
    };

    private static string Describe(BorderType border)
    {
        var style = StyleOf(border.Val?.Value);
        var eighths = border.Size?.Value ?? DefaultBorderEighths;
        var points = Math.Max(0.25, eighths / EighthsPerPoint);
        var color = ColorOf(border.Color?.Value);

        return $"{style},{points.ToString("0.##", CultureInfo.InvariantCulture)},{color}";
    }

    /// <summary>
    /// Estilo do OOXML → estilo que o CSS desenha.
    /// </summary>
    /// <remarks>
    /// O OOXML tem vinte e tantos; o CSS desenha cinco. Os que não têm
    /// equivalente viram `single`, que é a linha que eles todos são de longe. A
    /// aproximação **não** custa nada enquanto a pessoa não formatar a célula:
    /// até lá o XML original é que volta para o arquivo.
    /// </remarks>
    private static string StyleOf(BorderValues? value)
    {
        if (value is null) return "single";
        if (value == BorderValues.Nil || value == BorderValues.None) return "none";
        if (value == BorderValues.Double) return "double";
        if (value == BorderValues.Dotted) return "dotted";
        if (value == BorderValues.Dashed || value == BorderValues.DashSmallGap) return "dashed";
        return "single";
    }

    /// <summary>Se este estilo volta ao arquivo como veio, ou se a gravação o aproxima.</summary>
    private static bool IsExact(BorderValues? value) =>
        value is null
        || value == BorderValues.Nil
        || value == BorderValues.None
        || value == BorderValues.Single
        || value == BorderValues.Double
        || value == BorderValues.Dotted
        || value == BorderValues.Dashed;

    /// <summary>
    /// `w:color` → `#rrggbb`. `auto` vira preto, que é o que o Word desenha.
    /// </summary>
    private static string ColorOf(string? value)
    {
        if (value is null || value.Length != 6) return "#000000";
        return int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)
            ? "#" + value.ToLowerInvariant()
            : "#000000";
    }

    // --- gravação -----------------------------------------------------------

    /// <summary>
    /// Escreve o `w:shd` que o modelo pede, se ele mudou.
    /// </summary>
    /// <returns>Se o elemento foi reescrito.</returns>
    public static bool ApplyShading(TableCellProperties properties, string? model, Inventory inventory)
    {
        if (string.Equals(Shading(properties), model, StringComparison.Ordinal)) return false;

        // Trama — `w:shd w:val="pct25"` e companhia — não tem representação no
        // modelo: o que a tela mostra é a cor de preenchimento. Trocar a cor de
        // uma célula assim achata a trama, e é isto que o aviso diz.
        var previous = properties.Shading?.Val?.Value;
        if (previous is not null && previous != ShadingPatternValues.Clear && previous != ShadingPatternValues.Nil)
        {
            inventory.NoteLoss("trama de sombreamento de uma célula que você formatou");
        }

        if (model is null)
        {
            properties.Shading = null;
            return true;
        }

        properties.Shading = new Shading
        {
            Val = ShadingPatternValues.Clear,
            Color = "auto",
            Fill = model[1..].ToUpperInvariant(),
        };
        return true;
    }

    /// <summary>Escreve o `w:tcBorders` que o modelo pede, se ele mudou.</summary>
    /// <remarks>
    /// Lado a lado, dentro do elemento que a célula já tem — e não um
    /// `w:tcBorders` novo no lugar dele. O modelo só conhece os quatro lados; a
    /// diagonal (`w:tl2br`, `w:tr2bl`), as bordas internas, o `w:start`/`w:end`
    /// e, dentro de um lado, o `w:space`, a sombra e a cor de tema não têm onde
    /// morar nele. Trocar o elemento inteiro os apagava sem aviso quando a pessoa
    /// mudava um lado só, porque a conferência de perda olhava apenas os quatro.
    ///
    /// O lado cuja descrição não mudou não é tocado, nem aproximado: é o XML
    /// original que fica. No lado que mudou, só o campo que mudou é reescrito.
    /// </remarks>
    public static bool ApplyBorders(TableCellProperties properties, string? model, Inventory inventory)
    {
        if (string.Equals(Borders(properties), model, StringComparison.Ordinal)) return false;

        var wanted = ParseSides(model);
        var borders = properties.TableCellBorders ?? new TableCellBorders();
        var approximated = false;

        foreach (var side in Sides)
        {
            var current = BorderOf(borders, side);
            wanted.TryGetValue(side, out var target);

            if (current is not null && target is not null && Describe(current) == string.Join(',', target)) continue;
            if (current is null && target is null) continue;

            if (current is not null && !IsExact(current.Val?.Value)) approximated = true;

            if (target is null)
            {
                SetSide(borders, side, null);
                continue;
            }

            if (current is null)
            {
                SetSide(borders, side, side switch
                {
                    "top" => Build<TopBorder>(target),
                    "right" => Build<RightBorder>(target),
                    "bottom" => Build<BottomBorder>(target),
                    _ => Build<LeftBorder>(target),
                });
                continue;
            }

            Update(current, target);
        }

        // Trocar uma borda de estilo que a tela só aproxima achata o estilo — é
        // isto que o aviso diz, e só para o lado que de fato mudou.
        if (approximated) inventory.NoteLoss("estilo de borda de uma célula que você formatou");

        properties.TableCellBorders = borders.HasChildren ? borders : null;
        return true;
    }

    /// <summary>O atributo do modelo, lado a lado: estilo, espessura e cor.</summary>
    private static Dictionary<string, string[]> ParseSides(string? model)
    {
        var sides = new Dictionary<string, string[]>(StringComparer.Ordinal);
        foreach (var part in (model ?? "").Split(';'))
        {
            var colon = part.IndexOf(':', StringComparison.Ordinal);
            if (colon <= 0) continue;

            var pieces = part[(colon + 1)..].Split(',');
            if (pieces.Length < 3 || !Sides.Contains(part[..colon])) continue;
            sides[part[..colon]] = pieces;
        }

        return sides;
    }

    /// <summary>
    /// Pelos acessores com tipo, e nunca por `AppendChild`: o `w:tcBorders` é uma
    /// **sequência** — topo, esquerda, baixo, direita, internas, diagonais —, e
    /// anexado fora dela o documento sai do esquema e o Word o recusa.
    /// </summary>
    private static void SetSide(TableCellBorders borders, string side, BorderType? border)
    {
        switch (side)
        {
            case "top": borders.TopBorder = (TopBorder?)border; break;
            case "right": borders.RightBorder = (RightBorder?)border; break;
            case "bottom": borders.BottomBorder = (BottomBorder?)border; break;
            case "left": borders.LeftBorder = (LeftBorder?)border; break;
        }
    }

    /// <summary>
    /// O lado que já existe, com só o que mudou reescrito.
    /// </summary>
    /// <remarks>
    /// A cor de tema sai quando a cor muda: no Word ela vence o `w:color`, e o
    /// lado continuaria com a cor antiga. O `w:space`, a sombra e o quadro ficam.
    /// </remarks>
    private static void Update(BorderType border, string[] target)
    {
        var (style, width, color) = (target[0], target[1], target[2]);
        var before = Describe(border).Split(',');

        if (before[0] != style) border.Val = ValueOf(style);
        if (before[1] != width) border.Size = SizeOf(width);
        if (before[2] != color)
        {
            if (color.Length == 7 && color[0] == '#') border.Color = color[1..].ToUpperInvariant();
            border.ThemeColor = null;
            border.ThemeTint = null;
            border.ThemeShade = null;
        }
    }

    private static BorderValues ValueOf(string style) => style switch
    {
        "none" => BorderValues.Nil,
        "double" => BorderValues.Double,
        "dotted" => BorderValues.Dotted,
        "dashed" => BorderValues.Dashed,
        _ => BorderValues.Single,
    };

    /// <remarks>
    /// O `w:sz` é obrigatório em faixa: zero faz o Word ignorar a borda, e acima
    /// de 255 oitavos o documento sai fora do esquema.
    /// </remarks>
    private static uint SizeOf(string width)
    {
        var points = double.TryParse(width, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0.5;
        return (uint)Math.Clamp(Math.Round(points * EighthsPerPoint), 2, 255);
    }

    private static TBorder Build<TBorder>(string[] pieces) where TBorder : BorderType, new()
    {
        var (style, width, color) = (pieces[0], pieces[1], pieces[2]);
        var border = new TBorder { Val = ValueOf(style), Size = SizeOf(width) };
        if (color.Length == 7 && color[0] == '#') border.Color = color[1..].ToUpperInvariant();
        return border;
    }
}
