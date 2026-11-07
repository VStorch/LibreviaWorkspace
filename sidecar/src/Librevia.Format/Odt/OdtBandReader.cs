using System.Text;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// A <c>style:header</c> or <c>style:footer</c> that is not one plain line → the three columns of
/// <see cref="BandDto"/>, the inverse of <see cref="OdtBandWriter"/>: each paragraph goes to the
/// column of its alignment, and each tab moves to the column of its tab stop (center, then right,
/// without stops). Page fields stay fields, any other field shows its text, and a bottom border on
/// the last paragraph is the rule. One way, like the DOCX band: the file keeps the band as it is.
/// What the columns cannot draw (a table, a positioned frame, underline) is noted.
/// </summary>
internal sealed class OdtBandReader(OdtReadContext context)
{
    private const int Left = 0;
    private const int Center = 1;
    private const int Right = 2;
    private const string NoBorder = "none";

    private readonly OdtFrameReader _frames = new(context);
    private List<PieceDto>[] _columns = [];
    private readonly HashSet<int> _started = [];
    private List<string> _stops = [];
    private int _column;
    private int _tabs;
    private bool _afterSpace;
    private bool _newLine;

    public BandDto Read(XElement band)
    {
        _columns = [[], [], []];
        var rule = false;
        foreach (var block in band.Elements())
        {
            if (block.Name != Ns.Text + "p" && block.Name != Ns.Text + "h")
            {
                context.NoteLoss(OdtNotices.BandFormatting);
                continue;
            }

            var chain = context.Styles.Chain(OdtStyles.ParagraphFamily, block.Attr(Ns.Text + "style-name"));
            ReadParagraph(block, chain);
            rule = HasBottomBorder(chain);
        }

        return new BandDto(_columns[Left], _columns[Center], _columns[Right], rule, [], []);
    }

    private void ReadParagraph(XElement paragraph, List<OdtStyle> chain)
    {
        _started.Clear();
        _stops = TabStopsOf(chain);
        _column = context.Styles.Paragraph(chain).TextAlign switch
        {
            "center" => Center,
            "right" => Right,
            _ => Left,
        };
        _tabs = 0;
        _afterSpace = true;
        _newLine = false;
        ReadChildren(paragraph, context.Styles.Character(chain, context.Styles.DefaultCharacter()));
    }

    private void ReadChildren(XElement container, CharacterFormat format)
    {
        foreach (var child in container.Nodes())
        {
            if (child is XText text) AppendCollapsing(text.Value, format);
            else if (child is XElement element) ReadElement(element, format);
        }
    }

    private void ReadElement(XElement element, CharacterFormat format)
    {
        if (element.Name.Namespace == Ns.Draw)
        {
            ReadDrawing(element);
            return;
        }

        if (element.Name.Namespace != Ns.Text)
        {
            context.NoteLoss(OdtNotices.BandFormatting);
            return;
        }

        switch (element.Name.LocalName)
        {
            case "page-number":
                AddPiece(Styled(new PieceDto(PieceDto.KindPageNumber), format));
                return;
            case "page-count":
                AddPiece(Styled(new PieceDto(PieceDto.KindTotalPages), format));
                return;
            case "s":
                AppendRaw(new string(' ', Math.Clamp(OdtLength.Integer(element.Attr(Ns.Text + "c")) ?? 1, 1, OdtInlineReader.MaxSpaces)), format);
                return;
            case "tab":
                NextTab();
                return;
            case "line-break":
                _newLine = true;
                _afterSpace = true;
                return;
            case "span":
                ReadChildren(element, context.Styles.Character(context.Styles.Chain(OdtStyles.TextFamily, element.Attr(Ns.Text + "style-name")), format));
                return;
            default:
                ReadChildren(element, format);
                return;
        }
    }

    /// <summary>A picture as a character goes in the column; any other drawing is positioned on the sheet.</summary>
    private void ReadDrawing(XElement element)
    {
        var image = element.Name.LocalName == "frame" && (element.Attr(Ns.Text + "anchor-type") ?? "as-char") == "as-char"
            ? _frames.Read(element)
            : null;
        if (image?.Type != "image")
        {
            context.NoteLoss(OdtNotices.BandFormatting);
            return;
        }

        AddPiece(PieceDto.Image(Attr.String(image, "src") ?? string.Empty, Attr.Int(image, "width") ?? 0, Attr.Int(image, "height") ?? 0));
    }

    private void NextTab()
    {
        _tabs++;
        _column = _tabs <= _stops.Count
            ? _stops[_tabs - 1] switch
            {
                "center" => Center,
                "right" => Right,
                _ => Left,
            }
            : Math.Min(Right, _tabs);
        _afterSpace = true;
    }

    private void AppendCollapsing(string value, CharacterFormat format)
    {
        var text = new StringBuilder(value.Length);
        foreach (var letter in value)
        {
            var space = letter is ' ' or '\t' or '\r' or '\n';
            if (space && _afterSpace) continue;
            text.Append(space ? ' ' : letter);
            _afterSpace = space;
        }

        AppendRaw(text.ToString(), format);
    }

    private void AppendRaw(string text, CharacterFormat format)
    {
        if (text.Length == 0) return;
        var piece = Styled(new PieceDto(PieceDto.KindText, text), format);
        var pieces = _columns[_column];
        if (!_newLine && _started.Contains(_column) && pieces is [.., { Kind: PieceDto.KindText } last] &&
            last with { Text = text, Line = false, Literal = false } == piece)
        {
            pieces[^1] = last with { Text = last.Text + text };
            return;
        }

        AddPiece(piece);
    }

    /// <summary>The literal <c>{n}</c> is text, not the page field.</summary>
    private void AddPiece(PieceDto piece)
    {
        var pieces = _columns[_column];
        var opensLine = pieces.Count > 0 && (_newLine || !_started.Contains(_column));
        _started.Add(_column);
        _newLine = false;
        if (piece.Kind != PieceDto.KindText) _afterSpace = false;
        var literal = piece.Text is { } text && (text.Contains("{n}", StringComparison.Ordinal) || text.Contains("{total}", StringComparison.Ordinal));
        pieces.Add(piece with { Line = opensLine, Literal = literal });
    }

    private PieceDto Styled(PieceDto piece, CharacterFormat format)
    {
        if (format.Underline == true || format.Strike == true || format.Caps == true || format.SmallCaps == true ||
            format.Position is CharacterFormat.Superscript or CharacterFormat.Subscript ||
            format.Background is { } background && background != CharacterFormat.Transparent)
        {
            context.NoteLoss(OdtNotices.BandFormatting);
        }

        return piece with
        {
            Bold = format.Bold == true,
            Italic = format.Italic == true,
            Color = format.Color,
            FontSize = format.FontSizePt is { } size ? RunReader.FormatPoints(size) : null,
            FontFamily = format.FontFamily,
        };
    }

    /// <summary>The stops of the nearest style that declares any, in position order.</summary>
    private static List<string> TabStopsOf(List<OdtStyle> chain)
    {
        var stops = chain.AsEnumerable().Reverse()
            .Select(style => style.Properties("paragraph-properties")?.Element(Ns.Style + "tab-stops"))
            .FirstOrDefault(declared => declared is not null);
        return [.. (stops?.Elements(Ns.Style + "tab-stop") ?? [])
            .OrderBy(stop => OdtLength.Millimetres(stop.Attr(Ns.Style + "position")) ?? 0)
            .Select(stop => stop.Attr(Ns.Style + "type") ?? "left")];
    }

    private static bool HasBottomBorder(List<OdtStyle> chain)
    {
        var border = chain.AsEnumerable().Reverse()
            .Select(style => style.Properties("paragraph-properties"))
            .Select(properties => properties?.Attr(Ns.Fo + "border-bottom") ?? properties?.Attr(Ns.Fo + "border"))
            .FirstOrDefault(declared => declared is not null);
        return border is not null && border != NoBorder;
    }
}

/// <summary>
/// The bands of a master page. One plain centered line stays text the editor edits
/// (<see cref="OdtPageReader"/>); anything richer, or a page with its own first or left (even)
/// bands, becomes <see cref="BandDto"/>s. A variant the master page does not have repeats the
/// default band, as ODF draws it.
/// </summary>
internal sealed record OdtPageBands(
    string HeaderText,
    string FooterText,
    BandDto? Header,
    BandDto? Footer,
    BandDto? FirstHeader,
    BandDto? FirstFooter,
    BandDto? EvenHeader,
    BandDto? EvenFooter,
    bool TitlePage,
    bool EvenAndOdd)
{
    public static OdtPageBands Read(XElement? master, OdtStyles? styles, OdtPackage package, Inventory inventory)
    {
        var first = Shown(master, "header-first") || Shown(master, "footer-first");
        var even = Shown(master, "header-left") || Shown(master, "footer-left");
        var reader = styles is null ? null : new OdtBandReader(new OdtReadContext(package, styles, inventory));
        BandDto? Band(string kind) => reader is not null && Shown(master, kind) ? reader.Read(master!.Element(Ns.Style + kind)!) : null;

        var (headerText, header) = Default(master, "header", styles, first || even, inventory, Band);
        var (footerText, footer) = Default(master, "footer", styles, first || even, inventory, Band);
        return new OdtPageBands(
            headerText, footerText, header, footer,
            first ? Band("header-first") ?? header : null,
            first ? Band("footer-first") ?? footer : null,
            even ? Band("header-left") ?? header : null,
            even ? Band("footer-left") ?? footer : null,
            first,
            even);
    }

    /// <summary>The text when the band is one plain line and no variant needs it as a band.</summary>
    private static (string Text, BandDto? Band) Default(
        XElement? master, string kind, OdtStyles? styles, bool variants, Inventory inventory, Func<string, BandDto?> band)
    {
        var plain = new Inventory();
        var text = OdtPageReader.BandText(master, kind, styles, plain);
        if (!variants && plain.Lost.Count == 0) return (text ?? string.Empty, null);
        return (string.Empty, text is null && !variants ? null : band(kind));
    }

    private static bool Shown(XElement? master, string kind) =>
        master?.Element(Ns.Style + kind) is { } band && band.Attr(Ns.Style + "display") != "false";
}
