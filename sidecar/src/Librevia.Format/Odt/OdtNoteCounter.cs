using System.Globalization;
using System.Text;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The citation of each note, in text order, as <c>notes.ts</c> counts it: the start and format of
/// the model's numbering, back to the start at each section when it restarts there, and a custom
/// mark as it is, without counting.
/// </summary>
internal sealed class OdtNoteCounter(NotesDto? notes)
{
    public const string Footnote = "footnote";
    public const string Endnote = "endnote";

    private const string EachSection = "eachSect";
    private const string Chicago = "chicago";
    private const int LettersInAlphabet = 26;

    private static readonly string[] ChicagoMarks = ["*", "†", "‡", "§"];

    private static readonly (int Value, string Numeral)[] Romans =
    [
        (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
        (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
    ];

    private readonly Dictionary<bool, (int Ordinal, int Section)> _last = [];

    public static NotePrDto? NumberingOf(NotesDto? notes, bool endnote) => endnote ? notes?.EndnotePr : notes?.FootnotePr;

    public static string FormatOf(NotePrDto? numbering, bool endnote) => numbering?.NumFmt ?? (endnote ? "lowerRoman" : "decimal");

    public string Next(bool endnote, int section, string? mark)
    {
        if (!string.IsNullOrEmpty(mark)) return mark;
        var numbering = NumberingOf(notes, endnote);
        var restarts = numbering?.Restart == EachSection;
        var ordinal = _last.TryGetValue(endnote, out var last) && !(restarts && last.Section != section) ? last.Ordinal + 1 : 0;
        _last[endnote] = (ordinal, section);

        var value = (numbering?.Start ?? 1) + ordinal;
        var format = FormatOf(numbering, endnote);
        if (format == Chicago) return value <= 0 ? Decimal(value) : ChicagoOf(value);
        return Number(value, format) is { Length: > 0 } text ? text : Decimal(value);
    }

    /// <summary>The value in Word's format; <c>none</c> and <c>bullet</c> give nothing.</summary>
    public static string Number(int value, string format)
    {
        if (format is "none" or "bullet") return string.Empty;
        if (value <= 0 && format is not ("decimal" or "decimalZero")) return Decimal(value);
        return format switch
        {
            "decimalZero" => value is >= 0 and < 10 ? "0" + Decimal(value) : Decimal(value),
            "lowerLetter" => Letter(value),
            "upperLetter" => Letter(value).ToUpperInvariant(),
            "lowerRoman" => Roman(value),
            "upperRoman" => Roman(value).ToUpperInvariant(),
            _ => Decimal(value),
        };
    }

    private static string Decimal(int value) => value.ToString(CultureInfo.InvariantCulture);

    private static string ChicagoOf(int value) =>
        string.Concat(Enumerable.Repeat(ChicagoMarks[(value - 1) % ChicagoMarks.Length], ((value - 1) / ChicagoMarks.Length) + 1));

    private static string Letter(int value) =>
        new((char)('a' + ((value - 1) % LettersInAlphabet)), ((value - 1) / LettersInAlphabet) + 1);

    private static string Roman(int value)
    {
        var text = new StringBuilder();
        foreach (var (step, numeral) in Romans)
        {
            for (; value >= step; value -= step) text.Append(numeral);
        }

        return text.ToString();
    }
}
