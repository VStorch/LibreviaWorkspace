using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Librevia.Format.Pdf;

/// <summary>Text strings and dates, as the metadata dictionary stores them.</summary>
internal static partial class PdfText
{
    /// <summary>
    /// PDFDocEncoding: Latin-1, except these codes, which hold typographic signs instead of control
    /// characters. 0x9F and 0xAD are undefined.
    /// </summary>
    private static readonly Dictionary<byte, char> DocEncoding = new()
    {
        [0x18] = '˘', [0x19] = 'ˇ', [0x1A] = 'ˆ', [0x1B] = '˙', [0x1C] = '˝', [0x1D] = '˛', [0x1E] = '˚',
        [0x1F] = '˜', [0x80] = '•', [0x81] = '†', [0x82] = '‡', [0x83] = '…', [0x84] = '—', [0x85] = '–',
        [0x86] = 'ƒ', [0x87] = '⁄', [0x88] = '‹', [0x89] = '›', [0x8A] = '−', [0x8B] = '‰', [0x8C] = '„',
        [0x8D] = '“', [0x8E] = '”', [0x8F] = '‘', [0x90] = '’', [0x91] = '‚', [0x92] = '™', [0x93] = 'ﬁ',
        [0x94] = 'ﬂ', [0x95] = 'Ł', [0x96] = 'Œ', [0x97] = 'Š', [0x98] = 'Ÿ', [0x99] = 'Ž', [0x9A] = 'ı',
        [0x9B] = 'ł', [0x9C] = 'œ', [0x9D] = 'š', [0x9E] = 'ž', [0xA0] = '€',
    };

    public static string Decode(PdfString text)
    {
        var bytes = text.Bytes;
        if (bytes is [0xFE, 0xFF, ..]) return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes is [0xFF, 0xFE, ..]) return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes is [0xEF, 0xBB, 0xBF, ..]) return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);

        var builder = new StringBuilder(bytes.Length);
        foreach (var value in bytes) builder.Append(DocEncoding.TryGetValue(value, out var sign) ? sign : (char)value);
        return builder.ToString();
    }

    /// <summary>Plain ASCII stays one byte per letter; anything else goes in UTF-16, which every reader knows.</summary>
    public static PdfString Encode(string text)
    {
        if (text.All(letter => letter is >= ' ' and <= '~')) return new PdfString(Encoding.ASCII.GetBytes(text));
        return new PdfString([0xFE, 0xFF, .. Encoding.BigEndianUnicode.GetBytes(text)]);
    }

    [GeneratedRegex(@"^(?:D:)?(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?(?:([Zz])|([+-])(\d{2})'?(\d{2})?'?)?")]
    private static partial Regex PdfDate();

    /// <summary><c>D:20250102030405-03'00'</c> → <c>2025-01-02T03:04:05-03:00</c>. Null when unreadable.</summary>
    public static string? ToW3cdtf(string pdfDate)
    {
        var match = PdfDate().Match(pdfDate.Trim());
        if (!match.Success) return null;

        string Part(int group, string fallback) => match.Groups[group].Success ? match.Groups[group].Value : fallback;
        var date = $"{Part(1, "")}-{Part(2, "01")}-{Part(3, "01")}T{Part(4, "00")}:{Part(5, "00")}:{Part(6, "00")}";
        if (match.Groups[7].Success) return date + "Z";
        if (match.Groups[8].Success) return $"{date}{match.Groups[8].Value}{match.Groups[9].Value}:{Part(10, "00")}";
        return date;
    }

    /// <summary>Some writers fill the XMP dates with whatever their clock printed.</summary>
    public static bool IsW3cdtf(string value) =>
        DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out _);

    /// <summary>The other way; null when the date is not W3CDTF.</summary>
    public static string? ToPdfDate(string w3cdtf)
    {
        if (!DateTimeOffset.TryParse(w3cdtf, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var when))
        {
            return null;
        }

        var stamp = when.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
        if (when.Offset == TimeSpan.Zero) return $"D:{stamp}Z";
        var sign = when.Offset < TimeSpan.Zero ? '-' : '+';
        var offset = when.Offset.Duration();
        return $"D:{stamp}{sign}{offset.Hours:00}'{offset.Minutes:00}'";
    }
}
