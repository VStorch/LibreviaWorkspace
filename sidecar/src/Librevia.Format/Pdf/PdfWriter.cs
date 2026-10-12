using System.Globalization;
using System.Text;

namespace Librevia.Format.Pdf;

/// <summary>Objects back to PDF syntax: strings in hex, which needs no escaping.</summary>
internal static class PdfWriter
{
    public static string Write(object value) => value switch
    {
        PdfName name => Name(name.Value),
        PdfString text => "<" + Convert.ToHexString(text.Bytes) + ">",
        PdfNumber number => number.Raw,
        PdfRef reference => $"{reference.Number} {reference.Generation} R",
        bool flag => flag ? "true" : "false",
        PdfArray list => "[" + string.Join(' ', list.Select(Write)) + "]",
        PdfDictionary dictionary => Dictionary(dictionary),
        _ => "null",
    };

    private static string Dictionary(PdfDictionary dictionary)
    {
        var builder = new StringBuilder("<<");
        foreach (var (key, value) in dictionary.Entries) builder.Append(Name(key)).Append(' ').Append(Write(value));
        return builder.Append(">>").ToString();
    }

    private static string Name(string value)
    {
        var builder = new StringBuilder("/");
        foreach (var letter in Encoding.Latin1.GetBytes(value))
        {
            if (letter is > (byte)' ' and < 0x7F and not ((byte)'#' or (byte)'/' or (byte)'(' or (byte)')' or (byte)'<'
                or (byte)'>' or (byte)'[' or (byte)']' or (byte)'{' or (byte)'}' or (byte)'%'))
            {
                builder.Append((char)letter);
            }
            else
            {
                builder.Append('#').Append(letter.ToString("X2", CultureInfo.InvariantCulture));
            }
        }

        return builder.ToString();
    }
}
