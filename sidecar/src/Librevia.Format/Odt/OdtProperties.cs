using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>meta.xml</c> → the document's properties, the fields <see cref="DocumentProperties"/> reads
/// from a DOCX: <c>meta:initial-creator</c> is the author, <c>dc:creator</c> the last to save.
/// </summary>
internal static partial class OdtProperties
{
    private const int SecondsPerMinute = 60;

    /// <summary>The user-defined fields LibreOffice keeps for what OOXML has its own place for.</summary>
    private const string Category = "Category";
    private const string Company = "Company";
    private const string Manager = "Manager";

    public static PropertiesDto? Read(XDocument? meta)
    {
        var office = meta?.Root?.Element(Ns.Office + "meta");
        if (office is null) return null;

        string? Text(XName name) => office.Element(name)?.Value.Trim() is { Length: > 0 } value ? value : null;
        string? UserDefined(string field) => office.Elements(Ns.Meta + "user-defined")
            .FirstOrDefault(element => element.Attr(Ns.Meta + "name") == field)?.Value.Trim() is { Length: > 0 } value
            ? value
            : null;

        var keywords = office.Elements(Ns.Meta + "keyword").Select(keyword => keyword.Value.Trim())
            .Where(keyword => keyword.Length > 0).ToList();
        var properties = new PropertiesDto(
            Title: Text(Ns.Dc + "title"),
            Subject: Text(Ns.Dc + "subject"),
            Creator: Text(Ns.Meta + "initial-creator"),
            Keywords: keywords.Count == 0 ? null : string.Join("; ", keywords),
            Category: UserDefined(Category),
            Description: Text(Ns.Dc + "description"),
            LastModifiedBy: Text(Ns.Dc + "creator"),
            Revision: Text(Ns.Meta + "editing-cycles"),
            Created: Text(Ns.Meta + "creation-date"),
            Modified: Text(Ns.Dc + "date"),
            Company: UserDefined(Company),
            Manager: UserDefined(Manager),
            Language: Text(Ns.Dc + "language"),
            LastPrinted: Text(Ns.Meta + "print-date"),
            TotalTime: MinutesOf(Text(Ns.Meta + "editing-duration")));
        return properties == new PropertiesDto() ? null : properties;
    }

    /// <summary><c>PT1H2M3S</c> → 62.</summary>
    private static int? MinutesOf(string? duration)
    {
        if (duration is null || !Duration().IsMatch(duration)) return null;
        try
        {
            return (int)Math.Round(XmlConvert.ToTimeSpan(duration).TotalSeconds / SecondsPerMinute);
        }
        catch (Exception problem) when (problem is FormatException or OverflowException)
        {
            return null;
        }
    }

    [GeneratedRegex("^-?P[0-9YMDTHS.]+$", RegexOptions.CultureInvariant)]
    private static partial Regex Duration();
}
