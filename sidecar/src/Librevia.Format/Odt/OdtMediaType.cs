using System.Text;

namespace Librevia.Format.Odt;

/// <summary>
/// The package's media type, in <c>mimetype</c> and in the manifest's root entry: a template
/// (<c>.ott</c>) differs from a document only there.
/// </summary>
internal static class OdtMediaType
{
    public const string RootPath = "/";

    /// <summary>The type to write, or <c>null</c> when the package already declares it.</summary>
    public static string? Wanted(OdtPackage package, bool template)
    {
        var current = package.MimeType;
        var wanted = template ? OdtPackage.TemplateMimeType
            : current == OdtPackage.TemplateMimeType ? OdtPackage.TextMimeType
            : current;
        return wanted == current ? null : wanted;
    }

    public static byte[] Bytes(string mediaType) => Encoding.ASCII.GetBytes(mediaType);

    /// <summary>The root entry's tag with its <c>media-type</c> value swapped, quotes kept.</summary>
    public static string Retyped(string tag, string? from, string to)
    {
        if (from is null) return tag;
        return tag.Replace($"\"{from}\"", $"\"{to}\"", StringComparison.Ordinal)
            .Replace($"'{from}'", $"'{to}'", StringComparison.Ordinal);
    }
}
