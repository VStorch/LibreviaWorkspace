using System.Globalization;

namespace Librevia.Format.Docx;

internal static class InvariantText
{
    /// <summary>Without any locale's comma: works for twips, half-points and 240ths.</summary>
    public static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);
}
