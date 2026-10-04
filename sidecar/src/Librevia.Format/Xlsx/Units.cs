namespace Librevia.Format.Xlsx;

/// <summary>
/// A coluna do XLSX conta caracteres da fonte padrão, a linha conta pontos, e a
/// tela, pixels.
/// </summary>
public static class Units
{
    /// <summary>O dígito mais largo do Calibri 11, a fonte padrão desde 2007, em pixels.</summary>
    private const double MaxDigitWidth = 7.0;

    /// <summary>Caracteres → pixels. A largura padrão 8,43 dá 64 pixels.</summary>
    public static double WidthToPixels(double width) =>
        Math.Round(width * MaxDigitWidth + 5.0);

    public static double PixelsToWidth(double pixels) =>
        Math.Max(0, (pixels - 5.0) / MaxDigitWidth);

    /// <summary>Pontos → pixels, na densidade de 96 por polegada do CSS.</summary>
    public static double PointsToPixels(double points) => Math.Round(points * 96.0 / 72.0, 2);

    public static double PixelsToPoints(double pixels) => Math.Round(pixels * 72.0 / 96.0, 2);
}
