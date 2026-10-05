namespace Librevia.Format.Xlsx;

/// <summary>
/// An XLSX column counts characters of the default font, a row counts points, and the screen,
/// pixels.
/// </summary>
public static class Units
{
    /// <summary>The widest Calibri 11 digit, the default font since 2007, in pixels.</summary>
    private const double MaxDigitWidth = 7.0;

    /// <summary>Characters → pixels. The default width 8.43 gives 64 pixels.</summary>
    public static double WidthToPixels(double width) =>
        Math.Round(width * MaxDigitWidth + 5.0);

    public static double PixelsToWidth(double pixels) =>
        Math.Max(0, (pixels - 5.0) / MaxDigitWidth);

    public static double PointsToPixels(double points) =>
        Math.Round(points * Unit.PixelsPerInch / Unit.PointsPerInch, 2);

    public static double PixelsToPoints(double pixels) =>
        Math.Round(pixels * Unit.PointsPerInch / Unit.PixelsPerInch, 2);
}
