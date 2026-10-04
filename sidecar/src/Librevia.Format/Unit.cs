namespace Librevia.Format;

/// <summary>
/// O OOXML mede em twips (vigésimos de ponto), meios-pontos e EMUs; a tela, em
/// pixels do CSS, 96 por polegada.
/// </summary>
public static class Unit
{
    public const double MillimetersPerInch = 25.4;
    public const double CentimetersPerInch = 2.54;
    public const int PointsPerInch = 72;
    public const int PixelsPerInch = 96;
    public const int TwipsPerInch = 1440;
    public const long EmusPerInch = 914400;

    public const int TwipsPerPoint = 20;
    public const int TwipsPerPixel = TwipsPerInch / PixelsPerInch;
    public const int HalfPointsPerPoint = 2;
    public const long EmusPerPoint = EmusPerInch / PointsPerInch;
    public const long EmusPerPixel = EmusPerInch / PixelsPerInch;

    /// <summary>Meia polegada: o passo de recuo do Word.</summary>
    public const int IndentStepTwips = TwipsPerInch / 2;

    public static double TwipsToPoints(double twips) => twips / TwipsPerPoint;

    public static double TwipsToMillimeters(double twips) => twips * MillimetersPerInch / TwipsPerInch;

    public static double HalfPointsToPoints(double halfPoints) => halfPoints / HalfPointsPerPoint;

    public static double EmusToPixels(double emus) => emus / EmusPerPixel;
}
