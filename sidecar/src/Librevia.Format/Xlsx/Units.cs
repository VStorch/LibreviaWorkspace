namespace Librevia.Format.Xlsx;

/// <summary>
/// As três unidades que a planilha mistura.
/// </summary>
/// <remarks>
/// Largura de coluna no XLSX é contada em **caracteres** — quantos dígitos da
/// fonte padrão cabem —, altura de linha em **pontos**, e a tela trabalha em
/// pixels. Trocar uma pela outra não quebra nada visivelmente: só deixa a
/// coluna com a largura errada, que é o tipo de defeito que ninguém associa a
/// uma conversão de unidade.
/// </remarks>
public static class Units
{
    /// <summary>
    /// Largura do dígito mais largo da fonte padrão, em pixels.
    /// </summary>
    /// <remarks>
    /// Sete pixels é o valor do Calibri 11, que é a fonte padrão de toda
    /// planilha nova desde 2007. O formato define a conta em função dela.
    /// </remarks>
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
