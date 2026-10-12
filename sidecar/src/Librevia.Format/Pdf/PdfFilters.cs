using System.IO.Compression;

namespace Librevia.Format.Pdf;

/// <summary>
/// Only Flate, the filter of cross-reference streams, object streams and XMP. The page contents are
/// never decoded here.
/// </summary>
internal static class PdfFilters
{
    /// <summary>Against a few kilobytes that inflate to gigabytes.</summary>
    private const int MaxDecodedBytes = 256 * 1024 * 1024;

    public static byte[] Decode(PdfStream stream, Func<object?, object?> resolve)
    {
        var filters = Names(resolve(stream.Dictionary["Filter"]));
        if (filters.Count == 0) return stream.Data;
        if (filters.Count > 1 || filters[0] != "FlateDecode")
        {
            throw new PdfException("Este PDF usa uma compactação que o Librevia não sabe ler.");
        }

        var parameters = resolve(stream.Dictionary["DecodeParms"]) switch
        {
            PdfDictionary single => single,
            PdfArray { Count: > 0 } list => resolve(list[0]) as PdfDictionary,
            _ => null,
        };
        return Unpredict(Inflate(stream.Data), parameters);
    }

    private static List<string> Names(object? filter) => filter switch
    {
        PdfName name => [name.Value],
        PdfArray list => [.. list.OfType<PdfName>().Select(name => name.Value)],
        _ => [],
    };

    private static byte[] Inflate(byte[] data)
    {
        try
        {
            return Copy(new ZLibStream(new MemoryStream(data), CompressionMode.Decompress));
        }
        catch (InvalidDataException)
        {
            // Some writers forget the zlib header: the raw deflate data follows anyway.
            if (data.Length < 2) throw PdfParser.Broken();
            return Copy(new DeflateStream(new MemoryStream(data, 2, data.Length - 2), CompressionMode.Decompress));
        }
    }

    private static byte[] Copy(Stream source)
    {
        using (source)
        {
            var output = new MemoryStream();
            var buffer = new byte[81920];
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                if (output.Length + read > MaxDecodedBytes) throw PdfParser.Broken();
                output.Write(buffer, 0, read);
            }

            return output.ToArray();
        }
    }

    /// <summary>The PNG predictors, which cross-reference streams use to shrink offsets.</summary>
    private static byte[] Unpredict(byte[] data, PdfDictionary? parameters)
    {
        var predictor = Integer(parameters?["Predictor"]) ?? 1;
        if (predictor == 1) return data;
        if (predictor < 10) throw new PdfException("Este PDF usa uma compactação que o Librevia não sabe ler.");

        var colors = Integer(parameters?["Colors"]) ?? 1;
        var bits = Integer(parameters?["BitsPerComponent"]) ?? 8;
        var columns = Integer(parameters?["Columns"]) ?? 1;
        var pixel = Math.Max(1, (int)(colors * bits / 8));
        var row = (int)((columns * colors * bits) + 7) / 8;
        if (row <= 0) throw PdfParser.Broken();

        var output = new MemoryStream();
        var previous = new byte[row];
        for (var start = 0; start + 1 + row <= data.Length; start += row + 1)
        {
            var current = data.AsSpan(start + 1, row).ToArray();
            Reverse(data[start], current, previous, pixel);
            output.Write(current);
            previous = current;
        }

        return output.ToArray();
    }

    private static void Reverse(byte type, byte[] current, byte[] previous, int pixel)
    {
        for (var index = 0; index < current.Length; index++)
        {
            var left = index >= pixel ? current[index - pixel] : 0;
            var up = previous[index];
            var upLeft = index >= pixel ? previous[index - pixel] : 0;
            current[index] = (byte)(current[index] + type switch
            {
                1 => left,
                2 => up,
                3 => (left + up) / 2,
                4 => Paeth(left, up, upLeft),
                _ => 0,
            });
        }
    }

    private static int Paeth(int left, int up, int upLeft)
    {
        var estimate = left + up - upLeft;
        var toLeft = Math.Abs(estimate - left);
        var toUp = Math.Abs(estimate - up);
        var toUpLeft = Math.Abs(estimate - upLeft);
        if (toLeft <= toUp && toLeft <= toUpLeft) return left;
        return toUp <= toUpLeft ? up : upLeft;
    }

    private static long? Integer(object? value) => (value as PdfNumber)?.Integer;
}
