using System.Buffers.Binary;

namespace Librevia.Format.Docx;

/// <summary>
/// An image's size from its header bytes, for the <c>wp:extent</c> of an inserted image, which
/// arrives without a measure. PNG, GIF, BMP and JPEG, without a library.
/// </summary>
internal static class ImageSize
{
    /// <summary>Width and height in pixels, or <c>null</c> if the header does not say.</summary>
    public static (int Width, int Height)? Of(byte[] bytes)
    {
        var span = bytes.AsSpan();

        if (span.Length > 24 && span[..8].SequenceEqual((byte[])[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]))
        {
            // `IHDR` is always the first chunk: two 32-bit integers in network order.
            return (
                (int)BinaryPrimitives.ReadUInt32BigEndian(span[16..20]),
                (int)BinaryPrimitives.ReadUInt32BigEndian(span[20..24]));
        }

        if (span.Length > 10 && (span[..6].SequenceEqual("GIF87a"u8) || span[..6].SequenceEqual("GIF89a"u8)))
        {
            // GIF stores the size little-endian.
            return (
                BinaryPrimitives.ReadUInt16LittleEndian(span[6..8]),
                BinaryPrimitives.ReadUInt16LittleEndian(span[8..10]));
        }

        if (span.Length > 26 && span[0] == 0x42 && span[1] == 0x4D)
        {
            // A negative height says the rows come top-down.
            return (
                BinaryPrimitives.ReadInt32LittleEndian(span[18..22]),
                Math.Abs(BinaryPrimitives.ReadInt32LittleEndian(span[22..26])));
        }

        return span.Length > 3 && span[0] == 0xFF && span[1] == 0xD8 ? Jpeg(span) : null;
    }

    /// <summary>JPEG is a queue of segments; the size is in <c>SOFn</c>.</summary>
    private static (int Width, int Height)? Jpeg(ReadOnlySpan<byte> span)
    {
        var index = 2;
        while (index + 9 < span.Length)
        {
            if (span[index] != 0xFF)
            {
                index++;
                continue;
            }

            var marker = span[index + 1];
            var length = BinaryPrimitives.ReadUInt16BigEndian(span[(index + 2)..(index + 4)]);

            // `DHT`, `JPG` and `DAC` are in the `SOF` range, but are not frames.
            if (marker is >= 0xC0 and <= 0xCF && marker is not (0xC4 or 0xC8 or 0xCC))
            {
                return (
                    BinaryPrimitives.ReadUInt16BigEndian(span[(index + 7)..(index + 9)]),
                    BinaryPrimitives.ReadUInt16BigEndian(span[(index + 5)..(index + 7)]));
            }

            if (length < 2) return null;
            index += 2 + length;
        }

        return null;
    }
}
