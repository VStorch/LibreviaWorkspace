using System.Buffers.Binary;
using System.Text;

namespace Librevia.Format.Privacy;

/// <summary>
/// What a camera or an editing program wrote inside a JPEG or PNG: EXIF (with the GPS position),
/// XMP, IPTC and text comments. The pixels and the color profile stay. The orientation stays too,
/// in a minimal EXIF: without it a photo taken sideways would turn on screen.
/// </summary>
internal static class ImageMetadata
{
    private const byte Marker = 0xFF;
    private const byte StartOfImage = 0xD8;
    private const byte StartOfScan = 0xDA;
    private const byte App1 = 0xE1;
    private const byte App13 = 0xED;
    private const byte Comment = 0xFE;
    private const ushort OrientationTag = 0x0112;
    private const ushort ShortType = 3;
    private const ushort Upright = 1;

    private static readonly byte[] PngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    private static readonly byte[] ExifHeader = "Exif\0\0"u8.ToArray();
    private static readonly HashSet<string> PngMetadataChunks = new(StringComparer.Ordinal)
    {
        "eXIf", "tEXt", "zTXt", "iTXt", "tIME",
    };

    /// <summary>
    /// The image without metadata; null when it has none (the orientation alone is not
    /// metadata) or is not a JPEG or PNG.
    /// </summary>
    public static byte[]? Strip(byte[] image)
    {
        var clean = IsJpeg(image) ? StripJpeg(image)
            : image.AsSpan().StartsWith(PngSignature) ? StripPng(image)
            : null;
        return clean is null || clean.AsSpan().SequenceEqual(image) ? null : clean;
    }

    private static bool IsJpeg(byte[] image) =>
        image.Length > 3 && image[0] == Marker && image[1] == StartOfImage && image[2] == Marker;

    private static bool IsJpegMetadata(byte marker) => marker is App1 or App13 or Comment;

    /// <summary>Null when malformed: a damaged image is not this inspector's to repair.</summary>
    private static byte[]? StripJpeg(byte[] image)
    {
        var kept = new List<(int Start, int Length)>();
        var orientation = Upright;
        var removed = false;
        var position = 2;

        while (position + 4 <= image.Length && image[position] == Marker)
        {
            var marker = image[position + 1];
            if (marker == StartOfScan) break;

            var length = BinaryPrimitives.ReadUInt16BigEndian(image.AsSpan(position + 2));
            var end = position + 2 + length;
            if (length < 2 || end > image.Length) return null;

            if (IsJpegMetadata(marker))
            {
                if (marker == App1) orientation = Math.Max(orientation, OrientationOf(image.AsSpan(position + 4, length - 2)));
                removed = true;
            }
            else
            {
                kept.Add((position, end - position));
            }

            position = end;
        }

        if (!removed || position >= image.Length) return null;

        using var output = new MemoryStream(image.Length);
        output.Write(image, 0, 2);
        if (orientation != Upright) output.Write(JpegOrientationSegment(orientation));
        foreach (var (start, length) in kept) output.Write(image, start, length);
        output.Write(image, position, image.Length - position);
        return output.ToArray();
    }

    private static byte[] JpegOrientationSegment(ushort orientation)
    {
        var payload = new byte[ExifHeader.Length + MinimalTiffLength];
        ExifHeader.CopyTo(payload, 0);
        WriteMinimalTiff(payload.AsSpan(ExifHeader.Length), orientation);

        var segment = new byte[4 + payload.Length];
        segment[0] = Marker;
        segment[1] = App1;
        BinaryPrimitives.WriteUInt16BigEndian(segment.AsSpan(2), (ushort)(payload.Length + 2));
        payload.CopyTo(segment, 4);
        return segment;
    }

    private static byte[]? StripPng(byte[] image)
    {
        using var output = new MemoryStream(image.Length);
        output.Write(PngSignature);
        var removed = false;
        var position = PngSignature.Length;

        while (position + 12 <= image.Length)
        {
            var length = BinaryPrimitives.ReadInt32BigEndian(image.AsSpan(position));
            var end = position + 12 + length;
            if (length < 0 || end > image.Length) return null;

            var type = Encoding.ASCII.GetString(image, position + 4, 4);
            if (PngMetadataChunks.Contains(type))
            {
                removed = true;
                var orientation = type == "eXIf" ? OrientationOfTiff(image.AsSpan(position + 8, length)) : Upright;
                if (orientation != Upright) output.Write(PngOrientationChunk(orientation));
            }
            else
            {
                output.Write(image, position, end - position);
            }

            position = end;
        }

        return removed && position == image.Length ? output.ToArray() : null;
    }

    private static byte[] PngOrientationChunk(ushort orientation)
    {
        var chunk = new byte[12 + MinimalTiffLength];
        BinaryPrimitives.WriteInt32BigEndian(chunk, MinimalTiffLength);
        "eXIf"u8.CopyTo(chunk.AsSpan(4));
        WriteMinimalTiff(chunk.AsSpan(8, MinimalTiffLength), orientation);
        BinaryPrimitives.WriteUInt32BigEndian(chunk.AsSpan(8 + MinimalTiffLength), Crc32(chunk.AsSpan(4, 4 + MinimalTiffLength)));
        return chunk;
    }

    /// <summary>Header, one IFD with one entry, and the end of the chain.</summary>
    private const int MinimalTiffLength = 8 + 2 + 12 + 4;

    private static void WriteMinimalTiff(Span<byte> tiff, ushort orientation)
    {
        "MM"u8.CopyTo(tiff);
        BinaryPrimitives.WriteUInt16BigEndian(tiff[2..], 42);
        BinaryPrimitives.WriteUInt32BigEndian(tiff[4..], 8);
        BinaryPrimitives.WriteUInt16BigEndian(tiff[8..], 1);
        BinaryPrimitives.WriteUInt16BigEndian(tiff[10..], OrientationTag);
        BinaryPrimitives.WriteUInt16BigEndian(tiff[12..], ShortType);
        BinaryPrimitives.WriteUInt32BigEndian(tiff[14..], 1);
        BinaryPrimitives.WriteUInt16BigEndian(tiff[18..], orientation);
        BinaryPrimitives.WriteUInt16BigEndian(tiff[20..], 0);
        BinaryPrimitives.WriteUInt32BigEndian(tiff[22..], 0);
    }

    private static ushort OrientationOf(ReadOnlySpan<byte> app1) =>
        app1.StartsWith(ExifHeader) ? OrientationOfTiff(app1[ExifHeader.Length..]) : Upright;

    /// <summary>Tag 0x0112 of the first IFD; upright when absent or unreadable.</summary>
    private static ushort OrientationOfTiff(ReadOnlySpan<byte> tiff)
    {
        if (tiff.Length < 8) return Upright;
        var little = tiff[0] == (byte)'I';
        if (!little && tiff[0] != (byte)'M') return Upright;

        uint ReadUInt32(ReadOnlySpan<byte> span) =>
            little ? BinaryPrimitives.ReadUInt32LittleEndian(span) : BinaryPrimitives.ReadUInt32BigEndian(span);
        ushort ReadUInt16(ReadOnlySpan<byte> span) =>
            little ? BinaryPrimitives.ReadUInt16LittleEndian(span) : BinaryPrimitives.ReadUInt16BigEndian(span);

        var ifd = ReadUInt32(tiff[4..]);
        if (ifd > tiff.Length - 2) return Upright;
        var count = ReadUInt16(tiff[(int)ifd..]);

        for (var index = 0; index < count; index++)
        {
            var entry = (int)ifd + 2 + (index * 12);
            if (entry + 12 > tiff.Length) return Upright;
            if (ReadUInt16(tiff[entry..]) != OrientationTag) continue;

            var value = ReadUInt16(tiff[(entry + 8)..]);
            return value is >= 1 and <= 8 ? value : Upright;
        }

        return Upright;
    }

    private static readonly uint[] CrcTable = BuildCrcTable();

    private static uint[] BuildCrcTable()
    {
        var table = new uint[256];
        for (uint index = 0; index < table.Length; index++)
        {
            var value = index;
            for (var bit = 0; bit < 8; bit++) value = (value & 1) != 0 ? 0xEDB88320 ^ (value >> 1) : value >> 1;
            table[index] = value;
        }

        return table;
    }

    private static uint Crc32(ReadOnlySpan<byte> data)
    {
        var crc = 0xFFFFFFFF;
        foreach (var value in data) crc = CrcTable[(crc ^ value) & 0xFF] ^ (crc >> 8);
        return crc ^ 0xFFFFFFFF;
    }
}
