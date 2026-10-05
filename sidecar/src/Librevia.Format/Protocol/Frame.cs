namespace Librevia.Format.Protocol;

/// <summary>
/// A protocol frame: JSON plus a block of raw bytes. The other side is
/// <c>src/main/sidecar/protocol.ts</c>: they change together.
/// <code>
///   offset 0   uint32 BE   JSON byte count
///   offset 4   uint32 BE   binary byte count
///   offset 8   ...         UTF-8 JSON
///   then       ...         raw binary
/// </code>
/// The binary goes outside the JSON because base64 would cost a third more.
/// </summary>
public readonly record struct Frame(ReadOnlyMemory<byte> Json, ReadOnlyMemory<byte> Binary)
{
    public const int HeaderBytes = 8;

    /// <summary>
    /// The TypeScript side's caps, against a header asking for gigabytes. The JSON one is generous
    /// because DOCX images travel in it as data URIs.
    /// </summary>
    public const int MaxJsonBytes = 64 * 1024 * 1024;
    public const int MaxBinaryBytes = 64 * 1024 * 1024;
}

public static class FrameIo
{
    /// <summary>Null when main closes stdin, which is the request to shut down.</summary>
    public static async Task<Frame?> ReadAsync(Stream input, CancellationToken cancellation)
    {
        var header = new byte[Frame.HeaderBytes];
        if (!await ReadExactlyOrEofAsync(input, header, cancellation).ConfigureAwait(false))
        {
            return null;
        }

        var jsonLength = ReadUInt32BigEndian(header, 0);
        var binaryLength = ReadUInt32BigEndian(header, 4);

        if (jsonLength > Frame.MaxJsonBytes || binaryLength > Frame.MaxBinaryBytes)
        {
            throw new InvalidDataException(
                $"quadro anuncia {jsonLength} bytes de JSON e {binaryLength} de binário");
        }

        var json = new byte[jsonLength];
        if (!await ReadExactlyOrEofAsync(input, json, cancellation).ConfigureAwait(false))
        {
            throw new InvalidDataException("fluxo terminou no meio do JSON");
        }

        var binary = new byte[binaryLength];
        if (!await ReadExactlyOrEofAsync(input, binary, cancellation).ConfigureAwait(false))
        {
            throw new InvalidDataException("fluxo terminou no meio do binário");
        }

        return new Frame(json, binary);
    }

    public static async Task WriteAsync(
        Stream output,
        ReadOnlyMemory<byte> json,
        ReadOnlyMemory<byte> binary,
        CancellationToken cancellation)
    {
        var header = new byte[Frame.HeaderBytes];
        WriteUInt32BigEndian(header, 0, (uint)json.Length);
        WriteUInt32BigEndian(header, 4, (uint)binary.Length);

        await output.WriteAsync(header, cancellation).ConfigureAwait(false);
        await output.WriteAsync(json, cancellation).ConfigureAwait(false);
        if (!binary.IsEmpty)
        {
            await output.WriteAsync(binary, cancellation).ConfigureAwait(false);
        }

        // Without the flush, main waits for a response stuck in the buffer.
        await output.FlushAsync(cancellation).ConfigureAwait(false);
    }

    /// <summary>A pipe delivers as much as it wants on each read.</summary>
    private static async Task<bool> ReadExactlyOrEofAsync(
        Stream input,
        Memory<byte> destination,
        CancellationToken cancellation)
    {
        var filled = 0;
        while (filled < destination.Length)
        {
            var read = await input
                .ReadAsync(destination[filled..], cancellation)
                .ConfigureAwait(false);

            if (read == 0)
            {
                return false;
            }

            filled += read;
        }

        return true;
    }

    private static uint ReadUInt32BigEndian(ReadOnlySpan<byte> source, int offset) =>
        System.Buffers.Binary.BinaryPrimitives.ReadUInt32BigEndian(source[offset..]);

    private static void WriteUInt32BigEndian(Span<byte> destination, int offset, uint value) =>
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(destination[offset..], value);
}
