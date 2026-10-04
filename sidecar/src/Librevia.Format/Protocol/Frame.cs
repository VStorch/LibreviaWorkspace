namespace Librevia.Format.Protocol;

/// <summary>
/// Um quadro do protocolo: JSON mais um bloco de bytes crus. O outro lado é
/// <c>src/main/sidecar/protocol.ts</c>: mudam juntos.
/// <code>
///   offset 0   uint32 BE   bytes de JSON
///   offset 4   uint32 BE   bytes de binário
///   offset 8   ...         JSON em UTF-8
///   depois     ...         binário cru
/// </code>
/// O binário vai fora do JSON porque base64 custaria um terço a mais.
/// </summary>
public readonly record struct Frame(ReadOnlyMemory<byte> Json, ReadOnlyMemory<byte> Binary)
{
    public const int HeaderBytes = 8;

    /// <summary>
    /// Os tetos do lado TypeScript, contra um cabeçalho que peça gigabytes. O do JSON
    /// é largo porque as imagens do DOCX vão nele como data URI.
    /// </summary>
    public const int MaxJsonBytes = 64 * 1024 * 1024;
    public const int MaxBinaryBytes = 64 * 1024 * 1024;
}

public static class FrameIo
{
    /// <summary>Nulo quando o main fecha o stdin, que é o pedido de encerrar.</summary>
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

        // Sem o flush, o main espera por uma resposta presa no buffer.
        await output.FlushAsync(cancellation).ConfigureAwait(false);
    }

    /// <summary>Um pipe entrega quanto quiser a cada leitura.</summary>
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
