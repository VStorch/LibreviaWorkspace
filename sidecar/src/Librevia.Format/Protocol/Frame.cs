namespace Librevia.Format.Protocol;

/// <summary>
/// Um quadro do protocolo: JSON mais um bloco de bytes crus.
/// </summary>
/// <remarks>
/// O formato está descrito em <c>src/main/sidecar/protocol.ts</c>, que é o outro
/// lado desta mesma conversa. Os dois arquivos precisam mudar juntos.
///
/// <code>
///   offset 0   uint32 BE   bytes de JSON
///   offset 4   uint32 BE   bytes de binário
///   offset 8   ...         JSON em UTF-8
///   depois     ...         binário cru
/// </code>
///
/// O binário viaja fora do JSON porque base64 custaria 33% a mais sobre
/// documentos de até 20 MB — e isso apareceria no tempo de abrir cada arquivo.
/// </remarks>
public readonly record struct Frame(ReadOnlyMemory<byte> Json, ReadOnlyMemory<byte> Binary)
{
    public const int HeaderBytes = 8;

    /// <summary>
    /// Tetos de sanidade, espelhando os do lado TypeScript. Protegem contra um
    /// cabeçalho mentiroso nos fazer alocar gigabytes.
    /// </summary>
    public const int MaxJsonBytes = 8 * 1024 * 1024;
    public const int MaxBinaryBytes = 64 * 1024 * 1024;
}

public static class FrameIo
{
    /// <summary>
    /// Lê um quadro inteiro, ou devolve <c>null</c> quando o fluxo termina de
    /// forma limpa (o main fechou o stdin — é assim que pedimos para encerrar).
    /// </summary>
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

        // Sem o flush o quadro pode ficar preso no buffer e o main espera para
        // sempre por uma resposta que já foi escrita.
        await output.FlushAsync(cancellation).ConfigureAwait(false);
    }

    /// <summary>
    /// Preenche <paramref name="destination"/> por completo. Um pipe entrega o
    /// que quiser em cada leitura: assumir que uma leitura traz a mensagem
    /// inteira é o bug clássico desta integração.
    /// </summary>
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
