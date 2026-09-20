using System.Buffers.Binary;

namespace Librevia.Format.Docx;

/// <summary>
/// O tamanho real de uma imagem, lido do cabeçalho dos próprios bytes.
/// </summary>
/// <remarks>
/// Existe porque a imagem que a pessoa insere pela barra de ferramentas chega
/// sem medida nenhuma: o editor a mostra no tamanho natural e o `.docx` precisa
/// de um `wp:extent` em EMU. Na falta dele o escritor chutava 600 × 450 — uma
/// captura de tela quadrada saía deitada, e um ícone de 32 px ocupava meia
/// folha.
///
/// Quatro formatos, quinze linhas, nenhuma dependência: o cabeçalho de PNG, GIF,
/// BMP e JPEG diz a medida nos primeiros bytes, e é só isso que queremos saber.
/// Decodificar a imagem para descobrir o tamanho dela seria trazer uma
/// biblioteca de imagem para dentro de um programa que nunca desenha nada.
/// </remarks>
internal static class ImageSize
{
    /// <summary>Largura e altura em pixels, ou <c>null</c> se o cabeçalho não disser.</summary>
    public static (int Width, int Height)? Of(byte[] bytes)
    {
        var span = bytes.AsSpan();

        if (span.Length > 24 && span[..8].SequenceEqual((byte[])[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]))
        {
            // `IHDR` é sempre o primeiro pedaço, e a medida são dois inteiros de
            // 32 bits em ordem de rede.
            return (
                (int)BinaryPrimitives.ReadUInt32BigEndian(span[16..20]),
                (int)BinaryPrimitives.ReadUInt32BigEndian(span[20..24]));
        }

        if (span.Length > 10 && (span[..6].SequenceEqual("GIF87a"u8) || span[..6].SequenceEqual("GIF89a"u8)))
        {
            // O GIF é o único dos quatro que guarda a medida em ordem de byte
            // trocada, porque nasceu no mundo dos processadores da Intel.
            return (
                BinaryPrimitives.ReadUInt16LittleEndian(span[6..8]),
                BinaryPrimitives.ReadUInt16LittleEndian(span[8..10]));
        }

        if (span.Length > 26 && span[0] == 0x42 && span[1] == 0x4D)
        {
            // A altura do BMP pode ser negativa: é assim que ele diz que as
            // linhas vêm de cima para baixo. O sinal não é tamanho.
            return (
                BinaryPrimitives.ReadInt32LittleEndian(span[18..22]),
                Math.Abs(BinaryPrimitives.ReadInt32LittleEndian(span[22..26])));
        }

        return span.Length > 3 && span[0] == 0xFF && span[1] == 0xD8 ? Jpeg(span) : null;
    }

    /// <summary>
    /// O JPEG não tem cabeçalho fixo: é uma fila de segmentos, e a medida está
    /// no que começa a moldura (`SOFn`).
    /// </summary>
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

            // Todos os `SOF` guardam altura e largura nos mesmos lugares. Os
            // quatro fora da faixa (`DHT`, `JPG`, `DAC`, `RSTn`) não são moldura.
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
