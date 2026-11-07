using System.IO.Compression;

namespace Librevia.Format.Odt;

/// <summary>An entry a save adds to the package, dated now.</summary>
internal readonly record struct OdtAddedEntry(string Path, byte[] Bytes, CompressionLevel Level);

/// <summary>
/// The package written again: <c>mimetype</c> first, stored, so a reader finds the type at a fixed
/// offset; the rest in the original order, with the original dates.
/// </summary>
internal static class OdtZip
{
    /// <param name="bytesOf">An entry's new content; <c>null</c> leaves it out.</param>
    public static byte[] Write(OdtPackage package, Func<string, byte[]?> bytesOf, IEnumerable<OdtAddedEntry> added)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entries = package.Entries.OrderBy(entry => entry.FullName == OdtPackage.MimeTypePath ? 0 : 1);
            foreach (var entry in entries)
            {
                if (bytesOf(entry.FullName) is not { } bytes) continue;
                var level = entry.FullName == OdtPackage.MimeTypePath ? CompressionLevel.NoCompression : CompressionLevel.Optimal;
                Add(zip, entry.FullName, bytes, level, entry.LastWriteTime);
            }

            foreach (var (path, bytes, level) in added) Add(zip, path, bytes, level, DateTimeOffset.Now);
        }

        return output.ToArray();
    }

    private static void Add(ZipArchive zip, string path, byte[] bytes, CompressionLevel level, DateTimeOffset written)
    {
        var entry = zip.CreateEntry(path, level);
        entry.LastWriteTime = written;
        using var target = entry.Open();
        target.Write(bytes);
    }
}
