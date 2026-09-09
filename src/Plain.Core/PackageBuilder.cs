using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace Plain.Core;

/// <summary>
/// Writes a package from nothing. <see cref="OpcPackage"/> only ever rewrites a file that already existed, which is
/// the whole point of it; making a new file needs the other half, so it lives here and is used for exactly one thing.
///
/// The timestamp on every entry is fixed rather than "now", so two blank files made a week apart are the same bytes.
/// That makes a new file something a test can check, and it keeps a version control system quiet.
/// </summary>
public static class PackageBuilder
{
    private const ushort DosDate = ((2026 - 1980) << 9) | (1 << 5) | 1;   // 1 January 2026
    private const ushort DosTime = 0;

    public static byte[] Build(IReadOnlyList<(string Name, byte[] Content)> parts)
    {
        using var ms = new MemoryStream();
        var offsets = new long[parts.Count];
        var payloads = new byte[parts.Count][];
        var methods = new ushort[parts.Count];
        var crcs = new uint[parts.Count];

        for (int i = 0; i < parts.Count; i++)
        {
            var (name, content) = parts[i];
            var nameBytes = Encoding.UTF8.GetBytes(name);
            var deflated = Deflate(content);
            // Storing beats deflating when deflating makes it bigger, which happens with very small parts.
            payloads[i] = deflated.Length < content.Length ? deflated : content;
            methods[i] = deflated.Length < content.Length ? (ushort)8 : (ushort)0;
            crcs[i] = Crc32.Of(content);
            offsets[i] = ms.Position;

            var header = new byte[30];
            BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(0), 0x04034b50);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(4), 20);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(6), 0);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(8), methods[i]);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(10), DosTime);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(12), DosDate);
            BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(14), crcs[i]);
            BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(18), (uint)payloads[i].Length);
            BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(22), (uint)content.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(26), (ushort)nameBytes.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(28), 0);
            ms.Write(header);
            ms.Write(nameBytes);
            ms.Write(payloads[i]);
        }

        long directoryStart = ms.Position;
        for (int i = 0; i < parts.Count; i++)
        {
            var nameBytes = Encoding.UTF8.GetBytes(parts[i].Name);
            var record = new byte[46];
            BinaryPrimitives.WriteUInt32LittleEndian(record.AsSpan(0), 0x02014b50);
            BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(4), 20);
            BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(6), 20);
            BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(8), 0);
            BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(10), methods[i]);
            BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(12), DosTime);
            BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(14), DosDate);
            BinaryPrimitives.WriteUInt32LittleEndian(record.AsSpan(16), crcs[i]);
            BinaryPrimitives.WriteUInt32LittleEndian(record.AsSpan(20), (uint)payloads[i].Length);
            BinaryPrimitives.WriteUInt32LittleEndian(record.AsSpan(24), (uint)parts[i].Content.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(28), (ushort)nameBytes.Length);
            BinaryPrimitives.WriteUInt32LittleEndian(record.AsSpan(42), (uint)offsets[i]);
            ms.Write(record);
            ms.Write(nameBytes);
        }
        long directoryEnd = ms.Position;

        var end = new byte[22];
        BinaryPrimitives.WriteUInt32LittleEndian(end.AsSpan(0), 0x06054b50);
        BinaryPrimitives.WriteUInt16LittleEndian(end.AsSpan(8), (ushort)parts.Count);
        BinaryPrimitives.WriteUInt16LittleEndian(end.AsSpan(10), (ushort)parts.Count);
        BinaryPrimitives.WriteUInt32LittleEndian(end.AsSpan(12), (uint)(directoryEnd - directoryStart));
        BinaryPrimitives.WriteUInt32LittleEndian(end.AsSpan(16), (uint)directoryStart);
        ms.Write(end);

        return ms.ToArray();
    }

    private static byte[] Deflate(byte[] content)
    {
        using var output = new MemoryStream(content.Length);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) deflate.Write(content);
        return output.ToArray();
    }
}
