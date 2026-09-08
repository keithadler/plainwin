using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace Plain.Core;

/// <summary>
/// A Word, Excel or PowerPoint file, read as the ZIP package it really is.
///
/// The whole promise of Plain rests here. Every part is held as the exact bytes it occupies in the original file,
/// compression and all, and a part nobody edited is written back as that same byte range. An edited part is the only
/// thing rebuilt. Save a file you did not change and you get the same file back, byte for byte, which is a claim a
/// test can check rather than a claim a README can make.
/// </summary>
public sealed class OpcPackage
{
    private const uint SigLocal = 0x04034b50;
    private const uint SigCentral = 0x02014b50;
    private const uint SigEocd = 0x06054b50;
    private const uint SigZip64Eocd = 0x06064b50;
    private const uint SigZip64Locator = 0x07064b50;
    private const uint SigDescriptor = 0x08074b50;

    private readonly byte[] _raw;
    private readonly List<Part> _parts = new();
    private readonly Dictionary<string, Part> _byName = new(StringComparer.Ordinal);
    private int _eocdOffset;
    private int _eocdLength;

    /// <summary>True when the container uses ZIP64 records. Plain reads and re-saves these, but will not edit them yet.</summary>
    public bool Zip64Container { get; private init; }

    public string? SourcePath { get; private set; }

    /// <summary>Every part in the file, in the order the file stores them.</summary>
    public IReadOnlyList<Part> Parts => _parts;

    public sealed class Part
    {
        public required string Name { get; init; }
        public required int CentralOffset { get; init; }
        public required int CentralLength { get; init; }
        public required long LocalOffset { get; init; }
        public required int LocalLength { get; init; }
        public required long CompressedSize { get; init; }
        public required long UncompressedSize { get; init; }
        public required ushort Method { get; init; }
        public required ushort Flags { get; init; }
        public required uint Crc { get; init; }
        public required int DataOffset { get; init; }
        public required bool Zip64 { get; init; }

        internal byte[]? Replacement;

        /// <summary>True once something in this session replaced the part's content.</summary>
        public bool Edited => Replacement is not null;

        /// <summary>The size Plain will write, which differs from <see cref="UncompressedSize"/> only after an edit.</summary>
        public long Size => Replacement?.LongLength ?? UncompressedSize;

        public override string ToString() => Name;
    }

    public sealed class PackageException : Exception
    {
        public PackageException(string message) : base(message) { }
    }

    private OpcPackage(byte[] raw, bool zip64) { _raw = raw; Zip64Container = zip64; }

    // ---------- reading ----------

    public static OpcPackage Open(string path)
    {
        var pkg = Read(File.ReadAllBytes(path));
        pkg.SourcePath = path;
        return pkg;
    }

    public static OpcPackage Read(byte[] raw)
    {
        if (raw.Length < 22) throw new PackageException("Not a Word, Excel or PowerPoint file: too small to be a package.");

        int eocd = FindEocd(raw);
        if (eocd < 0) throw new PackageException("Not a Word, Excel or PowerPoint file: no ZIP end record found.");

        int commentLen = U16(raw, eocd + 20);
        long entryCount = U16(raw, eocd + 10);
        long cdSize = U32(raw, eocd + 12);
        long cdOffset = U32(raw, eocd + 16);
        bool zip64 = false;

        // ZIP64: the 32-bit fields are saturated and the real values live in the ZIP64 end record.
        if (entryCount == 0xFFFF || cdSize == 0xFFFFFFFF || cdOffset == 0xFFFFFFFF)
        {
            int loc = eocd - 20;
            if (loc >= 0 && U32(raw, loc) == SigZip64Locator)
            {
                long z64 = (long)U64(raw, loc + 8);
                if (z64 >= 0 && z64 + 56 <= raw.Length && U32(raw, (int)z64) == SigZip64Eocd)
                {
                    zip64 = true;
                    entryCount = (long)U64(raw, (int)z64 + 32);
                    cdSize = (long)U64(raw, (int)z64 + 40);
                    cdOffset = (long)U64(raw, (int)z64 + 48);
                }
            }
            if (!zip64) throw new PackageException("This package uses ZIP64 records Plain could not read.");
        }

        // A damaged directory can claim any number of parts; a real Office file has hundreds, not millions.
        if (entryCount < 0 || entryCount > 200_000)
            throw new PackageException("This file's directory claims an impossible number of parts; it is damaged.");
        if (cdOffset < 0 || cdOffset > raw.Length)
            throw new PackageException("This file's directory is not where the file says it is; it is damaged.");

        var pkg = new OpcPackage(raw, zip64) { _eocdOffset = eocd, _eocdLength = 22 + commentLen };

        int p = (int)cdOffset;
        for (long i = 0; i < entryCount; i++)
        {
            if (p + 46 > raw.Length || U32(raw, p) != SigCentral)
                throw new PackageException("The package's directory is damaged; Plain will not guess at it.");

            ushort flags = U16(raw, p + 8);
            ushort method = U16(raw, p + 10);
            uint crc = U32(raw, p + 16);
            long compSize = U32(raw, p + 20);
            long uncompSize = U32(raw, p + 24);
            int nameLen = U16(raw, p + 28);
            int extraLen = U16(raw, p + 30);
            int cmtLen = U16(raw, p + 32);
            long localOffset = U32(raw, p + 42);

            // The record says how long its own name, extra field and comment are; a damaged one can say more than
            // the file holds, and reading the name would then run off the end.
            if ((long)p + 46 + nameLen + extraLen + cmtLen > raw.Length)
                throw new PackageException("This file's directory runs past the end of the file; it is damaged.");

            string name = Encoding.UTF8.GetString(raw, p + 46, nameLen);

            bool entryZip64 = false;
            if (compSize == 0xFFFFFFFF || uncompSize == 0xFFFFFFFF || localOffset == 0xFFFFFFFF)
            {
                entryZip64 = true;
                ReadZip64Extra(raw, p + 46 + nameLen, extraLen, ref uncompSize, ref compSize, ref localOffset);
            }

            if ((flags & 0x1) != 0) throw new PackageException($"The part \"{name}\" is encrypted; Plain will not open password-protected files.");

            // Walk into the local header, because its name and extra fields can differ in length from the directory's.
            if (localOffset < 0 || localOffset + 30 > raw.Length || U32(raw, (int)localOffset) != SigLocal)
                throw new PackageException($"The package points at \"{name}\" in a place it is not; Plain will not guess at it.");
            int lNameLen = U16(raw, (int)localOffset + 26);
            int lExtraLen = U16(raw, (int)localOffset + 28);
            int dataOffset = (int)localOffset + 30 + lNameLen + lExtraLen;
            long end = dataOffset + compSize;
            if (end > raw.Length) throw new PackageException($"The part \"{name}\" runs past the end of the file.");

            // A data descriptor trails the data when bit 3 is set; its length depends on the signature and on ZIP64.
            if ((flags & 0x8) != 0)
            {
                int fieldWidth = entryZip64 ? 8 : 4;
                if (end + 4 <= raw.Length && U32(raw, (int)end) == SigDescriptor) end += 4;
                end += 4 + fieldWidth * 2;
                if (end > raw.Length) throw new PackageException($"The part \"{name}\" has a trailing record that runs past the end of the file.");
            }

            var part = new Part
            {
                Name = name,
                CentralOffset = p,
                CentralLength = 46 + nameLen + extraLen + cmtLen,
                LocalOffset = localOffset,
                LocalLength = (int)(end - localOffset),
                CompressedSize = compSize,
                UncompressedSize = uncompSize,
                Method = method,
                Flags = flags,
                Crc = crc,
                DataOffset = dataOffset,
                Zip64 = entryZip64,
            };
            pkg._parts.Add(part);
            pkg._byName[name] = part;
            p += part.CentralLength;
        }

        return pkg;
    }

    private static void ReadZip64Extra(byte[] raw, int start, int len, ref long uncomp, ref long comp, ref long localOffset)
    {
        int q = start, stop = start + len;
        while (q + 4 <= stop)
        {
            int id = U16(raw, q), size = U16(raw, q + 2);
            if (id == 0x0001)
            {
                int f = q + 4, endField = Math.Min(q + 4 + size, stop);
                if (uncomp == 0xFFFFFFFF && f + 8 <= endField) { uncomp = (long)U64(raw, f); f += 8; }
                if (comp == 0xFFFFFFFF && f + 8 <= endField) { comp = (long)U64(raw, f); f += 8; }
                if (localOffset == 0xFFFFFFFF && f + 8 <= endField) { localOffset = (long)U64(raw, f); }
                return;
            }
            q += 4 + size;
        }
        throw new PackageException("A part claims ZIP64 sizes but carries no ZIP64 record.");
    }

    private static int FindEocd(byte[] raw)
    {
        int max = Math.Min(raw.Length, 0xFFFF + 22);
        for (int i = raw.Length - 22; i >= raw.Length - max && i >= 0; i--)
            if (U32(raw, i) == SigEocd && i + 22 + U16(raw, i + 20) == raw.Length) return i;
        // Fall back to any end record, for files that carry trailing bytes after the comment.
        for (int i = raw.Length - 22; i >= raw.Length - max && i >= 0; i--)
            if (U32(raw, i) == SigEocd) return i;
        return -1;
    }

    // ---------- part access ----------

    public bool Has(string name) => _byName.ContainsKey(Normalize(name));

    public Part? Find(string name) => _byName.TryGetValue(Normalize(name), out var p) ? p : null;

    /// <summary>The part's content, decompressed. Throws when the part is not in the file.</summary>
    public byte[] Read(string name)
    {
        var part = Find(name) ?? throw new PackageException($"This file has no part called \"{name}\".");
        if (part.Replacement is not null) return (byte[])part.Replacement.Clone();
        return Inflate(part);
    }

    public string ReadText(string name) => DecodeUtf8(Read(name));

    /// <summary>Replace a part's content. Every other part stays exactly as it was found.</summary>
    public void Write(string name, byte[] content)
    {
        var part = Find(name) ?? throw new PackageException($"This file has no part called \"{name}\"; Plain does not add parts.");
        if (part.Zip64) throw new PackageException($"The part \"{name}\" uses ZIP64 records; Plain will not edit those yet.");
        part.Replacement = content;
    }

    public void WriteText(string name, string content) => Write(name, new UTF8Encoding(false).GetBytes(content));

    /// <summary>No single part may unpack to more than this. A damaged or hostile file cannot exhaust the machine.</summary>
    private const long MaxPartSize = 512L * 1024 * 1024;

    private byte[] Inflate(Part part)
    {
        if (part.Method == 0)
        {
            var stored = new byte[part.CompressedSize];
            Array.Copy(_raw, part.DataOffset, stored, 0, stored.Length);
            return stored;
        }
        if (part.Method != 8) throw new PackageException($"The part \"{part.Name}\" uses a compression method Plain does not read ({part.Method}).");

        try
        {
            using var input = new MemoryStream(_raw, part.DataOffset, (int)part.CompressedSize, writable: false);
            using var deflate = new DeflateStream(input, CompressionMode.Decompress);

            // Trust the recorded size only as a hint for how much room to take; never as a promise.
            int hint = part.UncompressedSize > 0 && part.UncompressedSize <= 8 * 1024 * 1024 ? (int)part.UncompressedSize : 0;
            using var output = new MemoryStream(hint);
            var buffer = new byte[81920];
            int read;
            while ((read = deflate.Read(buffer, 0, buffer.Length)) > 0)
            {
                if (output.Length + read > MaxPartSize)
                    throw new PackageException($"The part \"{part.Name}\" unpacks to more than half a gigabyte; Plain will not open it.");
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }
        catch (InvalidDataException)
        {
            throw new PackageException($"The part \"{part.Name}\" is damaged and could not be unpacked.");
        }
        catch (OutOfMemoryException)
        {
            throw new PackageException($"The part \"{part.Name}\" claims to be too large to unpack.");
        }
    }

    // ---------- writing ----------

    /// <summary>
    /// Write the file. The bytes go to a temporary file beside it first and only then take its place, so a save that
    /// is interrupted leaves the original whole rather than half written.
    ///
    /// Where the file already exists this uses Replace rather than Move, because Replace puts the new contents into
    /// the existing file and keeps what belongs to it: who may read it, when it was created, where it sits. Move
    /// would leave a brand new file wearing the old one's name, with whatever permissions the folder happened to
    /// hand out. For a program whose whole promise is not damaging your file, that difference matters.
    /// </summary>
    public void Save(string path)
    {
        var bytes = ToBytes();
        var tmp = path + ".plain-tmp";
        try
        {
            File.WriteAllBytes(tmp, bytes);
            if (File.Exists(path))
            {
                try { File.Replace(tmp, path, destinationBackupFileName: null); }
                catch (Exception ex) when (ex is PlatformNotSupportedException or IOException or UnauthorizedAccessException)
                {
                    // Replace needs both files on one volume and a filesystem that supports it; falling back to a
                    // move still leaves the file correct, only without the original's own permissions.
                    File.Move(tmp, path, overwrite: true);
                }
            }
            else File.Move(tmp, path);
        }
        finally
        {
            if (File.Exists(tmp)) { try { File.Delete(tmp); } catch { } }
        }
        SourcePath = path;
    }

    /// <summary>The whole file as bytes. Untouched parts are copied from the original byte for byte.</summary>
    public byte[] ToBytes()
    {
        if (Zip64Container && !_parts.Any(p => p.Edited)) return (byte[])_raw.Clone();
        if (Zip64Container) throw new PackageException("This package uses ZIP64 records; Plain will not edit it yet.");

        using var ms = new MemoryStream(_raw.Length + 4096);
        var newOffsets = new long[_parts.Count];

        for (int i = 0; i < _parts.Count; i++)
        {
            var part = _parts[i];
            newOffsets[i] = ms.Position;
            if (!part.Edited) { ms.Write(_raw, (int)part.LocalOffset, part.LocalLength); continue; }

            var content = part.Replacement!;
            byte[] payload = Deflate(content);
            ushort method = 8;
            if (payload.Length >= content.Length) { payload = content; method = 0; }
            uint crc = Crc32.Of(content);

            // Rebuild the local header, keeping the original's version, timestamps, name and extra bytes.
            int lNameLen = U16(_raw, (int)part.LocalOffset + 26);
            int lExtraLen = U16(_raw, (int)part.LocalOffset + 28);
            var header = new byte[30 + lNameLen + lExtraLen];
            Array.Copy(_raw, (int)part.LocalOffset, header, 0, header.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(6), (ushort)(part.Flags & ~0x8)); // no data descriptor
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(8), method);
            BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(14), crc);
            BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(18), (uint)payload.Length);
            BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(22), (uint)content.Length);
            ms.Write(header);
            ms.Write(payload);
        }

        long cdStart = ms.Position;
        for (int i = 0; i < _parts.Count; i++)
        {
            var part = _parts[i];
            var rec = new byte[part.CentralLength];
            Array.Copy(_raw, part.CentralOffset, rec, 0, rec.Length);
            BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(42), (uint)newOffsets[i]);
            if (part.Edited)
            {
                var content = part.Replacement!;
                byte[] payload = Deflate(content);
                ushort method = 8;
                if (payload.Length >= content.Length) { payload = content; method = 0; }
                BinaryPrimitives.WriteUInt16LittleEndian(rec.AsSpan(8), (ushort)(part.Flags & ~0x8));
                BinaryPrimitives.WriteUInt16LittleEndian(rec.AsSpan(10), method);
                BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(16), Crc32.Of(content));
                BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(20), (uint)payload.Length);
                BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(24), (uint)content.Length);
            }
            ms.Write(rec);
        }
        long cdEnd = ms.Position;

        var eocd = new byte[_eocdLength];
        Array.Copy(_raw, _eocdOffset, eocd, 0, eocd.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(eocd.AsSpan(12), (uint)(cdEnd - cdStart));
        BinaryPrimitives.WriteUInt32LittleEndian(eocd.AsSpan(16), (uint)cdStart);
        ms.Write(eocd);

        return ms.ToArray();
    }

    private static byte[] Deflate(byte[] content)
    {
        using var output = new MemoryStream(content.Length);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) deflate.Write(content);
        return output.ToArray();
    }

    // ---------- helpers ----------

    /// <summary>OPC part names are written with a leading slash; ZIP entry names are not. Accept either.</summary>
    public static string Normalize(string name) => name.StartsWith('/') ? name[1..] : name;

    /// <summary>UTF-8 text without letting a byte order mark reach the caller.</summary>
    public static string DecodeUtf8(byte[] bytes)
    {
        int start = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF ? 3 : 0;
        return Encoding.UTF8.GetString(bytes, start, bytes.Length - start);
    }

    private static ushort U16(byte[] b, int i) => BinaryPrimitives.ReadUInt16LittleEndian(b.AsSpan(i));
    private static uint U32(byte[] b, int i) => BinaryPrimitives.ReadUInt32LittleEndian(b.AsSpan(i));
    private static ulong U64(byte[] b, int i) => BinaryPrimitives.ReadUInt64LittleEndian(b.AsSpan(i));
}

/// <summary>The CRC-32 a ZIP entry carries. Written out so Plain needs no package beyond the framework.</summary>
public static class Crc32
{
    private static readonly uint[] Table = Build();

    private static uint[] Build()
    {
        var t = new uint[256];
        for (uint i = 0; i < 256; i++)
        {
            uint c = i;
            for (int k = 0; k < 8; k++) c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
            t[i] = c;
        }
        return t;
    }

    public static uint Of(ReadOnlySpan<byte> data)
    {
        uint c = 0xFFFFFFFFu;
        foreach (byte b in data) c = Table[(c ^ b) & 0xFF] ^ (c >> 8);
        return c ^ 0xFFFFFFFFu;
    }
}
