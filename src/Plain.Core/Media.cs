namespace Plain.Core;

/// <summary>A picture that is in the file: what it is called, what kind, how big, and its bytes.</summary>
public sealed record Picture(string Part, string Kind, long Bytes, Func<byte[]> Read)
{
    public string Size => Bytes >= 1_000_000 ? $"{Bytes / 1_048_576.0:0.0} MB"
                        : Bytes >= 1_000 ? $"{Bytes / 1024.0:0.0} KB"
                        : $"{Bytes} B";

    public string Name => Part[(Part.LastIndexOf('/') + 1)..];
}

/// <summary>
/// The pictures inside a file. Plain preserves them whether or not it draws them, which for anyone about to send a
/// document onward is the wrong way round: a scanned signature page or a screenshot with somebody's details in it is
/// exactly the thing to look at before it goes. So they can be listed, looked at, and saved out.
/// </summary>
public static class Media
{
    private static readonly Dictionary<string, string> Kinds = new(StringComparer.OrdinalIgnoreCase)
    {
        [".png"] = "PNG image", [".jpg"] = "JPEG image", [".jpeg"] = "JPEG image", [".gif"] = "GIF image",
        [".bmp"] = "Bitmap", [".tif"] = "TIFF image", [".tiff"] = "TIFF image", [".emf"] = "Windows metafile",
        [".wmf"] = "Windows metafile", [".svg"] = "SVG image", [".webp"] = "WebP image", [".ico"] = "Icon",
    };

    /// <summary>Can this be drawn on screen without help? Metafiles and SVG cannot, and are listed rather than shown.</summary>
    public static bool Drawable(string part)
    {
        var extension = System.IO.Path.GetExtension(part).ToLowerInvariant();
        return extension is ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".tif" or ".tiff";
    }

    public static IReadOnlyList<Picture> In(PlainFile file)
    {
        var pictures = new List<Picture>();
        foreach (var part in file.Package.Parts)
        {
            var extension = System.IO.Path.GetExtension(part.Name).ToLowerInvariant();
            if (!Kinds.TryGetValue(extension, out var kind)) continue;
            if (!part.Name.Contains("/media/", StringComparison.OrdinalIgnoreCase)
                && !part.Name.Contains("/images/", StringComparison.OrdinalIgnoreCase)) continue;

            var name = part.Name;
            pictures.Add(new Picture(name, kind, part.Size, () => file.Package.Read(name)));
        }
        return pictures;
    }

    /// <summary>Write a picture out to a file of its own, so it can be looked at in anything.</summary>
    public static string SaveTo(Picture picture, string folder)
    {
        Directory.CreateDirectory(folder);
        var path = System.IO.Path.Combine(folder, picture.Name);
        File.WriteAllBytes(path, picture.Read());
        return path;
    }
}
