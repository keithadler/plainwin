using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// The Open XML namespaces Plain needs, and the one way it writes XML back. Office files declare a standalone
/// document with no indentation; writing them back any other way is legal but noisy in a diff, so Plain matches.
/// </summary>
/// <summary>
/// Which family of names a file was written in. Office can save the same document two ways: the transitional shape
/// nearly everything uses, and the strict ISO shape that Excel offers as "Strict Open XML". The elements are the
/// same, the namespace they live in is not, so a reader that knows only one finds an empty file and says nothing.
/// Plain works out which family a part is in and reads it either way; it writes back the family it found.
/// </summary>
public sealed class Dialect
{
    public required XNamespace Sheet { get; init; }
    public required XNamespace Word { get; init; }
    public required XNamespace Pres { get; init; }
    public required XNamespace Draw { get; init; }
    public required XNamespace Rel { get; init; }
    public required bool Strict { get; init; }

    public static readonly Dialect Transitional = new()
    {
        Sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        Pres = "http://schemas.openxmlformats.org/presentationml/2006/main",
        Draw = "http://schemas.openxmlformats.org/drawingml/2006/main",
        Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        Strict = false,
    };

    public static readonly Dialect Iso = new()
    {
        Sheet = "http://purl.oclc.org/ooxml/spreadsheetml/main",
        Word = "http://purl.oclc.org/ooxml/wordprocessingml/main",
        Pres = "http://purl.oclc.org/ooxml/presentationml/main",
        Draw = "http://purl.oclc.org/ooxml/drawingml/main",
        Rel = "http://purl.oclc.org/ooxml/officeDocument/relationships",
        Strict = true,
    };

    /// <summary>The family a part is written in, taken from the namespace its root element sits in.</summary>
    public static Dialect Of(XElement root) =>
        root.Name.NamespaceName.Contains("purl.oclc.org/ooxml", StringComparison.Ordinal) ? Iso : Transitional;
}

public static class Ns
{
    public static readonly XNamespace Sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public static readonly XNamespace Pres = "http://schemas.openxmlformats.org/presentationml/2006/main";
    public static readonly XNamespace Draw = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public static readonly XNamespace Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace Pkg = "http://schemas.openxmlformats.org/package/2006/relationships";
    public static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
}

public static class Xml
{
    /// <summary>
    /// Parse a part. A damaged file is a sentence someone can act on, not a stack trace, so the XML reader's own
    /// complaint is turned into one here rather than at every call site.
    /// </summary>
    public static XDocument Parse(byte[] bytes)
    {
        try { return XDocument.Parse(OpcPackage.DecodeUtf8(bytes), LoadOptions.PreserveWhitespace); }
        catch (XmlException ex)
        {
            throw new OpcPackage.PackageException(
                $"Part of this file is damaged and could not be read (line {ex.LineNumber}: {ex.Message})");
        }
        catch (Exception ex) when (ex is ArgumentException or FormatException or InvalidOperationException)
        {
            throw new OpcPackage.PackageException("Part of this file is damaged and could not be read.");
        }
    }

    /// <summary>
    /// A whole number from an attribute, or null. Casting an attribute to int? throws on anything that is not a
    /// number, and a damaged file puts anything anywhere, so nothing in Plain casts one directly.
    /// </summary>
    public static int? Int(XAttribute? attribute) =>
        attribute is not null && int.TryParse(attribute.Value, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var value) ? value : null;

    public static int Int(XAttribute? attribute, int fallback) => Int(attribute) ?? fallback;

    /// <summary>Serialize without adding whitespace, with the standalone declaration Office writes.</summary>
    public static byte[] ToBytes(XDocument doc)
    {
        var settings = new XmlWriterSettings
        {
            Indent = false,
            OmitXmlDeclaration = true,
            Encoding = new UTF8Encoding(false),
            NewLineHandling = NewLineHandling.None,
        };
        using var ms = new MemoryStream();
        ms.Write(new UTF8Encoding(false).GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"));
        using (var w = XmlWriter.Create(ms, settings)) doc.Save(w);
        return ms.ToArray();
    }
}
