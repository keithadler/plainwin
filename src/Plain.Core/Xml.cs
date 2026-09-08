using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// The Open XML namespaces Plain needs, and the one way it writes XML back. Office files declare a standalone
/// document with no indentation; writing them back any other way is legal but noisy in a diff, so Plain matches.
/// </summary>
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
    public static XDocument Parse(byte[] bytes) =>
        XDocument.Parse(OpcPackage.DecodeUtf8(bytes), LoadOptions.PreserveWhitespace);

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
