using System.Text;

namespace Plain.Core;

/// <summary>
/// A new, empty Word, Excel or PowerPoint file, written part by part. Everything here is the smallest set of parts
/// the format actually requires, with nothing carried over from whatever tool happened to make a template: no other
/// program's name in the properties, no leftover styles, no theme nobody chose.
///
/// The parts are exactly the ones Plain understands, which means a file it makes is a file it can open, edit and
/// hand back unchanged. The tests check that round trip rather than trusting this code to be right.
/// </summary>
public static class Blank
{
    private static byte[] B(string xml) => new UTF8Encoding(false).GetBytes(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" + xml);

    private const string RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string DocRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static byte[] Make(FileKind kind) => kind switch
    {
        FileKind.Spreadsheet => Spreadsheet(),
        FileKind.Document => Document(),
        FileKind.Presentation => Presentation(),
        _ => throw new OpcPackage.PackageException("Plain makes Word, Excel and PowerPoint files. That is not one of them."),
    };

    /// <summary>The two properties parts every Office file carries. Plain names itself and claims nothing else.</summary>
    private static (string, byte[])[] Properties(string application) =>
    [
        ("docProps/core.xml", B(
            "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
            "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
            "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
            "<dc:title></dc:title><dc:creator></dc:creator><cp:lastModifiedBy></cp:lastModifiedBy>" +
            "</cp:coreProperties>")),
        ("docProps/app.xml", B(
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
            "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
            $"<Application>{application}</Application></Properties>")),
    ];

    private static byte[] RootRels(string mainPart) => B(
        $"<Relationships xmlns=\"{RelNs}\">" +
        $"<Relationship Id=\"rId1\" Type=\"{DocRel}/officeDocument\" Target=\"{mainPart}\"/>" +
        $"<Relationship Id=\"rId2\" Type=\"{DocRel}/extended-properties\" Target=\"docProps/app.xml\"/>" +
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
        "</Relationships>");

    private const string ContentTypeHead =
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>" +
        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>";

    // ---------- Excel ----------

    private static byte[] Spreadsheet()
    {
        const string ml = "application/vnd.openxmlformats-officedocument.spreadsheetml";
        var parts = new List<(string, byte[])>
        {
            ("[Content_Types].xml", B(ContentTypeHead +
                $"<Override PartName=\"/xl/workbook.xml\" ContentType=\"{ml}.sheet.main+xml\"/>" +
                $"<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"{ml}.worksheet+xml\"/>" +
                $"<Override PartName=\"/xl/styles.xml\" ContentType=\"{ml}.styles+xml\"/>" +
                $"<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"{ml}.sharedStrings+xml\"/>" +
                "</Types>")),
            ("_rels/.rels", RootRels("xl/workbook.xml")),
            ("xl/workbook.xml", B(
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                $"xmlns:r=\"{DocRel}\">" +
                "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                "<calcPr calcId=\"0\" fullCalcOnLoad=\"1\"/></workbook>")),
            ("xl/_rels/workbook.xml.rels", B(
                $"<Relationships xmlns=\"{RelNs}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{DocRel}/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                $"<Relationship Id=\"rId2\" Type=\"{DocRel}/styles\" Target=\"styles.xml\"/>" +
                $"<Relationship Id=\"rId3\" Type=\"{DocRel}/sharedStrings\" Target=\"sharedStrings.xml\"/>" +
                "</Relationships>")),
            ("xl/worksheets/sheet1.xml", B(
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                "<sheetData/></worksheet>")),
            // Excel expects the first fill to be none and the second gray125; it is a convention older than the format.
            ("xl/styles.xml", B(
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                "<fonts count=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>" +
                "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
                "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>" +
                "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
                "</styleSheet>")),
            // An empty shared string table, so the first piece of text typed has somewhere to go.
            ("xl/sharedStrings.xml", B(
                "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>")),
        };
        parts.AddRange(Properties("Plain for Windows"));
        return PackageBuilder.Build(parts);
    }

    // ---------- Word ----------

    private static byte[] Document()
    {
        const string ml = "application/vnd.openxmlformats-officedocument.wordprocessingml";
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var parts = new List<(string, byte[])>
        {
            ("[Content_Types].xml", B(ContentTypeHead +
                $"<Override PartName=\"/word/document.xml\" ContentType=\"{ml}.document.main+xml\"/>" +
                $"<Override PartName=\"/word/styles.xml\" ContentType=\"{ml}.styles+xml\"/>" +
                "</Types>")),
            ("_rels/.rels", RootRels("word/document.xml")),
            ("word/_rels/document.xml.rels", B(
                $"<Relationships xmlns=\"{RelNs}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{DocRel}/styles\" Target=\"styles.xml\"/>" +
                "</Relationships>")),
            // One empty paragraph, so there is a line to start typing on, and an A4 page for whatever opens it next.
            ("word/document.xml", B(
                $"<w:document xmlns:w=\"{w}\"><w:body>" +
                "<w:p/>" +
                "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>" +
                "<w:pgMar w:top=\"1134\" w:right=\"1134\" w:bottom=\"1134\" w:left=\"1134\" w:header=\"709\" w:footer=\"709\" w:gutter=\"0\"/>" +
                "</w:sectPr></w:body></w:document>")),
            ("word/styles.xml", B(
                $"<w:styles xmlns:w=\"{w}\">" +
                "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault>" +
                "<w:pPrDefault><w:pPr><w:spacing w:after=\"160\" w:line=\"259\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults>" +
                "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
                Heading(1, 32) + Heading(2, 26) + Heading(3, 24) +
                "</w:styles>")),
        };
        parts.AddRange(Properties("Plain for Windows"));
        return PackageBuilder.Build(parts);

        static string Heading(int level, int halfPoints) =>
            $"<w:style w:type=\"paragraph\" w:styleId=\"Heading{level}\"><w:name w:val=\"heading {level}\"/>" +
            "<w:basedOn w:val=\"Normal\"/><w:pPr><w:outlineLvl w:val=\"" + (level - 1) + "\"/>" +
            "<w:spacing w:before=\"240\" w:after=\"120\"/></w:pPr>" +
            $"<w:rPr><w:b/><w:sz w:val=\"{halfPoints}\"/></w:rPr></w:style>";
    }

    // ---------- PowerPoint ----------

    private static byte[] Presentation()
    {
        const string ml = "application/vnd.openxmlformats-officedocument.presentationml";
        const string p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const string a = "http://schemas.openxmlformats.org/drawingml/2006/main";

        // An empty shape tree is the frame every slide, layout and master hangs its shapes on.
        string tree =
            "<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
            "<p:grpSpPr/>";

        string Placeholder(int id, string name, string type, long x, long y, long cx, long cy, string text, int size) =>
            $"<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"{name}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>" +
            $"<p:nvPr><p:ph type=\"{type}\"/></p:nvPr></p:nvSpPr>" +
            $"<p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>" +
            "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
            $"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"{size}\"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>";

        var parts = new List<(string, byte[])>
        {
            ("[Content_Types].xml", B(ContentTypeHead +
                $"<Override PartName=\"/ppt/presentation.xml\" ContentType=\"{ml}.presentation.main+xml\"/>" +
                $"<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"{ml}.slideMaster+xml\"/>" +
                $"<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"{ml}.slideLayout+xml\"/>" +
                $"<Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"{ml}.slide+xml\"/>" +
                "<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>" +
                "</Types>")),
            ("_rels/.rels", RootRels("ppt/presentation.xml")),
            ("ppt/presentation.xml", B(
                $"<p:presentation xmlns:a=\"{a}\" xmlns:r=\"{DocRel}\" xmlns:p=\"{p}\">" +
                "<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>" +
                "<p:sldIdLst><p:sldId id=\"256\" r:id=\"rId2\"/></p:sldIdLst>" +
                "<p:sldSz cx=\"12192000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/>" +
                "</p:presentation>")),
            ("ppt/_rels/presentation.xml.rels", B(
                $"<Relationships xmlns=\"{RelNs}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{DocRel}/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>" +
                $"<Relationship Id=\"rId2\" Type=\"{DocRel}/slide\" Target=\"slides/slide1.xml\"/>" +
                $"<Relationship Id=\"rId3\" Type=\"{DocRel}/theme\" Target=\"theme/theme1.xml\"/>" +
                "</Relationships>")),
            ("ppt/slideMasters/slideMaster1.xml", B(
                $"<p:sldMaster xmlns:a=\"{a}\" xmlns:r=\"{DocRel}\" xmlns:p=\"{p}\">" +
                "<p:cSld>" + tree + "</p:spTree></p:cSld>" +
                "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" " +
                "accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>" +
                "<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst>" +
                "</p:sldMaster>")),
            ("ppt/slideMasters/_rels/slideMaster1.xml.rels", B(
                $"<Relationships xmlns=\"{RelNs}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{DocRel}/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>" +
                $"<Relationship Id=\"rId2\" Type=\"{DocRel}/theme\" Target=\"../theme/theme1.xml\"/>" +
                "</Relationships>")),
            ("ppt/slideLayouts/slideLayout1.xml", B(
                $"<p:sldLayout xmlns:a=\"{a}\" xmlns:r=\"{DocRel}\" xmlns:p=\"{p}\" type=\"obj\" preserve=\"1\">" +
                "<p:cSld name=\"Title and Content\">" + tree + "</p:spTree></p:cSld>" +
                "</p:sldLayout>")),
            ("ppt/slideLayouts/_rels/slideLayout1.xml.rels", B(
                $"<Relationships xmlns=\"{RelNs}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{DocRel}/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/>" +
                "</Relationships>")),
            // A title and a body, so a new deck has something to type into rather than an empty white rectangle.
            ("ppt/slides/slide1.xml", B(
                $"<p:sld xmlns:a=\"{a}\" xmlns:r=\"{DocRel}\" xmlns:p=\"{p}\">" +
                "<p:cSld>" + tree +
                Placeholder(2, "Title", "title", 838200, 610000, 10515600, 1325563, "Title", 4400) +
                Placeholder(3, "Content", "body", 838200, 2110000, 10515600, 3600000, "Your first point", 1800) +
                "</p:spTree></p:cSld></p:sld>")),
            ("ppt/slides/_rels/slide1.xml.rels", B(
                $"<Relationships xmlns=\"{RelNs}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{DocRel}/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>" +
                "</Relationships>")),
            ("ppt/theme/theme1.xml", B(Theme(a))),
        };
        parts.AddRange(Properties("Plain for Windows"));
        return PackageBuilder.Build(parts);
    }

    /// <summary>
    /// A theme. The format insists on exactly three fill styles, three line styles, three effect styles and three
    /// background fills, so the shortest legal theme is still this long. The colours are the ones Office ships with.
    /// </summary>
    private static string Theme(string a)
    {
        string fills =
            "<a:fillStyleLst>" +
            "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
            "<a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/></a:schemeClr></a:solidFill>" +
            "<a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"50000\"/></a:schemeClr></a:solidFill>" +
            "</a:fillStyleLst>";
        string lines =
            "<a:lnStyleLst>" +
            "<a:ln w=\"6350\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln>" +
            "<a:ln w=\"12700\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln>" +
            "<a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln>" +
            "</a:lnStyleLst>";
        string effects = "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>" +
            "<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>";
        string backgrounds =
            "<a:bgFillStyleLst>" +
            "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
            "<a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/></a:schemeClr></a:solidFill>" +
            "<a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/></a:schemeClr></a:solidFill>" +
            "</a:bgFillStyleLst>";

        return $"<a:theme xmlns:a=\"{a}\" name=\"Office\"><a:themeElements>" +
            "<a:clrScheme name=\"Office\">" +
            "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>" +
            "<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>" +
            "<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>" +
            "<a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>" +
            "<a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>" +
            "<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>" +
            "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>" +
            "</a:clrScheme>" +
            "<a:fontScheme name=\"Office\">" +
            "<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
            "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>" +
            "</a:fontScheme>" +
            "<a:fmtScheme name=\"Office\">" + fills + lines + effects + backgrounds + "</a:fmtScheme>" +
            "</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>";
    }
}
