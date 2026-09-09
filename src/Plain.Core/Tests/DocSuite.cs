using System.Xml.Linq;
namespace Plain.Core.Tests;

public static class DocSuite
{
    public static Suite Run()
    {
        var s = new Suite("doc");
        if (Fixtures.Missing(s, "doc.docx")) return s;
        var path = Fixtures.Path_("doc.docx")!;

        var doc = new Document(OpcPackage.Open(path));
        s.Check("the document has blocks", doc.BlockCount > 4);

        var blocks = doc.Blocks().ToList();
        s.Check("the title is the first thing with text",
                blocks.First(b => b.Text.Length > 0).Text.Contains("Master Services Agreement"));
        s.Check("a heading is recognised", blocks.Any(b => b.IsHeading));
        s.Check("body text is read", doc.PlainText().Contains("thirty (30) days"));
        s.Check("table text is read", doc.PlainText().Contains("Discovery and audit"));
        s.Check("a table cell is marked as one", blocks.Any(b => b.Kind == BlockKind.TableCell));

        var work = Fixtures.Copy("doc.docx");
        try
        {
            var d = new Document(OpcPackage.Open(work));
            int target = d.Blocks().First(b => b.Text.Contains("thirty (30) days")).Index;
            int untouched = d.Blocks().First(b => b.Text.Contains("Discovery and audit")).Index;
            string keep = d.Read(untouched).Text;
            d.SetText(target, "The Client shall pay within forty-five (45) days of the invoice date.");
            d.Save(work);

            var again = new Document(OpcPackage.Open(work));
            s.Equal("the edited block reads back",
                    "The Client shall pay within forty-five (45) days of the invoice date.", again.Read(target).Text);
            s.Equal("a neighbouring block is untouched", keep, again.Read(untouched).Text);
            s.Equal("the block count is unchanged", d.BlockCount, again.BlockCount);
            s.Check("the heading survived the edit", again.Blocks().Any(b => b.IsHeading));

            // Text with edge-case characters must survive intact.
            var d2 = new Document(OpcPackage.Open(work));
            d2.SetText(target, "  spaces & <angles> and \"quotes\" kept  ");
            d2.Save(work);
            s.Equal("awkward characters survive", "  spaces & <angles> and \"quotes\" kept  ",
                    new Document(OpcPackage.Open(work)).Read(target).Text);
        }
        finally { try { File.Delete(work); } catch { } }

        // ---- what a heading looks like when it did not come from Word ----
        // LibreOffice writes a heading as an ordinary style with an outline level, and adds numbering that says
        // "none". Reading only the style name made every one of those a bullet, which is what a whole document of
        // headings looked like until this was checked.
        {
            var where = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "plain-outline-" + Guid.NewGuid().ToString("N") + ".docx");
            var file = PlainFile.Create(where);

            var written = Xml.Parse(file.Package.Read(Document.BodyPart));
            var d = Dialect.Of(written.Root!);
            var body = written.Root!.Element(d.Word + "body")!;

            XElement Para(string text, XElement? properties)
            {
                var p = new XElement(d.Word + "p");
                if (properties is not null) p.Add(properties);
                p.Add(new XElement(d.Word + "r", new XElement(d.Word + "t", text)));
                return p;
            }

            // A heading the LibreOffice way: Normal style, numbering that means none, an outline level.
            body.AddFirst(Para("A level two heading", new XElement(d.Word + "pPr",
                new XElement(d.Word + "pStyle", new XAttribute(d.Word + "val", "Normal")),
                new XElement(d.Word + "numPr",
                    new XElement(d.Word + "ilvl", new XAttribute(d.Word + "val", "0")),
                    new XElement(d.Word + "numId", new XAttribute(d.Word + "val", "0"))),
                new XElement(d.Word + "outlineLvl", new XAttribute(d.Word + "val", "1")))));

            body.AddFirst(Para("A heading with no style name", new XElement(d.Word + "pPr",
                new XElement(d.Word + "pStyle", new XAttribute(d.Word + "val", "Normal")),
                new XElement(d.Word + "numPr",
                    new XElement(d.Word + "ilvl", new XAttribute(d.Word + "val", "0")),
                    new XElement(d.Word + "numId", new XAttribute(d.Word + "val", "0"))),
                new XElement(d.Word + "outlineLvl", new XAttribute(d.Word + "val", "0")))));

            // And a real list item, which must still be one.
            body.Add(Para("An actual bullet", new XElement(d.Word + "pPr",
                new XElement(d.Word + "numPr",
                    new XElement(d.Word + "ilvl", new XAttribute(d.Word + "val", "0")),
                    new XElement(d.Word + "numId", new XAttribute(d.Word + "val", "3"))))));

            file.Package.Write(Document.BodyPart, Xml.ToBytes(written));
            var back = PlainFile.Read(file.Package.ToBytes(), file.Path).Document!;
            var outlined = back.Blocks().ToList();

            s.Check("a heading marked only by its outline level is a heading",
                outlined.Any(b => b.Text == "A heading with no style name" && b.Kind == BlockKind.Heading1));
            s.Check("and the level is read from it",
                outlined.Any(b => b.Text == "A level two heading" && b.Kind == BlockKind.Heading2));
            s.Check("numbering that says none does not make a list item",
                !outlined.Any(b => b.Text.StartsWith("A heading") && b.Kind == BlockKind.ListItem));
            s.Check("a paragraph with real numbering is still a list item",
                outlined.Any(b => b.Text == "An actual bullet" && b.Kind == BlockKind.ListItem));
        }

        return s;
    }
}
