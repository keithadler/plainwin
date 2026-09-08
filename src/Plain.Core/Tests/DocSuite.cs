namespace Plain.Core.Tests;

public static class DocSuite
{
    public static Suite Run()
    {
        var s = new Suite("doc");
        var path = Fixtures.Path_("doc.docx");
        if (path is null) { s.Check("fixtures found", false, "tests/fixtures/doc.docx is missing"); return s; }

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

        return s;
    }
}
