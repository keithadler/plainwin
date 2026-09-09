namespace Plain.Core.Tests;

/// <summary>
/// Adding and taking rows out of a table. A new row is a copy of a neighbour with the words removed, so the checks
/// care about two things: that the count changed, and that the row that arrived is empty but still looks like the
/// table it joined. And, as everywhere else, that the file still opens afterwards.
/// </summary>
public static class TableSuite
{
    public static Suite Run()
    {
        var s = new Suite("table");

        var corpus = Environment.GetEnvironmentVariable("PLAIN_CORPUS");
        var demo = corpus is not null ? System.IO.Path.Combine(corpus, "review.docx") : null;
        if (demo is null || !File.Exists(demo))
        {
            if (Fixtures.Missing(s, "doc.docx")) return s;
            demo = Fixtures.Path_("doc.docx");
        }

        PlainFile Work()
        {
            var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "plain-table-" + Guid.NewGuid().ToString("N") + ".docx");
            File.Copy(demo!, path, overwrite: true);
            return PlainFile.Open(path);
        }

        var first = Work();
        var shape = first.Document!.TableShape();
        if (shape.Count == 0)
        {
            s.Check("this document has no tables, so there is nothing to add a row to", true);
            return s;
        }

        s.Check("the document says how many rows its table has", shape[0] > 0);

        // ---- inserting ----
        {
            var file = Work();
            var doc = file.Document!;
            int was = doc.TableShape()[0];

            var done = doc.InsertRow(0, was);
            s.Check("a row can be added at the end", done is TableRows.Done);
            s.Check("the table has one more row", doc.TableShape()[0] == was + 1);

            file.Flush();
            var back = PlainFile.Read(file.Package.ToBytes(), file.Path);
            s.Check("the document still opens", back.Document is not null);
            s.Check("and still has the extra row", back.Document!.TableShape()[0] == was + 1);

            // The added row must be empty, or it would look like the row it was copied from.
            var cells = back.Document.Blocks().Where(b => b.InTable && b.Table == 0 && b.Row == was).ToList();
            s.Check("the new row has cells", cells.Count > 0);
            s.Check("and every one of them is empty", cells.All(c => c.Text.Trim().Length == 0));

            // What was already there is untouched.
            var older = back.Document.Blocks().Where(b => b.InTable && b.Table == 0 && b.Row == 0).ToList();
            s.Check("the row that was first is still there with its words",
                older.Any(c => c.Text.Trim().Length > 0));
        }

        // ---- inserting at the top ----
        {
            var file = Work();
            var doc = file.Document!;
            var firstRowWas = doc.Blocks().First(b => b.InTable && b.Table == 0 && b.Row == 0).Text;
            doc.InsertRow(0, 0);
            var nowSecond = doc.Blocks().First(b => b.InTable && b.Table == 0 && b.Row == 1).Text;
            s.Check("a row added at the top pushes the old first row down", nowSecond == firstRowWas);
            s.Check("and the new top row is empty",
                doc.Blocks().Where(b => b.InTable && b.Table == 0 && b.Row == 0).All(c => c.Text.Trim().Length == 0));
        }

        // ---- deleting ----
        {
            var file = Work();
            var doc = file.Document!;
            int was = doc.TableShape()[0];
            var done = doc.DeleteRow(0, was);
            s.Check("a row can be taken out", done is TableRows.Done);
            s.Check("the table has one fewer", doc.TableShape()[0] == was - 1);

            file.Flush();
            var back = PlainFile.Read(file.Package.ToBytes(), file.Path);
            s.Check("the document still opens after a removal", back.Document is not null);
            s.Check("and still has one fewer", back.Document!.TableShape()[0] == was - 1);
        }

        // ---- the refusals ----
        {
            var file = Work();
            var doc = file.Document!;
            s.Check("a table that is not there is refused", doc.InsertRow(99, 1) is TableRows.Refused);
            s.Check("a row that is not there is refused", doc.DeleteRow(0, 999) is TableRows.Refused);

            while (doc.TableShape()[0] > 1)
                if (doc.DeleteRow(0, 1) is TableRows.Refused) break;
            s.Check("a table is never left with no rows", doc.DeleteRow(0, 1) is TableRows.Refused);
        }

        return s;
    }
}
