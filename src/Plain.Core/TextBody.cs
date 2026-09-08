using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// Word and PowerPoint both store text as paragraphs of runs: a paragraph carries the shape of the text, a run
/// carries a stretch of it with one set of formatting. The element names differ between the two, the structure does
/// not, so Plain reads and writes both through here.
/// </summary>
public sealed class TextShape
{
    private readonly XNamespace _ns;
    private readonly XElement _root;
    private readonly string _pName, _rName, _tName, _rPrName;

    public TextShape(XElement root, XNamespace ns, string paragraph = "p", string run = "r", string text = "t", string runProps = "rPr")
    {
        _root = root; _ns = ns; _pName = paragraph; _rName = run; _tName = text; _rPrName = runProps;
    }

    public IEnumerable<XElement> Paragraphs => _root.Descendants(_ns + _pName);

    public string TextOf(XElement paragraph) =>
        string.Concat(paragraph.Descendants(_ns + _tName).Select(t => t.Value));

    /// <summary>True when the paragraph is one unbroken run, so Plain can retype it without losing any formatting.</summary>
    public bool IsSingleRun(XElement paragraph) => paragraph.Elements(_ns + _rName).Count() == 1;

    /// <summary>
    /// Replace a paragraph's text. A one-run paragraph is edited in place and nothing is lost. A paragraph made of
    /// several differently formatted runs collapses to the first run's formatting, and the caller is told so it can
    /// say as much rather than silently flattening someone's bold word.
    /// </summary>
    public bool SetText(XElement paragraph, string value)
    {
        var runs = paragraph.Elements(_ns + _rName).ToList();

        if (runs.Count == 1)
        {
            var only = runs[0];
            var texts = only.Elements(_ns + _tName).ToList();
            if (texts.Count == 1) { SetRunText(texts[0], value); return true; }
        }

        var keptProps = runs.Count > 0 ? runs[0].Element(_ns + _rPrName) : null;
        foreach (var r in runs) r.Remove();

        var newRun = new XElement(_ns + _rName);
        if (keptProps is not null) newRun.Add(new XElement(keptProps));
        var t = new XElement(_ns + _tName);
        SetRunText(t, value);
        newRun.Add(t);

        // A paragraph's properties element must stay first; everything else follows it.
        var pPr = paragraph.Elements().FirstOrDefault(e => e.Name.LocalName is "pPr");
        if (pPr is not null) pPr.AddAfterSelf(newRun); else paragraph.AddFirst(newRun);

        return runs.Count <= 1;
    }

    private void SetRunText(XElement t, string value)
    {
        t.Value = value;
        // Word drops leading and trailing spaces unless the run says to keep them.
        if (_ns == Ns.Word && value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])))
            t.SetAttributeValue(XNamespace.Xml + "space", "preserve");
    }
}
