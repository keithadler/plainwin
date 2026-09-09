using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// The properties a file carries about itself: who wrote it, what it is called, which company, when it was last saved
/// and by whom. Office fills these in without telling anyone, which is how a document prepared for one client reaches
/// another still carrying the first one's name. Plain can show them, change them, and clear them out.
/// </summary>
public sealed class Properties
{
    private const string CorePart = "docProps/core.xml";
    private const string AppPart = "docProps/app.xml";

    private static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DcTerms = "http://purl.org/dc/terms/";
    private static readonly XNamespace Ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    private readonly OpcPackage _pkg;
    private XDocument? _core, _app;
    private bool _coreDirty, _appDirty;

    public Properties(OpcPackage pkg)
    {
        _pkg = pkg;
        if (pkg.Has(CorePart)) { try { _core = Xml.Parse(pkg.Read(CorePart)); } catch { } }
        if (pkg.Has(AppPart)) { try { _app = Xml.Parse(pkg.Read(AppPart)); } catch { } }
    }

    /// <summary>
    /// Everything Plain can read, in the order a person would want to check it. The name is what the panel shows;
    /// the value is what is in the file, empty when the property is absent.
    /// </summary>
    public IReadOnlyList<(string Name, string Value)> All()
    {
        var list = new List<(string, string)>();
        void Add(string label, string? value) => list.Add((label, value ?? ""));

        Add("Title", Core("title"));
        Add("Subject", Core("subject"));
        Add("Author", Core("creator"));
        Add("Last saved by", Core("lastModifiedBy"));
        Add("Company", App("Company"));
        Add("Manager", App("Manager"));
        Add("Keywords", Core("keywords"));
        Add("Category", Core("category"));
        Add("Comments", Core("description"));
        Add("Created", Core("created"));
        Add("Modified", Core("modified"));
        Add("Revision", Core("revision"));
        Add("Total editing time", App("TotalTime"));
        Add("Made with", App("Application"));
        return list;
    }

    /// <summary>The ones that name a person or an organisation, which are the ones that leak.</summary>
    public static readonly string[] Identifying =
        { "Author", "Last saved by", "Company", "Manager", "Title", "Subject", "Keywords", "Category", "Comments" };

    /// <summary>What is filled in that names somebody. This is what a person wants to see before sending a file out.</summary>
    public IReadOnlyList<(string Name, string Value)> Revealing() =>
        All().Where(p => Identifying.Contains(p.Name) && p.Value.Length > 0).ToList();

    private string? Core(string name)
    {
        if (_core?.Root is null) return null;
        var ns = name is "title" or "subject" or "creator" or "description" ? Dc
               : name is "created" or "modified" ? DcTerms
               : Cp;
        return _core.Root.Element(ns + name)?.Value;
    }

    private string? App(string name) => _app?.Root?.Element(Ep + name)?.Value;

    /// <summary>Set a property by the name <see cref="All"/> uses. Setting it empty removes it.</summary>
    public bool Set(string name, string value)
    {
        switch (name)
        {
            case "Title": return SetCore(Dc + "title", value);
            case "Subject": return SetCore(Dc + "subject", value);
            case "Author": return SetCore(Dc + "creator", value);
            case "Comments": return SetCore(Dc + "description", value);
            case "Last saved by": return SetCore(Cp + "lastModifiedBy", value);
            case "Keywords": return SetCore(Cp + "keywords", value);
            case "Category": return SetCore(Cp + "category", value);
            case "Revision": return SetCore(Cp + "revision", value);
            case "Created": return SetCore(DcTerms + "created", value);
            case "Modified": return SetCore(DcTerms + "modified", value);
            case "Company": return SetApp(Ep + "Company", value);
            case "Manager": return SetApp(Ep + "Manager", value);
            case "Total editing time": return SetApp(Ep + "TotalTime", value);
            case "Made with": return SetApp(Ep + "Application", value);
            default: return false;
        }
    }

    private bool SetCore(XName name, string value)
    {
        if (_core?.Root is null) return false;
        var element = _core.Root.Element(name);
        if (value.Length == 0) { element?.Remove(); _coreDirty = true; return true; }
        if (element is null) _core.Root.Add(new XElement(name, value));
        else element.Value = value;
        _coreDirty = true;
        return true;
    }

    private bool SetApp(XName name, string value)
    {
        if (_app?.Root is null) return false;
        var element = _app.Root.Element(name);
        if (value.Length == 0) { element?.Remove(); _appDirty = true; return true; }
        if (element is null) _app.Root.Add(new XElement(name, value));
        else element.Value = value;
        _appDirty = true;
        return true;
    }

    /// <summary>
    /// Clear every property that names a person or an organisation, and say how many were cleared. The dates and the
    /// counts stay: they say nothing about who you are, and removing them makes a file look tampered with rather than
    /// clean. This changes only the two properties parts; the document itself is untouched.
    /// </summary>
    public int Strip()
    {
        int cleared = 0;
        foreach (var (name, value) in All())
            if (Identifying.Contains(name) && value.Length > 0 && Set(name, "")) cleared++;
        return cleared;
    }

    public void Flush()
    {
        if (_coreDirty && _core is not null) { _pkg.Write(CorePart, Xml.ToBytes(_core)); _coreDirty = false; }
        if (_appDirty && _app is not null) { _pkg.Write(AppPart, Xml.ToBytes(_app)); _appDirty = false; }
    }
}
