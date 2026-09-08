using System.Xml.Linq;

namespace Plain.Core;

/// <summary>Open XML never names another part directly; it names a relationship id that a _rels part resolves.</summary>
public sealed class Rels
{
    private readonly Dictionary<string, string> _targets = new(StringComparer.Ordinal);
    private readonly string _baseDir;

    public Rels(OpcPackage pkg, string partName)
    {
        partName = OpcPackage.Normalize(partName);
        int slash = partName.LastIndexOf('/');
        _baseDir = slash < 0 ? "" : partName[..(slash + 1)];
        string relsPart = _baseDir + "_rels/" + partName[(slash + 1)..] + ".rels";
        if (!pkg.Has(relsPart)) return;

        foreach (var r in Xml.Parse(pkg.Read(relsPart)).Root!.Elements(Ns.Pkg + "Relationship"))
        {
            string? id = (string?)r.Attribute("Id"), target = (string?)r.Attribute("Target");
            if (id is null || target is null) continue;
            if (string.Equals((string?)r.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
            _targets[id] = Resolve(target);
        }
    }

    private string Resolve(string target)
    {
        if (target.StartsWith('/')) return target[1..];
        var stack = new List<string>(_baseDir.Split('/', StringSplitOptions.RemoveEmptyEntries));
        foreach (var seg in target.Split('/'))
        {
            if (seg is "." or "") continue;
            if (seg == "..") { if (stack.Count > 0) stack.RemoveAt(stack.Count - 1); }
            else stack.Add(seg);
        }
        return string.Join('/', stack);
    }

    /// <summary>The part a relationship id points at, or null when the id is unknown.</summary>
    public string? this[string id] => _targets.TryGetValue(id, out var t) ? t : null;
}
