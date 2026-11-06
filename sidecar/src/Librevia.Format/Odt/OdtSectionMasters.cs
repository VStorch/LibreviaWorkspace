using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>A master page whose layout follows a section: an existing one, or a copy of <c>From</c>.</summary>
internal sealed record OdtMasterChange(string Name, string? From, PageSetupDto Known, PageSetupDto Wanted);

/// <summary>
/// The master page each of the model's sections opens. A section read from the file keeps its own
/// (by its id; the page by being last); a new one takes the template's <c>SectionN</c>, or the
/// previous section's. Sections that share a master page and want different setups split it: the
/// one closest to the file keeps it, each other setup gets a copy.
/// </summary>
internal sealed class OdtSectionMasters
{
    private const string CopyPrefix = "LvPage";

    public required IReadOnlyList<OdtReadSection> Read { get; init; }

    /// <summary>By the model's section: the read section it is, or <c>null</c> for a new one.</summary>
    public required IReadOnlyList<int?> ReadOf { get; init; }

    public required IReadOnlyList<string> Names { get; init; }
    public required IReadOnlyList<OdtMasterChange> Changes { get; init; }
    public required ISet<string> Available { get; init; }

    public static OdtSectionMasters Assign(OdtSource source, List<PageSetupDto> sections, ISet<string> existing)
    {
        var read = source.Sections;
        var readOf = new int?[sections.Count];
        var names = new string[sections.Count];
        for (var index = 0; index < sections.Count; index++)
        {
            readOf[index] = ReadIndexOf(read, sections, index);
            names[index] = readOf[index] is { } known
                ? read[known].Master ?? OdtPageReader.DefaultMaster
                : Fallback(read, existing, names, index, sections.Count);
        }

        var available = new HashSet<string>(existing, StringComparer.Ordinal);
        var changes = new List<OdtMasterChange>();
        foreach (var name in names.Distinct().ToList()) Split(source, sections, names, name, available, changes);
        return new OdtSectionMasters { Read = read, ReadOf = readOf, Names = names, Changes = changes, Available = available };
    }

    private static int? ReadIndexOf(IReadOnlyList<OdtReadSection> read, List<PageSetupDto> sections, int index)
    {
        if (index == sections.Count - 1) return read.Count - 1;
        if (sections[index].Id is not { } id) return null;
        for (var known = 0; known < read.Count - 1; known++)
        {
            if (read[known].Setup.Id == id) return known;
        }

        return null;
    }

    private static string Fallback(IReadOnlyList<OdtReadSection> read, ISet<string> existing, string[] names, int index, int count)
    {
        var template = OdtSectionPlan.MasterNameOf(index, count);
        if (existing.Contains(template) && read.All(section => section.Master != template)) return template;
        return index > 0 ? names[index - 1] : read[0].Master ?? OdtPageReader.DefaultMaster;
    }

    /// <summary>The sections on <paramref name="name"/>, grouped by setup.</summary>
    private static void Split(
        OdtSource source, List<PageSetupDto> sections, string[] names, string name, HashSet<string> available, List<OdtMasterChange> changes)
    {
        var known = source.Sections.FirstOrDefault(section => (section.Master ?? OdtPageReader.DefaultMaster) == name)?.Setup
                    ?? OdtPageReader.Read(source.Package, name, new Inventory());
        var groups = new List<(PageSetupDto Wanted, List<int> Sections)>();
        for (var index = 0; index < sections.Count; index++)
        {
            if (names[index] != name) continue;
            var group = groups.FindIndex(candidate => OdtStylesPart.SameLayout(candidate.Wanted, sections[index]));
            if (group < 0) groups.Add((sections[index], [index]));
            else groups[group].Sections.Add(index);
        }

        var keeper = groups.FindIndex(group => OdtStylesPart.SameLayout(known, group.Wanted));
        if (keeper < 0)
        {
            keeper = 0;
            changes.Add(new OdtMasterChange(name, null, known, groups[0].Wanted));
        }

        for (var group = 0; group < groups.Count; group++)
        {
            if (group == keeper) continue;
            var copy = NewName(available);
            changes.Add(new OdtMasterChange(copy, name, known, groups[group].Wanted));
            foreach (var index in groups[group].Sections) names[index] = copy;
        }
    }

    private static string NewName(HashSet<string> available)
    {
        var number = 1;
        while (!available.Add($"{CopyPrefix}{number}")) number++;
        return $"{CopyPrefix}{number}";
    }
}
