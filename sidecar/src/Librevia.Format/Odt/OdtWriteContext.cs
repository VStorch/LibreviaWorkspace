using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>What the writers of one <c>content.xml</c> share.</summary>
internal sealed class OdtWriteContext(OdtStyles styles, OdtAutomaticStyles automatic, OdtPictures pictures, Inventory inventory)
{
    private const string TablePrefix = "LvTable";
    private const string ImagePrefix = "LvImage";

    private readonly HashSet<string> _names = new(StringComparer.Ordinal);

    public OdtStyles Styles => styles;
    public OdtAutomaticStyles Automatic => automatic;
    public OdtPictures Pictures => pictures;
    public Inventory Inventory => inventory;

    /// <summary>The list style of each numbering key written so far, so a list can continue another.</summary>
    public Dictionary<string, string> ListStyles { get; } = new(StringComparer.Ordinal);

    /// <summary>The table and frame names the part already has: a new one never repeats them.</summary>
    public void TakeNames(XElement root)
    {
        foreach (var element in root.Descendants())
        {
            if (element.Attr(Ns.Table + "name") is { } table) _names.Add(table);
            if (element.Attr(Ns.Draw + "name") is { } frame) _names.Add(frame);
        }
    }

    /// <summary>The name of the table being rewritten: formulas and links point to it.</summary>
    public string? KeptTableName { get; set; }

    /// <summary>The cell padding of the table being rewritten.</summary>
    public string? KeptCellPadding { get; set; }

    public string NewTableName()
    {
        var kept = KeptTableName;
        KeptTableName = null;
        return kept ?? NewName(TablePrefix);
    }

    public string NewImageName() => NewName(ImagePrefix);

    private string NewName(string prefix)
    {
        var number = 1;
        string name;
        while (!_names.Add(name = $"{prefix}{number}")) number++;
        return name;
    }

    /// <summary>The named style of the family for the editor's id, when the package has it.</summary>
    public string? NamedStyle(string family, string? id)
    {
        if (string.IsNullOrEmpty(id)) return null;
        var name = OdtWordStyles.OdfNameOf(family, id);
        return styles.IsNamed(family, name) ? name : null;
    }

    /// <summary>The fields of the block being rewritten, in order: a written field takes the attributes of its own.</summary>
    public List<XElement> SourceFields { get; } = [];

    public XElement? TakeSourceField(string localName)
    {
        var found = SourceFields.FirstOrDefault(field => field.Name == Ns.Text + localName);
        if (found is not null) SourceFields.Remove(found);
        return found;
    }

    public void NoteLoss(string key, string? value = null) => inventory.NoteLoss(key, value);
}
