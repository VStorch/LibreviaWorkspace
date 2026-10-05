using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// Multilevel lists: the definition the reader hands the screen and what saving does with it:
/// nothing, when nothing changed; a new `w:num`, when the list restarts; one definition per list,
/// when the list was born in the editor.
/// </summary>
public class ListNumberingTests
{
    private static List<Node> ListsOf(DocumentModelDto model) =>
        Roundtrip.Walk(model.Doc).Where(node => node.Type is "orderedList" or "bulletList").ToList();

    private static JsonNode? AttrOf(Node node, string name) =>
        node.Attrs is { } attrs && attrs.TryGetValue(name, out var value) ? value : null;

    private static int? IntOf(Node node, string name) => AttrOf(node, name)?.GetValue<int>();

    private static JsonObject DefinitionOf(Node list) => (JsonObject)AttrOf(list, "numbering")!;

    [Fact]
    public void LeitorEntregaOsNiveisEAChaveDaContagem()
    {
        var lists = ListsOf(Roundtrip.Open(Fixtures.WithMultilevelList()));

        // One, the sublist, Three (same `numId`, across the paragraph) and Ten.
        Assert.Equal(4, lists.Count);

        var first = DefinitionOf(lists[0]);
        Assert.Equal("a3", first["key"]!.GetValue<string>());
        Assert.Equal("lowerLetter", first["levels"]![1]!["fmt"]!.GetValue<string>());
        Assert.Equal("%1.%2)", first["levels"]![1]!["text"]!.GetValue<string>());

        // The sublist has the same numbering: it finds the definition in the outer list.
        Assert.Null(AttrOf(lists[1], "numbering"));
        Assert.Equal(5, IntOf(lists[1], "numId"));

        // The same `numId` after the paragraph continues the same count.
        Assert.Equal("a3", DefinitionOf(lists[2])["key"]!.GetValue<string>());

        // The `w:num` with a restart counts on its own, from 10.
        var restarted = DefinitionOf(lists[3]);
        Assert.Equal("n6", restarted["key"]!.GetValue<string>());
        Assert.Equal(10, restarted["overrides"]!["0"]!.GetValue<int>());
    }

    [Fact]
    public void OutraNumeracaoNoMesmoNivelEOutraLista()
    {
        // Joined, "Dez" would be an item of "Três": 4, not 10.
        var lists = ListsOf(Roundtrip.Open(Fixtures.WithMultilevelList()));
        Assert.Equal(6, IntOf(lists[3], "numId"));
        Assert.Single(lists[3].Content!);
    }

    [Fact]
    public void AbrirEGravarSemEditarNaoTocaNaNumeracao()
    {
        var original = Fixtures.WithMultilevelList();
        var (saved, result) = Roundtrip.Save(original, Roundtrip.Clone(Roundtrip.Open(original)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Equal(Roundtrip.PartsOf(original)["word/numbering.xml"], Roundtrip.PartsOf(saved)["word/numbering.xml"]);
    }

    [Fact]
    public void ListaReiniciadaGanhaNumComStartOverrideDaMesmaDefinicao()
    {
        // "Restart at 1": a new `w:num` of the same definition, with `w:startOverride`, as in Word.
        var original = Fixtures.WithMultilevelList();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var third = ListsOf(model)[2];

        var definition = (JsonObject)DefinitionOf(third).DeepClone();
        definition["key"] = "nova-1";
        definition["overrides"] = new JsonObject { ["0"] = 1 };
        third.With("numId", null).With("numbering", definition);

        var (saved, _) = Roundtrip.Save(original, model);
        var numbering = Roundtrip.XmlOf(saved, "word/numbering.xml");

        Assert.Single(Regex.Matches(numbering, "<w:abstractNum "));
        Assert.Matches("<w:num w:numId=\"7\"><w:abstractNumId w:val=\"3\" /><w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"1\" />", numbering);

        var reopened = ListsOf(Roundtrip.Open(saved))[2];
        Assert.Equal(7, IntOf(reopened, "numId"));
        Assert.Equal("n7", DefinitionOf(reopened)["key"]!.GetValue<string>());
        Assert.Equal(1, DefinitionOf(reopened)["overrides"]!["0"]!.GetValue<int>());
    }

    [Fact]
    public void ItemDescidoComTabVaiAoNivelDeBaixoSemReescreverOParagrafo()
    {
        // It changes the surrounding list, not the item; returned as it came, it would go back to
        // level 0.
        var original = Fixtures.WithBulletList();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var list = ListsOf(model)[0];
        var second = list.Content![1];
        list.Content.RemoveAt(1);
        list.Content[0].Content!.Add(Node.Of("bulletList", second));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Matches("<w:ilvl w:val=\"1\" />(<w:numId w:val=\"1\" />)?.*Segundo item", xml);
        Assert.Equal(Roundtrip.PartsOf(original)["word/numbering.xml"], Roundtrip.PartsOf(saved)["word/numbering.xml"]);
    }

    [Fact]
    public void CadaListaNovaGanhaDefinicaoPropriaComOsNiveisDoWord()
    {
        // With a single definition, the second new list would continue the first one's count in
        // Word.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        Node Item(string text) => Node.Of("listItem", Node.Of("paragraph", new Node { Type = "text", Text = text }));

        model.Doc.Content!.Add(Node.Of("orderedList", Item("um"), Item("dois")));
        model.Doc.Content!.Add(Node.Of("paragraph", new Node { Type = "text", Text = "meio" }));
        model.Doc.Content!.Add(Node.Of("orderedList", Item("um de novo")));

        var (saved, _) = Roundtrip.Save(original, model);
        var numbering = Roundtrip.XmlOf(saved, "word/numbering.xml");

        Assert.Equal(2, Regex.Matches(numbering, "<w:abstractNum ").Count);
        Assert.Contains("w:numFmt w:val=\"lowerLetter\"", numbering, StringComparison.Ordinal);
        Assert.Contains("w:numFmt w:val=\"lowerRoman\"", numbering, StringComparison.Ordinal);

        // And the levels that come back are the defaults the screen drew.
        var reopened = ListsOf(Roundtrip.Open(saved));
        Assert.Equal(2, reopened.Count);
        Assert.True(JsonNode.DeepEquals(ListLevels.Defaults("orderedList"), DefinitionOf(reopened[0])["levels"]));
        Assert.NotEqual(
            DefinitionOf(reopened[0])["key"]!.GetValue<string>(),
            DefinitionOf(reopened[1])["key"]!.GetValue<string>());
    }

    [Fact]
    public void ListasNovasComAMesmaChaveSaoUmaNumeracaoSo()
    {
        // "Continue numbering" between two lists not saved yet: the key is what joins them, and
        // both must come out in the same `w:num`.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var definition = new JsonObject { ["key"] = "nova-junta", ["levels"] = ListLevels.Defaults("orderedList") };
        Node List(string text) => Node.Of(
                "orderedList",
                Node.Of("listItem", Node.Of("paragraph", new Node { Type = "text", Text = text })))
            .With("numbering", definition.DeepClone());

        model.Doc.Content!.Add(List("um"));
        model.Doc.Content!.Add(Node.Of("paragraph", new Node { Type = "text", Text = "meio" }));
        model.Doc.Content!.Add(List("dois"));

        var (saved, _) = Roundtrip.Save(original, model);
        var numbering = Roundtrip.XmlOf(saved, "word/numbering.xml");

        Assert.Single(Regex.Matches(numbering, "<w:num "));
        var lists = ListsOf(Roundtrip.Open(saved));
        Assert.Equal(IntOf(lists[0], "numId"), IntOf(lists[1], "numId"));
    }

    [Fact]
    public void ListaComDefinicaoDaGaleriaSaiComOsNiveisDela()
    {
        // A definition chosen in the editor (gallery, pasted list) comes back the same.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var levels = ListLevels.Defaults("orderedList");
        levels[0] = ListLevels.Level("upperRoman", "%1.", 1, 12.7, 6.35);
        levels[1] = ListLevels.Level("decimal", "%1.%2.", 1, 25.4, 6.35, legal: true);
        var definition = new JsonObject { ["key"] = "galeria-1", ["levels"] = levels };

        var list = Node.Of("orderedList", Node.Of("listItem", Node.Of("paragraph", new Node { Type = "text", Text = "I" })));
        list.With("numbering", definition);
        model.Doc.Content!.Add(list);

        var (saved, _) = Roundtrip.Save(original, model);
        var numbering = Roundtrip.XmlOf(saved, "word/numbering.xml");
        Assert.Contains("w:numFmt w:val=\"upperRoman\"", numbering, StringComparison.Ordinal);
        Assert.Contains("<w:isLgl />", numbering, StringComparison.Ordinal);

        var back = DefinitionOf(ListsOf(Roundtrip.Open(saved))[0]);
        Assert.True(JsonNode.DeepEquals(levels, back["levels"]));
    }

    [Fact]
    public void MarcadoresDaTelaVoltamAoGlifoEAFonteDoWord()
    {
        Assert.Equal(("", "Symbol"), ListLevels.GlyphOf("•"));
        Assert.Equal(("", "Wingdings"), ListLevels.GlyphOf("▪"));
        Assert.Equal(("–", (string?)null), ListLevels.GlyphOf("–"));
        Assert.Equal("▪", ListLevels.Shown(""));
    }

    [Fact]
    public void ListaColadaComNumIdQueODestinoUsaParaOutraCoisaGanhaNumeracaoPropria()
    {
        // A list from another document carries `numId` 5, which here is another numbering.
        var original = Fixtures.WithMultilevelList();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var definition = new JsonObject { ["key"] = "a7", ["abstractId"] = 7, ["levels"] = ListLevels.Defaults("bulletList") };
        var pasted = Node.Of("bulletList", Node.Of("listItem", Node.Of("paragraph", new Node { Type = "text", Text = "Colado" })))
            .With("numId", 5)
            .With("numbering", definition);
        model.Doc.Content!.Add(pasted);

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Matches("<w:numId w:val=\"7\" />.*Colado", xml);
        var reopened = ListsOf(Roundtrip.Open(saved)).Last();
        Assert.Equal("bullet", DefinitionOf(reopened)["levels"]![0]!["fmt"]!.GetValue<string>());
    }

    [Fact]
    public void NumIdSemDefinicaoNaoEListaENaoEReescrito()
    {
        var original = Fixtures.WithDanglingNumbering();
        var opened = Roundtrip.Open(original);
        Assert.Empty(ListsOf(opened));

        var (saved, result) = Roundtrip.Save(original, Roundtrip.Clone(opened));
        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Equal(Roundtrip.PartsOf(original)["word/numbering.xml"], Roundtrip.PartsOf(saved)["word/numbering.xml"]);
    }

    [Fact]
    public void NivelNaoMudadoECopiadoDoOriginalAoRecriarADefinicao()
    {
        // The gallery only touched level 2: level 1 comes back from the file as it was
        // (right-aligned, red, `ordinal`), not rebuilt from what the screen knows.
        var original = Fixtures.WithRichNumbering();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var list = ListsOf(model)[0];
        var definition = (JsonObject)DefinitionOf(list).DeepClone();
        Assert.Equal("ordinal", definition["levels"]![0]!["fmt"]!.GetValue<string>());
        Assert.Equal("right", definition["levels"]![0]!["jc"]!.GetValue<string>());

        definition["key"] = "galeria-2";
        definition["levels"]![1] = ListLevels.Level("upperRoman", "%2)", 1, 25.4, 6.35);
        list.With("numId", null).With("numbering", definition);

        var (saved, result) = Roundtrip.Save(original, model);
        var numbering = Roundtrip.XmlOf(saved, "word/numbering.xml");

        Assert.Equal(2, Regex.Matches(numbering, "<w:abstractNum ").Count);
        Assert.Equal(2, Regex.Matches(numbering, "w:val=\"FF0000\"").Count);
        Assert.Equal(2, Regex.Matches(numbering, "w:numFmt w:val=\"ordinal\"").Count);
        Assert.Contains("upperRoman", numbering, StringComparison.Ordinal);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void NivelRecriadoSemOriginalLevaAlinhamentoEFormatoEAvisaOQueFicou()
    {
        // Pasted from another document: there is no `w:lvl` to copy. What the definition carries
        // comes back (right-aligned, `ordinal`); the number color does not, and a warning goes out.
        var original = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var levels = ListLevels.Defaults("orderedList");
        var rich = ListLevels.Level("ordinal", "%1", 1, 12.7, 6.35);
        rich["jc"] = "right";
        rich["extra"] = true;
        levels[0] = rich;
        model.Doc.Content!.Add(Node.Of("orderedList", Node.Of("listItem", Node.Of("paragraph", new Node { Type = "text", Text = "x" })))
            .With("numbering", new JsonObject { ["key"] = "colada-1", ["levels"] = levels }));

        var (saved, result) = Roundtrip.Save(original, model);
        var numbering = Roundtrip.XmlOf(saved, "word/numbering.xml");

        Assert.Contains("w:numFmt w:val=\"ordinal\"", numbering, StringComparison.Ordinal);
        Assert.Contains("w:lvlJc w:val=\"right\"", numbering, StringComparison.Ordinal);
        Assert.Single(result.Inventory.Lost);
    }
}
