using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// Listas multinível: a definição que o leitor entrega à tela e o que a gravação
/// faz com ela — nada, quando nada mudou; um `w:num` novo, quando a lista
/// recomeça; uma definição por lista, quando a lista nasceu no editor.
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

        // Um, a sublista, Três (mesmo `numId`, do outro lado do parágrafo) e Dez.
        Assert.Equal(4, lists.Count);

        var first = DefinitionOf(lists[0]);
        Assert.Equal("a3", first["key"]!.GetValue<string>());
        Assert.Equal("lowerLetter", first["levels"]![1]!["fmt"]!.GetValue<string>());
        Assert.Equal("%1.%2)", first["levels"]![1]!["text"]!.GetValue<string>());

        // A sublista é da mesma numeração: acha a definição na lista de fora.
        Assert.Null(AttrOf(lists[1], "numbering"));
        Assert.Equal(5, IntOf(lists[1], "numId"));

        // O mesmo `numId` depois do parágrafo continua a mesma conta.
        Assert.Equal("a3", DefinitionOf(lists[2])["key"]!.GetValue<string>());

        // O `w:num` com reinício conta à parte, a partir de 10.
        var restarted = DefinitionOf(lists[3]);
        Assert.Equal("n6", restarted["key"]!.GetValue<string>());
        Assert.Equal(10, restarted["overrides"]!["0"]!.GetValue<int>());
    }

    [Fact]
    public void OutraNumeracaoNoMesmoNivelEOutraLista()
    {
        // Juntadas, "Dez" era contado como item de "Três" — 4, e não 10.
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
        // "Reiniciar em 1" no editor: a lista perde o `numId` e leva a definição
        // de onde saiu com o reinício. O Word grava exatamente isto — um `w:num`
        // novo, da mesma definição abstrata, com `w:startOverride`.
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
        // O item não muda — muda a lista em volta. Devolvido como veio, voltava
        // ao nível 0 ao reabrir.
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
        // Com uma definição só para todas, a segunda lista nova continuava a conta
        // da primeira ao reabrir no Word: 1, 2 e 3 onde a tela mostrou 1, 2 e 1.
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

        // E os níveis que voltam são os padrão que a tela desenhou.
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
        // "Continuar numeração" entre duas listas que ainda não foram gravadas: a
        // chave é o que as junta, e as duas têm de sair no mesmo `w:num`.
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
        // A definição escolhida no editor (galeria, lista colada de outro
        // documento) é gravada como veio — e volta igual.
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
}
