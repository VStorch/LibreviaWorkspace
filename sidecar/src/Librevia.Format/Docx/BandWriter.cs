using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O texto digitado no cabeçalho ou no rodapé volta para a parte que o guarda.
/// </summary>
/// <remarks>
/// A gravação cirúrgica deixava `word/header1.xml` intacto porque ninguém o
/// abria; a faixa era desenho, não edição. Abrir a edição não muda a aposta —
/// muda só o que conta como "tocado". O que se escreve aqui é **um `w:t` por
/// peça editada**, e nada mais: a moldura, a tabela, o logotipo, o campo `PAGE`
/// e o parágrafo que os carrega seguem byte a byte como estavam.
///
/// É por isso que o endereço da peça existe. Comparar a faixa inteira não daria
/// para decidir o que mudou — ela é uma projeção com perda, e a volta seria uma
/// regeneração disfarçada de edição, que é justamente o que este projeto não
/// faz.
///
/// Parte cujo texto não mudou **não é tocada**, e nem sequer entra na lista de
/// partes graváveis: o SDK reserializa toda parte cujo DOM tipado foi
/// materializado, e sair diferente sem intenção é a porta por onde a fidelidade
/// escapa em silêncio.
/// </remarks>
internal static class BandWriter
{
    /// <summary>
    /// Aplica o texto editado das faixas e devolve os caminhos que mudaram.
    /// </summary>
    internal static HashSet<string> Apply(MainDocumentPart part, PageSetupDto? page, Inventory inventory)
    {
        var touched = new HashSet<string>(StringComparer.Ordinal);
        if (page is null) return touched;

        // Endereço → texto que se quer ali. Um mesmo cabeçalho aparece em várias
        // folhas mas é um só no arquivo, e o modelo o traz uma vez por papel:
        // se dois papéis apontarem para a mesma parte, o texto é o mesmo.
        var wanted = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var band in BandsOf(page))
        {
            foreach (var piece in PiecesOf(band)) Remember(wanted, piece);
        }

        // O mesmo, um nível acima: o cabeçalho corporativo não é feito de
        // parágrafos soltos, e o título dele mora dentro de uma caixa.
        var boxes = new Dictionary<string, List<Node>>(StringComparer.Ordinal);
        foreach (var band in BandsOf(page))
        {
            foreach (var float_ in band.Floats ?? [])
            {
                if (float_.BoxId is null || float_.Kind != "text") continue;
                boxes[float_.BoxId] = float_.Content ?? [];
            }
        }

        if (wanted.Count == 0 && boxes.Count == 0) return touched;

        var fonts = new FontTable(part);
        var relationships = wanted.Keys.Select(Relationship)
            .Concat(boxes.Keys.Select(BoxRelationship))
            .Where(id => id.Length > 0)
            .Distinct(StringComparer.Ordinal);

        foreach (var relationship in relationships)
        {
            if (BandNav.PartOf(part, relationship) is not { } target) continue;

            // A mesma tabela de fontes do leitor, e não outra: é ela que decide
            // se dois runs vizinhos são a mesma peça, e uma fusão diferente aqui
            // daria outra numeração de peças — o texto iria para a peça errada.
            var changed = ApplyTo(
                target.Root,
                wanted.Where(entry => Relationship(entry.Key) == relationship),
                inventory,
                fonts);

            changed |= ApplyBoxes(
                target.Root,
                boxes.Where(entry => BoxRelationship(entry.Key) == relationship),
                part,
                inventory,
                fonts);

            if (!changed) continue;

            target.Root.Save();
            touched.Add(BandNav.PathOf(target.Owner));
        }

        return touched;
    }

    /// <summary>
    /// Regenera as caixas cujo texto mudou, e só essas.
    /// </summary>
    /// <remarks>
    /// A caixa vem inteira porque digitar dentro dela abre e fecha parágrafos:
    /// um endereço por parágrafo quebraria no primeiro Enter. Caixa cujo texto
    /// não mudou não é tocada — o XML dela segue como estava, com a moldura, o
    /// preenchimento e o giro que este escritor não sabe reproduzir.
    /// </remarks>
    private static bool ApplyBoxes(
        OpenXmlPartRootElement root,
        IEnumerable<KeyValuePair<string, List<Node>>> wanted,
        MainDocumentPart part,
        Inventory inventory,
        FontTable fonts)
    {
        var boxes = BandNav.BoxesOf(root);
        var writer = new ParagraphWriter(part, inventory);
        var touched = false;

        foreach (var (address, content) in wanted.Select(entry => (entry.Key, entry.Value)))
        {
            if (BandNav.ParseBox(address) is not { } at) continue;
            if (at.Box >= boxes.Count) continue;

            var box = boxes[at.Box];

            // Contra o texto que a tela mostra, e não contra o do arquivo: o
            // campo `PAGE` sai do leitor como `{n}`, e comparar com o XML diria
            // que a caixa mudou sempre — a gravação trocaria o campo por um
            // `{n}` literal, que foi o que o cabeçalho do corpus passou a
            // mostrar no lugar do número da página.
            if (HeaderReader.BoxTextOf(box, inventory, fonts) == PlainTextOf(content)) continue;

            // Caixa com campo dentro não é reescrita nem quando o texto mudou:
            // regenerá-la apagaria o campo, e um cabeçalho que deixa de contar
            // páginas é pior do que um título que não se pôde corrigir.
            if (HasField(box))
            {
                inventory.NoteLoss("texto de uma caixa de cabeçalho com campo calculado");
                continue;
            }

            box.RemoveAllChildren();
            foreach (var block in content)
            {
                foreach (var element in writer.Write(block)) box.AppendChild(element);
            }

            // `w:txbxContent` vazio invalida o documento para o Word.
            if (!box.HasChildren) box.AppendChild(new Paragraph());
            touched = true;
        }

        if (touched) Mirror(root);
        return touched;
    }

    private static bool HasField(TextBoxContent box) =>
        box.Descendants<FieldChar>().Any()
        || box.Descendants<FieldCode>().Any()
        || box.Descendants<SimpleField>().Any();

    /// <summary>O texto de um conteúdo de caixa, na mesma forma que o do arquivo.</summary>
    private static string PlainTextOf(List<Node> content) =>
        string.Join("\n", content.Select(node => string.Concat(Texts(node))));

    private static IEnumerable<string> Texts(Node node)
    {
        if (node.Text is { } text) yield return text;
        foreach (var child in node.Content ?? [])
        {
            foreach (var value in Texts(child)) yield return value;
        }
    }

    private static string BoxRelationship(string address) =>
        BandNav.ParseBox(address) is { } parsed ? parsed.RelationshipId : string.Empty;

    private static IEnumerable<BandDto> BandsOf(PageSetupDto page)
    {
        foreach (var band in new[]
                 {
                     page.Header, page.Footer,
                     page.FirstHeader, page.FirstFooter,
                     page.EvenHeader, page.EvenFooter,
                 })
        {
            if (band is not null) yield return band;
        }
    }

    private static IEnumerable<PieceDto> PiecesOf(BandDto band)
    {
        foreach (var piece in band.Left) yield return piece;
        foreach (var piece in band.Center) yield return piece;
        foreach (var piece in band.Right) yield return piece;

        foreach (var row in band.Rows ?? [])
        {
            foreach (var cell in row.Cells)
            {
                foreach (var piece in cell.Pieces) yield return piece;
            }
        }
    }

    private static void Remember(Dictionary<string, string> wanted, PieceDto piece)
    {
        if (piece.Pid is null || piece.Kind != PieceDto.KindText) return;
        wanted[piece.Pid] = piece.Text ?? string.Empty;
    }

    private static string Relationship(string address) =>
        BandNav.Parse(address) is { } parsed ? parsed.RelationshipId : string.Empty;

    /// <summary>
    /// Escreve numa parte o que mudou nela. Devolve se alguma coisa mudou mesmo.
    /// </summary>
    private static bool ApplyTo(
        OpenXmlPartRootElement root,
        IEnumerable<KeyValuePair<string, string>> wanted,
        Inventory inventory,
        FontTable fonts)
    {
        var paragraphs = BandNav.ParagraphsOf(root);
        var cache = new Dictionary<int, List<HeaderReader.TracedPiece>>();
        var touched = false;

        foreach (var (address, text) in wanted.Select(entry => (entry.Key, entry.Value)))
        {
            if (BandNav.Parse(address) is not { } at) continue;
            if (at.Paragraph >= paragraphs.Count) continue;

            if (!cache.TryGetValue(at.Paragraph, out var pieces))
            {
                pieces = HeaderReader.TracedRuns(paragraphs[at.Paragraph], inventory, fonts);
                cache[at.Paragraph] = pieces;
            }

            if (at.Piece >= pieces.Count) continue;

            var piece = pieces[at.Piece];
            if (piece.Source.Count == 0) continue;
            if (piece.Piece.Text == text) continue;

            Rewrite(piece.Source, text);
            touched = true;
        }

        if (!touched) return false;

        Mirror(root);
        return true;
    }

    /// <summary>
    /// O ramo de reserva repete a mesma caixa em VML.
    /// </summary>
    /// <remarks>
    /// Escrever só no ramo que vale deixaria o arquivo dizendo duas coisas, e
    /// qual delas aparece depende de quem abre.
    /// </remarks>
    private static void Mirror(OpenXmlPartRootElement root)
    {
        foreach (var alternate in root.Descendants<AlternateContent>().ToList())
        {
            TextBoxNav.MirrorFallback(alternate);
        }
    }

    /// <summary>
    /// O texto novo entra no primeiro `w:t` da peça; os demais esvaziam.
    /// </summary>
    /// <remarks>
    /// A peça pode ter saído de vários runs — o Word pica uma frase de mesmo
    /// estilo em pedaços — e o leitor os fundiu numa peça só. Esvaziar em vez de
    /// remover mantém a contagem de runs, e com ela o endereço, estável de uma
    /// gravação para a seguinte.
    ///
    /// `xml:space="preserve"` sempre: sem ele o Word come o espaço da ponta, e
    /// "Manual do " voltaria como "Manual do" colado no que vem depois.
    /// </remarks>
    private static void Rewrite(List<Text> source, string text)
    {
        source[0].Text = text;
        source[0].Space = SpaceProcessingModeValues.Preserve;

        for (var index = 1; index < source.Count; index++) source[index].Text = string.Empty;
    }
}
