using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Como se chega às caixas de texto de um parágrafo, e em que ordem.
/// </summary>
/// <remarks>
/// Leitor e escritor precisam ver **as mesmas caixas na mesma ordem**: o modelo
/// entrega os objetos de um bloco numa lista, e é por posição nessa lista que o
/// texto digitado volta para o `w:txbxContent` certo. Duas travessias parecidas
/// escritas em dois arquivos divergiriam, e a divergência escreveria o subtítulo
/// dentro do título.
/// </remarks>
internal static class TextBoxNav
{
    /// <summary>
    /// As caixas mais externas de um desenho.
    /// </summary>
    /// <remarks>
    /// Caixa dentro de caixa aparece: o conteúdo da de dentro já é lido junto
    /// com o da de fora, e descê-la de novo a mostraria duas vezes.
    /// </remarks>
    internal static IEnumerable<TextBoxContent> Outermost(OpenXmlElement root)
    {
        foreach (var child in root.ChildElements)
        {
            if (child is TextBoxContent box)
            {
                yield return box;
                continue;
            }

            foreach (var nested in Outermost(child)) yield return nested;
        }
    }

    /// <summary>
    /// O ramo que vale de um `mc:AlternateContent`.
    /// </summary>
    /// <remarks>
    /// O Word grava a mesma forma duas vezes — `mc:Choice` em DrawingML e
    /// `mc:Fallback` no VML antigo. Um ramo só, ou cada caixa aparece em dobro.
    /// </remarks>
    internal static OpenXmlElement? BranchOf(AlternateContent alternate) =>
        (OpenXmlElement?)alternate.GetFirstChild<AlternateContentChoice>()
        ?? alternate.GetFirstChild<AlternateContentFallback>();

    /// <summary>
    /// As caixas dos desenhos **posicionados** do parágrafo, na ordem do arquivo.
    /// </summary>
    /// <remarks>
    /// A mesma peneira que o leitor usa para decidir o que vira objeto na folha:
    /// desenho ancorado que não está onde o fluxo já o poria. Caixa de desenho
    /// no fluxo tem o texto lido na linha, não vira objeto, e não pode entrar
    /// nesta contagem sob pena de deslocar todas as outras.
    /// </remarks>
    internal static IEnumerable<TextBoxContent> AnchoredBoxesOf(OpenXmlElement paragraph)
    {
        foreach (var run in paragraph.Elements<Run>())
        {
            foreach (var child in run.ChildElements)
            {
                var shape = child is AlternateContent alternate ? BranchOf(alternate) : child;
                if (shape is null) continue;
                if (AnchorReader.AnchorOf(shape) is not { } anchor) continue;
                if (AnchorReader.FlowsWithText(anchor)) continue;

                foreach (var box in Outermost(shape)) yield return box;
            }
        }
    }

    /// <summary>Os parágrafos que são da caixa, e não de uma caixa de dentro.</summary>
    internal static IEnumerable<Paragraph> ParagraphsOf(TextBoxContent box) =>
        box.Descendants<Paragraph>().Where(p => p.Ancestors<TextBoxContent>().First() == box);

    /// <summary>
    /// O texto de uma caixa, para comparar com o que voltou do editor.
    /// </summary>
    /// <remarks>
    /// Comparar texto é o que deixa a caixa **não editada** intocada: o XML dela
    /// segue byte a byte, com a moldura, o preenchimento e a formatação que este
    /// escritor não sabe reproduzir.
    /// </remarks>
    internal static string TextOf(TextBoxContent box) =>
        string.Join("\n", ParagraphsOf(box).Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))));
}
