using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Leitor e escritor veem as mesmas caixas na mesma ordem: o texto volta ao
/// <c>w:txbxContent</c> pela posição na lista de objetos do bloco.
/// </summary>
internal static class TextBoxNav
{
    /// <summary>Caixa dentro de caixa já é lida com a de fora, e não desce de novo.</summary>
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
    /// O Word grava a forma em <c>mc:Choice</c> (DrawingML) e em <c>mc:Fallback</c>
    /// (VML): um ramo só, senão a caixa aparece em dobro.
    /// </summary>
    internal static OpenXmlElement? BranchOf(AlternateContent alternate) =>
        (OpenXmlElement?)alternate.GetFirstChild<AlternateContentChoice>()
        ?? alternate.GetFirstChild<AlternateContentFallback>();

    /// <summary>
    /// A mesma peneira do leitor: a caixa de desenho no fluxo é lida na linha e não
    /// entra na contagem, senão desloca as outras.
    /// </summary>
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

    /// <summary>
    /// O ramo VML repete o texto do que vale, senão o arquivo diz duas coisas e cada
    /// programa lê uma. A forma VML fica, porque é a moldura de quem lê esse ramo.
    /// </summary>
    internal static void MirrorFallback(AlternateContent alternate)
    {
        var choice = alternate.GetFirstChild<AlternateContentChoice>();
        var fallback = alternate.GetFirstChild<AlternateContentFallback>();
        if (choice is null || fallback is null) return;

        var source = Outermost(choice).ToList();
        var mirror = Outermost(fallback).ToList();
        if (source.Count != mirror.Count) return;

        for (var index = 0; index < source.Count; index++)
        {
            mirror[index].RemoveAllChildren();
            foreach (var child in source[index].ChildElements)
            {
                mirror[index].AppendChild(child.CloneNode(true));
            }
        }
    }

    /// <summary>Os parágrafos que são da caixa, e não de uma caixa de dentro.</summary>
    internal static IEnumerable<Paragraph> ParagraphsOf(TextBoxContent box) =>
        box.Descendants<Paragraph>().Where(p => p.Ancestors<TextBoxContent>().First() == box);

    /// <summary>A caixa de texto igual não é regravada, e guarda o que o escritor não reproduz.</summary>
    internal static string TextOf(TextBoxContent box) =>
        string.Join("\n", ParagraphsOf(box).Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))));
}
