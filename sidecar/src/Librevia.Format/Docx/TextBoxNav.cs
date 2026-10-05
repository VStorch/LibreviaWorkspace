using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Reader and writer see the same boxes in the same order: text goes back to <c>w:txbxContent</c>
/// by its position in the block's object list.
/// </summary>
internal static class TextBoxNav
{
    /// <summary>
    /// A box inside a box is already read with the outer one, and is not descended into again.
    /// </summary>
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
    /// Word writes the shape in <c>mc:Choice</c> (DrawingML) and in <c>mc:Fallback</c> (VML): one
    /// branch only, or the box appears twice.
    /// </summary>
    internal static OpenXmlElement? BranchOf(AlternateContent alternate) =>
        (OpenXmlElement?)alternate.GetFirstChild<AlternateContentChoice>()
        ?? alternate.GetFirstChild<AlternateContentFallback>();

    /// <summary>
    /// The same sieve as the reader: a drawing box in the flow is read in the line and does not
    /// enter the count, or it would shift the others.
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
    /// The VML branch repeats the text of the one that counts, or the file says two things and each
    /// program reads one. The VML shape stays, because it is the frame for whoever reads that
    /// branch.
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

    /// <summary>Paragraphs belonging to the box, not to a box inside it.</summary>
    internal static IEnumerable<Paragraph> ParagraphsOf(TextBoxContent box) =>
        box.Descendants<Paragraph>().Where(p => p.Ancestors<TextBoxContent>().First() == box);

    /// <summary>
    /// An equal text box is not rewritten, and keeps what the writer does not reproduce.
    /// </summary>
    internal static string TextOf(TextBoxContent box) =>
        string.Join("\n", ParagraphsOf(box).Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))));
}
