using Librevia.Format.Odt;

namespace Librevia.Format.Tests;

public class TextEditsTests
{
    [Fact]
    public void EditsApplyFromTheEndSoEarlierOffsetsStayValid()
    {
        var edits = new TextEdits("abcdef");
        edits.Replace(0, 1, "X");
        edits.Insert(3, "-");
        edits.Replace(4, 6, "");

        Assert.Equal("Xbc-d", edits.Apply());
    }

    [Fact]
    public void OverlappingEditsAreRefusedInsteadOfCorruptingTheText()
    {
        var edits = new TextEdits("abcdef");
        edits.Replace(1, 4, "X");
        edits.Replace(3, 5, "Y");

        Assert.Throws<InvalidOperationException>(edits.Apply);
    }
}
