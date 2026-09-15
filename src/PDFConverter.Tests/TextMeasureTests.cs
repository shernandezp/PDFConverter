using Xunit;

namespace PDFConverter.Tests;

public class TextMeasureTests
{
    static readonly RunFormat Format = new(null, null, false, false, false, 10.0);

    [Fact]
    public void SplitOverlongWords_NoWidthLimit_ReturnsTextUnchanged()
    {
        Assert.Equal(["{{placeholder}}"], TextMeasure.SplitOverlongWords("{{placeholder}}", Format, 0));
    }

    [Fact]
    public void SplitOverlongWords_TextThatFits_IsNotSplit()
    {
        Assert.Equal(["ok"], TextMeasure.SplitOverlongWords("ok", Format, 200));
    }

    [Fact]
    public void SplitOverlongWords_WordWiderThanTheCell_IsBrokenUp()
    {
        OpenXmlHelpers.EnsureFontResolverInitialized();

        var lines = TextMeasure.SplitOverlongWords("{{minimum}}", Format, 30);

        Assert.True(lines.Count > 1, "expected the word to be broken across lines");
        Assert.Equal("{{minimum}}", string.Concat(lines));
    }

    [Fact]
    public void SplitOverlongWords_KeepsShortWordsIntact()
    {
        OpenXmlHelpers.EnsureFontResolverInitialized();

        var lines = TextMeasure.SplitOverlongWords("a b c", Format, 30);

        Assert.Equal(["a b c"], lines);
    }
}
