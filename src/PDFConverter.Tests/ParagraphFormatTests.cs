using Xunit;
using MigraDoc.DocumentObjectModel;

namespace PDFConverter.Tests;

public class ParagraphFormatTests
{
    [Fact]
    public void ParagraphFormat_DefaultValues_AllZeroOrNull()
    {
        var fmt = new ParagraphFormat(
            ParagraphAlignment.Left, 0, 0, 0, 0, 0, null, null, false, false);

        Assert.Equal(ParagraphAlignment.Left, fmt.Alignment);
        Assert.Equal(0, fmt.LeftIndent);
        Assert.Equal(0, fmt.RightIndent);
        Assert.Equal(0, fmt.FirstLineIndent);
        Assert.Equal(0, fmt.SpacingBefore);
        Assert.Equal(0, fmt.SpacingAfter);
        Assert.Null(fmt.LineSpacing);
        Assert.Null(fmt.LineRule);
        Assert.False(fmt.HasExplicitSpacingBefore);
        Assert.False(fmt.HasExplicitSpacingAfter);
    }

    [Fact]
    public void ParagraphFormat_RightIndent_IsTracked()
    {
        var fmt = new ParagraphFormat(
            ParagraphAlignment.Right, 10.0, 5.0, 2.0, 6.0, 3.0, 1.15, "Auto", true, true);

        Assert.Equal(5.0, fmt.RightIndent);
    }

    [Fact]
    public void ParagraphFormat_RecordEquality()
    {
        var a = new ParagraphFormat(ParagraphAlignment.Center, 10, 5, 2, 6, 3, 1.5, "Auto", true, false);
        var b = new ParagraphFormat(ParagraphAlignment.Center, 10, 5, 2, 6, 3, 1.5, "Auto", true, false);

        Assert.Equal(a, b);
    }

    [Fact]
    public void ParagraphFormat_RecordInequality_DifferentRightIndent()
    {
        var a = new ParagraphFormat(ParagraphAlignment.Left, 0, 5.0, 0, 0, 0, null, null, false, false);
        var b = new ParagraphFormat(ParagraphAlignment.Left, 0, 10.0, 0, 0, 0, null, null, false, false);

        Assert.NotEqual(a, b);
    }

    [Fact]
    public void ParagraphFormat_RecordInequality_DifferentAlignment()
    {
        var a = new ParagraphFormat(ParagraphAlignment.Left, 0, 0, 0, 0, 0, null, null, false, false);
        var b = new ParagraphFormat(ParagraphAlignment.Center, 0, 0, 0, 0, 0, null, null, false, false);

        Assert.NotEqual(a, b);
    }

    [Fact]
    public void ParagraphFormat_NegativeFirstLineIndent_Allowed()
    {
        // Hanging indent produces negative FirstLineIndent
        var fmt = new ParagraphFormat(
            ParagraphAlignment.Left, 36.0, 0, -21.1, 0, 0, null, null, false, false);

        Assert.Equal(-21.1, fmt.FirstLineIndent, 1);
    }
}
