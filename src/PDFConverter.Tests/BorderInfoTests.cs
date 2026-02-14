using Xunit;

namespace PDFConverter.Tests;

public class BorderInfoTests
{
    [Fact]
    public void Empty_HasAllZeroWidths()
    {
        var empty = BorderInfo.Empty;

        Assert.Equal(0, empty.TopWidth);
        Assert.Equal(0, empty.BottomWidth);
        Assert.Equal(0, empty.LeftWidth);
        Assert.Equal(0, empty.RightWidth);
        Assert.Null(empty.TopColor);
        Assert.Null(empty.BottomColor);
        Assert.Null(empty.LeftColor);
        Assert.Null(empty.RightColor);
        Assert.Null(empty.TopStyle);
        Assert.Null(empty.BottomStyle);
        Assert.Null(empty.LeftStyle);
        Assert.Null(empty.RightStyle);
    }

    [Fact]
    public void Empty_HasZeroPadding()
    {
        var empty = BorderInfo.Empty;

        Assert.Equal(0, empty.PaddingTop);
        Assert.Equal(0, empty.PaddingBottom);
        Assert.Equal(0, empty.PaddingLeft);
        Assert.Equal(0, empty.PaddingRight);
    }

    [Fact]
    public void Constructor_SetsAllProperties()
    {
        var border = new BorderInfo(
            1.0, "#000000", "Single",
            2.0, "#FF0000", "Double",
            0.5, "#00FF00", "Dotted",
            1.5, "#0000FF", "Dashed",
            3.0, 4.0, 5.0, 6.0);

        Assert.Equal(1.0, border.TopWidth);
        Assert.Equal("#000000", border.TopColor);
        Assert.Equal("Single", border.TopStyle);
        Assert.Equal(2.0, border.BottomWidth);
        Assert.Equal("#FF0000", border.BottomColor);
        Assert.Equal("Double", border.BottomStyle);
        Assert.Equal(0.5, border.LeftWidth);
        Assert.Equal("#00FF00", border.LeftColor);
        Assert.Equal("Dotted", border.LeftStyle);
        Assert.Equal(1.5, border.RightWidth);
        Assert.Equal("#0000FF", border.RightColor);
        Assert.Equal("Dashed", border.RightStyle);
        Assert.Equal(3.0, border.PaddingTop);
        Assert.Equal(4.0, border.PaddingBottom);
        Assert.Equal(5.0, border.PaddingLeft);
        Assert.Equal(6.0, border.PaddingRight);
    }

    [Fact]
    public void RecordEquality_Works()
    {
        var a = new BorderInfo(1, "#000", "Single", 1, "#000", "Single",
            1, "#000", "Single", 1, "#000", "Single", 2, 2, 2, 2);
        var b = new BorderInfo(1, "#000", "Single", 1, "#000", "Single",
            1, "#000", "Single", 1, "#000", "Single", 2, 2, 2, 2);

        Assert.Equal(a, b);
    }

    [Fact]
    public void DefaultPadding_IsZero()
    {
        // Constructor without padding uses default 0 values
        var border = new BorderInfo(1, null, null, 1, null, null,
            1, null, null, 1, null, null);

        Assert.Equal(0, border.PaddingTop);
        Assert.Equal(0, border.PaddingBottom);
        Assert.Equal(0, border.PaddingLeft);
        Assert.Equal(0, border.PaddingRight);
    }

    [Fact]
    public void Empty_IsSingleton()
    {
        Assert.Same(BorderInfo.Empty, BorderInfo.Empty);
    }
}
