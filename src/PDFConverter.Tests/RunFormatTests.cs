using PDFConverter.Layout;
using SkiaSharp;
using Xunit;

namespace PDFConverter.Tests;

public class RunFormatTests
{
    static readonly TextStyle Defaults = new("Calibri", 11, Color: SKColors.Black);

    [Fact]
    public void ToTextStyle_AllProperties_AppliedCorrectly()
    {
        var fmt = new RunFormat("Arial", "FF0000", true, true, true, 14.0);
        var style = fmt.ToTextStyle(Defaults);

        Assert.Equal("Arial", style.FontFamily);
        Assert.True(style.Bold);
        Assert.True(style.Italic);
        Assert.Equal(UnderlineStyle.Single, style.Underline);
        Assert.Equal(14.0, style.Size);
        Assert.Equal(SKColors.Red, style.Color);
    }

    [Fact]
    public void ToTextStyle_NullProperties_FallBackToDefaults()
    {
        var fmt = new RunFormat(null, null, false, false, false, null);
        var style = fmt.ToTextStyle(Defaults);

        Assert.Equal("Calibri", style.FontFamily);
        Assert.Equal(11, style.Size);
        Assert.False(style.Bold);
        Assert.False(style.Italic);
        Assert.Null(style.Underline);
        Assert.Equal(SKColors.Black, style.Color);
    }

    [Fact]
    public void ToTextStyle_OnlySize_SetsSize()
    {
        var style = new RunFormat(null, null, false, false, false, 9.5).ToTextStyle(Defaults);

        Assert.Equal(9.5, style.Size);
    }

    [Fact]
    public void ToTextStyle_ColorWithoutHash_ParsesCorrectly()
    {
        var style = new RunFormat(null, "0000FF", false, false, false, null).ToTextStyle(Defaults);

        Assert.Equal(SKColors.Blue, style.Color);
    }

    [Fact]
    public void ToTextStyle_InvalidColor_KeepsDefault()
    {
        var style = new RunFormat(null, "ZZZZZZ", false, false, false, null).ToTextStyle(Defaults);

        Assert.Equal(SKColors.Black, style.Color);
    }

    [Fact]
    public void ToTextStyle_UnderlineStyle_IsCarriedOver()
    {
        var fmt = new RunFormat(null, null, false, false, true, null, UnderlineStyle: UnderlineStyle.Dotted);

        Assert.Equal(UnderlineStyle.Dotted, fmt.ToTextStyle(Defaults).Underline);
    }

    [Fact]
    public void RecordEquality_SameValues_AreEqual()
    {
        var a = new RunFormat("Arial", "FF0000", true, false, true, 12.0);
        var b = new RunFormat("Arial", "FF0000", true, false, true, 12.0);

        Assert.Equal(a, b);
    }

    [Fact]
    public void RecordEquality_DifferentValues_AreNotEqual()
    {
        var a = new RunFormat("Arial", "FF0000", true, false, true, 12.0);
        var b = new RunFormat("Verdana", "FF0000", true, false, true, 12.0);

        Assert.NotEqual(a, b);
    }
}
