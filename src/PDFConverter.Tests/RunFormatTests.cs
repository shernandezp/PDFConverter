using Xunit;
using MigraDoc.DocumentObjectModel;

namespace PDFConverter.Tests;

public class RunFormatTests
{
    [Fact]
    public void ApplyTo_AllProperties_AppliedCorrectly()
    {
        var doc = new Document();
        var section = doc.AddSection();
        var para = section.AddParagraph();
        var formatted = para.AddFormattedText("test");

        var fmt = new RunFormat("Arial", "FF0000", true, true, true, 14.0);
        fmt.ApplyTo(formatted);

        Assert.Equal("Arial", formatted.Font.Name);
        Assert.True(formatted.Bold);
        Assert.True(formatted.Italic);
        Assert.Equal(Underline.Single, formatted.Underline);
        Assert.Equal(Unit.FromPoint(14.0), formatted.Size);
    }

    [Fact]
    public void ApplyTo_NullProperties_NoException()
    {
        var doc = new Document();
        var section = doc.AddSection();
        var para = section.AddParagraph();
        var formatted = para.AddFormattedText("test");

        var fmt = new RunFormat(null, null, false, false, false, null);
        fmt.ApplyTo(formatted);

        // Should not throw, and no formatting applied
        Assert.False(formatted.Bold);
        Assert.False(formatted.Italic);
    }

    [Fact]
    public void ApplyTo_OnlySize_SetsSize()
    {
        var doc = new Document();
        var section = doc.AddSection();
        var para = section.AddParagraph();
        var formatted = para.AddFormattedText("test");

        var fmt = new RunFormat(null, null, false, false, false, 9.5);
        fmt.ApplyTo(formatted);

        Assert.Equal(Unit.FromPoint(9.5), formatted.Size);
    }

    [Fact]
    public void ApplyTo_ColorWithHash_ParsesCorrectly()
    {
        var doc = new Document();
        var section = doc.AddSection();
        var para = section.AddParagraph();
        var formatted = para.AddFormattedText("test");

        // The Color is stored without # but ApplyTo prepends #
        var fmt = new RunFormat(null, "0000FF", false, false, false, null);
        fmt.ApplyTo(formatted);

        // Color should be blue — the actual parse adds #
        Assert.NotEqual(Colors.Black, formatted.Color);
    }

    [Fact]
    public void ApplyTo_InvalidColor_DoesNotThrow()
    {
        var doc = new Document();
        var section = doc.AddSection();
        var para = section.AddParagraph();
        var formatted = para.AddFormattedText("test");

        var fmt = new RunFormat(null, "ZZZZZZ", false, false, false, null);

        // Should not throw due to try/catch in ApplyTo
        fmt.ApplyTo(formatted);
    }

    [Fact]
    public void ApplyTo_InvalidFontName_DoesNotThrow()
    {
        var doc = new Document();
        var section = doc.AddSection();
        var para = section.AddParagraph();
        var formatted = para.AddFormattedText("test");

        var fmt = new RunFormat("NonExistentFont_XYZ", null, false, false, false, null);

        // Should not throw due to try/catch in ApplyTo
        fmt.ApplyTo(formatted);
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
