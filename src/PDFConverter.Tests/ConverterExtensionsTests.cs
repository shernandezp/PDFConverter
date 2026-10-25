using Xunit;

namespace PDFConverter.Tests;

public class ConverterExtensionsTests
{
    #region ToRoman

    [Theory]
    [InlineData(1, "I")]
    [InlineData(4, "IV")]
    [InlineData(9, "IX")]
    [InlineData(14, "XIV")]
    [InlineData(40, "XL")]
    [InlineData(99, "XCIX")]
    [InlineData(2024, "MMXXIV")]
    public void ToRoman_ConvertsCorrectly(int input, string expected)
    {
        Assert.Equal(expected, ConverterExtensions.ToRoman(input));
    }

    [Fact]
    public void ToRoman_ZeroOrNegative_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, ConverterExtensions.ToRoman(0));
        Assert.Equal(string.Empty, ConverterExtensions.ToRoman(-5));
    }

    #endregion

    #region DetectImageExtension

    [Fact]
    public void DetectImageExtension_PngHeader_ReturnsPng()
    {
        var bytes = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A };
        Assert.Equal(".png", ConverterExtensions.DetectImageExtension(bytes));
    }

    [Fact]
    public void DetectImageExtension_JpgHeader_ReturnsJpg()
    {
        var bytes = new byte[] { 0xFF, 0xD8, 0xFF, 0xE0 };
        Assert.Equal(".jpg", ConverterExtensions.DetectImageExtension(bytes));
    }

    [Fact]
    public void DetectImageExtension_GifHeader_ReturnsGif()
    {
        var bytes = new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 };
        Assert.Equal(".gif", ConverterExtensions.DetectImageExtension(bytes));
    }

    [Fact]
    public void DetectImageExtension_Unknown_DefaultsPng()
    {
        var bytes = new byte[] { 0x00, 0x01, 0x02, 0x03 };
        Assert.Equal(".png", ConverterExtensions.DetectImageExtension(bytes));
    }

    [Fact]
    public void DetectImageExtension_TooShort_DefaultsPng()
    {
        var bytes = new byte[] { 0x89, 0x50 };
        Assert.Equal(".png", ConverterExtensions.DetectImageExtension(bytes));
    }

    #endregion

    #region GetParagraphText

    [Fact]
    public void GetParagraphText_WithTextRunsAndTabs_ExtractsAll()
    {
        var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
            new DocumentFormat.OpenXml.Wordprocessing.Run(
                new DocumentFormat.OpenXml.Wordprocessing.Text("Hello")),
            new DocumentFormat.OpenXml.Wordprocessing.Run(
                new DocumentFormat.OpenXml.Wordprocessing.TabChar(),
                new DocumentFormat.OpenXml.Wordprocessing.Text("World")));

        var text = ConverterExtensions.GetParagraphText(p);

        Assert.Equal("Hello\tWorld", text);
    }

    [Fact]
    public void GetParagraphText_WithBreak_InsertsNewline()
    {
        var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
            new DocumentFormat.OpenXml.Wordprocessing.Run(
                new DocumentFormat.OpenXml.Wordprocessing.Text("Line1"),
                new DocumentFormat.OpenXml.Wordprocessing.Break(),
                new DocumentFormat.OpenXml.Wordprocessing.Text("Line2")));

        var text = ConverterExtensions.GetParagraphText(p);

        Assert.Equal("Line1\nLine2", text);
    }

    [Fact]
    public void GetParagraphText_EmptyParagraph_ReturnsEmpty()
    {
        var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
        Assert.Equal(string.Empty, ConverterExtensions.GetParagraphText(p));
    }

    [Fact]
    public void GetParagraphText_NullParagraph_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, ConverterExtensions.GetParagraphText(null!));
    }

    #endregion
}
