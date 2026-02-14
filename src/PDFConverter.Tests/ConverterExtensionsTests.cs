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

    #region ContainsEmoji

    [Fact]
    public void ContainsEmoji_PlainText_ReturnsFalse()
    {
        Assert.False(ConverterExtensions.ContainsEmoji("Hello World"));
    }

    [Fact]
    public void ContainsEmoji_BmpEmoji_ReturnsTrue()
    {
        // ☀ = U+2600 (Misc symbols range)
        Assert.True(ConverterExtensions.ContainsEmoji("Sunny ☀"));
    }

    [Fact]
    public void ContainsEmoji_SurrogatePairEmoji_ReturnsTrue()
    {
        // 😀 = U+1F600 (requires surrogate pair)
        Assert.True(ConverterExtensions.ContainsEmoji("Happy 😀"));
    }

    [Fact]
    public void ContainsEmoji_ZeroWidthJoiner_ReturnsTrue()
    {
        // ZWJ = U+200D
        Assert.True(ConverterExtensions.ContainsEmoji("text\u200Dtext"));
    }

    [Fact]
    public void ContainsEmoji_VariationSelector_ReturnsTrue()
    {
        // U+FE0F = variation selector-16
        Assert.True(ConverterExtensions.ContainsEmoji("star\uFE0F"));
    }

    [Fact]
    public void ContainsEmoji_EmptyString_ReturnsFalse()
    {
        Assert.False(ConverterExtensions.ContainsEmoji(""));
    }

    #endregion

    #region SplitEmojiSegments

    [Fact]
    public void SplitEmojiSegments_PlainText_SingleNonEmojiSegment()
    {
        var segments = ConverterExtensions.SplitEmojiSegments("Hello");

        Assert.Single(segments);
        Assert.Equal("Hello", segments[0].Text);
        Assert.False(segments[0].IsEmoji);
    }

    [Fact]
    public void SplitEmojiSegments_MixedContent_SplitsCorrectly()
    {
        // ☺ = U+263A (BMP emoji)
        var segments = ConverterExtensions.SplitEmojiSegments("Hi ☺ there");

        Assert.Equal(3, segments.Count);
        Assert.Equal("Hi ", segments[0].Text);
        Assert.False(segments[0].IsEmoji);
        Assert.Equal("☺", segments[1].Text);
        Assert.True(segments[1].IsEmoji);
        Assert.Equal(" there", segments[2].Text);
        Assert.False(segments[2].IsEmoji);
    }

    [Fact]
    public void SplitEmojiSegments_SurrogatePairEmoji_KeptTogether()
    {
        // 😀 = U+1F600 (2 chars in UTF-16)
        var input = "A😀B";
        var segments = ConverterExtensions.SplitEmojiSegments(input);

        Assert.Equal(3, segments.Count);
        Assert.Equal("A", segments[0].Text);
        Assert.False(segments[0].IsEmoji);
        Assert.Equal("😀", segments[1].Text);
        Assert.True(segments[1].IsEmoji);
        Assert.Equal(2, segments[1].Text.Length); // Surrogate pair = 2 chars
        Assert.Equal("B", segments[2].Text);
        Assert.False(segments[2].IsEmoji);
    }

    [Fact]
    public void SplitEmojiSegments_Empty_ReturnsEmpty()
    {
        Assert.Empty(ConverterExtensions.SplitEmojiSegments(""));
        Assert.Empty(ConverterExtensions.SplitEmojiSegments(null!));
    }

    [Fact]
    public void SplitEmojiSegments_ConsecutiveEmojis_GroupedTogether()
    {
        // ☀☁ = two BMP emojis in sequence
        var segments = ConverterExtensions.SplitEmojiSegments("☀☁");

        Assert.Single(segments);
        Assert.True(segments[0].IsEmoji);
        Assert.Equal("☀☁", segments[0].Text);
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

    #region ApplySrcRectCrop

    [Fact]
    public void ApplySrcRectCrop_AllZeros_ReturnsSameBytes()
    {
        var bytes = new byte[] { 1, 2, 3, 4 };
        var result = ConverterExtensions.ApplySrcRectCrop(bytes, 0, 0, 0, 0);
        Assert.Same(bytes, result); // Should return exact same reference
    }

    [Fact]
    public void ApplySrcRectCrop_InvalidImage_ReturnsOriginal()
    {
        var bytes = new byte[] { 1, 2, 3, 4 }; // Not a valid image
        var result = ConverterExtensions.ApplySrcRectCrop(bytes, 10000, 10000, 10000, 10000);
        Assert.Same(bytes, result); // Falls back to original on error
    }

    #endregion

    #region SaveTempImage / TryDeleteTempFile

    [Fact]
    public void SaveTempImage_CreatesFile_AndTryDeleteRemovesIt()
    {
        var pngHeader = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                                     0, 0, 0, 0, 0, 0, 0, 0 };
        var path = ConverterExtensions.SaveTempImage(pngHeader);

        try
        {
            Assert.True(File.Exists(path));
            Assert.EndsWith(".png", path);
            Assert.Equal(pngHeader, File.ReadAllBytes(path));
        }
        finally
        {
            ConverterExtensions.TryDeleteTempFile(path);
            Assert.False(File.Exists(path));
        }
    }

    [Fact]
    public void TryDeleteTempFile_NonExistentPath_DoesNotThrow()
    {
        // Should not throw for non-existent file
        ConverterExtensions.TryDeleteTempFile(@"C:\nonexistent_path_12345.tmp");
    }

    #endregion
}
