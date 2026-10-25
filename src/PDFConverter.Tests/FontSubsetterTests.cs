using PDFConverter.Layout;
using SkiaSharp;
using Xunit;

namespace PDFConverter.Tests;

public class FontSubsetterTests
{
    static byte[] EmojiFont()
    {
        using var stream = typeof(FontRegistry).Assembly.GetManifestResourceStream("PDFConverter.Fonts.NotoEmoji-Regular.ttf")!;
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    [Fact]
    public void StripUnusedGlyphs_KeepsGlyphIdsAndShrinksTheFont()
    {
        var original = EmojiFont();
        using var typeface = SKTypeface.FromData(SKData.CreateCopy(original));
        using var font = new SKFont(typeface);
        var used = new HashSet<ushort>(font.GetGlyphs("😀☀"));

        var stripped = FontSubsetter.StripUnusedGlyphs(original, used);

        Assert.NotNull(stripped);
        Assert.True(stripped.Length < original.Length / 4, $"{stripped.Length} vs {original.Length}");

        using var subset = SKTypeface.FromData(SKData.CreateCopy(stripped));
        Assert.NotNull(subset);
        Assert.Equal(typeface.GlyphCount, subset.GlyphCount);
        using var subsetFont = new SKFont(subset);
        Assert.Equal(font.GetGlyphs("😀☀"), subsetFont.GetGlyphs("😀☀"));
        Assert.Equal(font.MeasureText("😀☀"), subsetFont.MeasureText("😀☀"), 0.01);
    }

    [Fact]
    public void StripUnusedGlyphs_RejectsNonTrueTypeData()
    {
        Assert.Null(FontSubsetter.StripUnusedGlyphs([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], new HashSet<ushort>()));
        Assert.Null(FontSubsetter.StripUnusedGlyphs("OTTO"u8.ToArray().Concat(new byte[64]).ToArray(), new HashSet<ushort>()));
    }
}
