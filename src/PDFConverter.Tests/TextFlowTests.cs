using PDFConverter.Layout;
using SkiaSharp;
using Xunit;

namespace PDFConverter.Tests;

public class TextFlowTests
{
    static readonly TextStyle Style = new("Arial", 10);

    static TextFlowTests() => OpenXmlHelpers.EnsureFontResolverInitialized();

    static TextMeasurer Measurer() => new(FontRegistry.Current);

    static Paragraph ParagraphOf(string text)
    {
        var paragraph = new Paragraph { DefaultStyle = Style };
        paragraph.Add(new TextRun(text, Style));
        return paragraph;
    }

    static ImageSource Picture()
    {
        using var bitmap = new SKBitmap(4, 4);
        bitmap.Erase(SKColors.Blue);
        using var encoded = SKImage.FromBitmap(bitmap).Encode(SKEncodedImageFormat.Png, 100);
        return ImageSource.Create(encoded.ToArray())!;
    }

    static (Section Section, List<Page> Pages) LayoutBesidePicture(double pictureWidth, WrapSide side)
    {
        var document = new Document();
        var section = document.AddSection();
        section.Blocks.Add(new FloatingBlock
        {
            Image = Picture(), Left = 0, Top = 0, Width = pictureWidth, Height = 40, WrapSide = side,
            DistanceRight = 5, DistanceLeft = 5,
        });
        section.Blocks.Add(ParagraphOf(string.Join(" ", Enumerable.Repeat("flowing text", 80))));
        using var engine = new LayoutEngine(FontRegistry.Current);
        return (section, engine.Layout(document));
    }

    [Fact]
    public void SquareWrap_TextBesideThePicture_StartsAfterItAndReturnsBelow()
    {
        var (section, pages) = LayoutBesidePicture(100, WrapSide.Right);

        var ops = pages[0].Body.OfType<TextOp>().ToList();
        var beside = ops.Where(op => op.Baseline < section.TopMargin + 40).ToList();
        var below = ops.Where(op => op.Baseline > section.TopMargin + 40 + 12).ToList();
        Assert.NotEmpty(beside);
        Assert.NotEmpty(below);
        Assert.All(beside, op => Assert.True(op.X >= section.LeftMargin + 100 + 5, $"x={op.X}"));
        Assert.Contains(below, op => Math.Abs(op.X - section.LeftMargin) < 0.01);
    }

    [Fact]
    public void SquareWrap_LeftSide_KeepsTextLeftOfThePicture()
    {
        var document = new Document();
        var section = document.AddSection();
        section.Blocks.Add(new FloatingBlock
        {
            Image = Picture(), Left = section.ContentWidth - 100, Top = 0, Width = 100, Height = 40, WrapSide = WrapSide.Left,
        });
        section.Blocks.Add(ParagraphOf(string.Join(" ", Enumerable.Repeat("flowing text", 80))));
        using var engine = new LayoutEngine(FontRegistry.Current);
        var pages = engine.Layout(document);

        var beside = pages[0].Body.OfType<TextOp>().Where(op => op.Baseline < section.TopMargin + 40).ToList();
        Assert.NotEmpty(beside);
        var pictureLeft = section.LeftMargin + section.ContentWidth - 100;
        Assert.All(beside, op => Assert.True(op.X < pictureLeft, $"x={op.X} picture={pictureLeft}"));
    }

    [Fact]
    public void SquareWrap_WhenTooLittleRoomRemains_TextStartsBelowThePicture()
    {
        var (section, pages) = LayoutBesidePicture(section_width(), WrapSide.Both);

        var first = pages[0].Body.OfType<TextOp>().First();
        Assert.True(first.Baseline > section.TopMargin + 40, $"baseline={first.Baseline}");
    }

    static double section_width() => new Section().ContentWidth - 20;

    [Fact]
    public void Shaping_AppliesKerning()
    {
        var measurer = Measurer();
        var face = FontRegistry.Current.Resolve("Arial", false, false);

        var pair = measurer.Width(face, 12, "AV");
        var separate = measurer.Width(face, 12, "A") + measurer.Width(face, 12, "V");

        Assert.True(pair < separate - 0.3, $"pair={pair} separate={separate}");
    }

    [Fact]
    public void Shaping_ArabicLetters_TakeContextualForms()
    {
        var measurer = Measurer();
        var face = FontRegistry.Current.Resolve("Arial", false, false);

        var isolated = measurer.Shape(face, 12, "س", rtl: true).Glyphs;
        var word = measurer.Shape(face, 12, "سلام", rtl: true).Glyphs;

        Assert.InRange(word.Length, 3, 4);
        Assert.DoesNotContain(isolated[0], word);
        Assert.Equal(TextDirection.RightToLeft, TextLayout.DirectionOf("سلام"));
        Assert.Equal(TextDirection.LeftToRight, TextLayout.DirectionOf("salam"));
        Assert.Equal(TextDirection.Neutral, TextLayout.DirectionOf("--"));
    }

    static LaidOutLine SingleLine(Paragraph paragraph) =>
        Assert.Single(TextLayout.Layout(paragraph, 400, Measurer(), _ => string.Empty).Lines);

    [Fact]
    public void RightToLeftParagraph_PlacesLogicallyFirstTextOnTheRight()
    {
        var paragraph = new Paragraph { DefaultStyle = Style, RightToLeft = true };
        paragraph.Add(new TextRun("abc ", Style)).Add(new TextRun("אבג", Style));

        var items = SingleLine(paragraph).Items.Where(i => i.Kind == ItemKind.Text).ToList();

        var latin = items.Single(i => i.Text == "abc");
        var hebrew = items.Single(i => i.Text == "אבג");
        Assert.True(hebrew.X < latin.X, $"hebrew={hebrew.X} latin={latin.X}");
    }

    [Fact]
    public void LeftToRightParagraph_KeepsAnEmbeddedRightToLeftWordInPlace()
    {
        var paragraph = ParagraphOf("hello עברית world");

        var items = SingleLine(paragraph).Items.Where(i => i.Kind == ItemKind.Text).ToList();

        Assert.Equal(["hello", "עברית", "world"], items.Select(i => i.Text));
        Assert.True(items[0].X < items[1].X && items[1].X < items[2].X);
    }

    [Fact]
    public void RightToLeftParagraph_NeutralsBetweenTwoRightToLeftWords_StayBetweenThem()
    {
        var paragraph = new Paragraph { DefaultStyle = Style, RightToLeft = true };
        paragraph.Add(new TextRun("אבג - דהו", Style));

        var items = SingleLine(paragraph).Items.Where(i => i.Kind == ItemKind.Text).OrderBy(i => i.X).ToList();

        Assert.Equal(["דהו", "-", "אבג"], items.Select(i => i.Text));
    }

    [Fact]
    public void Writer_EmbedsShapedTextThatExtractsAsTheOriginalString()
    {
        var document = new Document();
        var section = document.AddSection();
        section.Blocks.Add(ParagraphOf("office AVAST"));
        using var engine = new LayoutEngine(FontRegistry.Current);
        var pages = engine.Layout(document);

        using var output = new MemoryStream();
        SkiaPdfWriter.Write(pages, output, null);

        Assert.Equal(["office", "AVAST"], PdfProbe.Words(output.ToArray()));
    }

    [Fact]
    public void Writer_EmbedsImagesLosslessly()
    {
        var document = new Document();
        var section = document.AddSection();
        section.Blocks.Add(new FloatingBlock { Image = Picture(), Left = 0, Top = 0, Width = 40, Height = 40 });
        using var engine = new LayoutEngine(FontRegistry.Current);
        var pages = engine.Layout(document);

        using var output = new MemoryStream();
        SkiaPdfWriter.Write(pages, output, null);
        var pdf = System.Text.Encoding.Latin1.GetString(output.ToArray());

        Assert.Single(PdfProbe.Images(output.ToArray()));
        Assert.DoesNotContain("/DCTDecode", pdf);
        Assert.Contains("/FlateDecode", pdf);
    }
}
