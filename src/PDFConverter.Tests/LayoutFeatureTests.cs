using PDFConverter.Layout;
using SkiaSharp;
using Xunit;

namespace PDFConverter.Tests;

public class LayoutFeatureTests
{
    static readonly TextStyle Style = new("Arial", 10);

    static LayoutFeatureTests() => OpenXmlHelpers.EnsureFontResolverInitialized();

    static TextMeasurer Measurer() => new(FontRegistry.Current);

    static Paragraph ParagraphOf(string text, TextStyle? style = null)
    {
        var paragraph = new Paragraph { DefaultStyle = Style };
        paragraph.Add(new TextRun(text, style ?? Style));
        return paragraph;
    }

    static List<DrawOp> EmitSingleLine(Paragraph paragraph)
    {
        var measurer = Measurer();
        var line = Assert.Single(TextLayout.Layout(paragraph, 400, measurer, _ => string.Empty).Lines);
        var ops = new List<DrawOp>();
        TextLayout.Emit(line, 0, 0, _ => string.Empty, measurer, ops);
        return ops;
    }

    [Fact]
    public void LetterSpacing_WidensTheRunByOneSpacingPerCharacter()
    {
        var plain = TextLayout.Layout(ParagraphOf("abcd"), 400, Measurer(), _ => string.Empty).Lines[0].Width;
        var spaced = TextLayout.Layout(ParagraphOf("abcd", Style with { LetterSpacing = 2 }), 400, Measurer(), _ => string.Empty)
            .Lines[0].Width;

        Assert.Equal(plain + 8, spaced, 0.01);
    }

    [Fact]
    public void WidthScale_ScalesTheRunWidth()
    {
        var plain = TextLayout.Layout(ParagraphOf("abcd"), 400, Measurer(), _ => string.Empty).Lines[0].Width;
        var narrow = TextLayout.Layout(ParagraphOf("abcd", Style with { WidthScale = 0.5 }), 400, Measurer(), _ => string.Empty)
            .Lines[0].Width;

        Assert.Equal(plain / 2, narrow, 0.1);
    }

    [Fact]
    public void Strike_DrawsALineAboveTheBaseline()
    {
        var ops = EmitSingleLine(ParagraphOf("struck", Style with { Strike = StrikeStyle.Single }));

        var text = ops.OfType<TextOp>().Single();
        var strike = ops.OfType<LineOp>().Single();
        Assert.True(strike.Y1 < text.Baseline && strike.Y1 > text.Baseline - 10);
        Assert.Equal(text.Width(), strike.X2 - strike.X1, 0.01);
    }

    [Fact]
    public void Highlight_FillsBehindTheTextBeforeDrawingIt()
    {
        var ops = EmitSingleLine(ParagraphOf("marked", Style with { Highlight = SKColors.Yellow }));

        Assert.IsType<FillRectOp>(ops[0]);
        Assert.IsType<TextOp>(ops[1]);
        Assert.Equal(SKColors.Yellow, ((FillRectOp)ops[0]).Color);
    }

    [Fact]
    public void ParagraphBorders_SurroundTheParagraph()
    {
        var document = new Document();
        var section = document.AddSection();
        var paragraph = ParagraphOf("boxed");
        paragraph.Borders = new Borders();
        paragraph.Borders.SetAll(new BorderEdge(1, SKColors.Black));
        section.Blocks.Add(paragraph);

        var pages = new LayoutEngine(FontRegistry.Current).Layout(document);

        var lines = pages[0].Body.OfType<LineOp>().ToList();
        Assert.Equal(4, lines.Count);
        Assert.Equal(2, lines.Count(l => Math.Abs(l.Y1 - l.Y2) < 0.01));
        Assert.Equal(2, lines.Count(l => Math.Abs(l.X1 - l.X2) < 0.01));
    }

    static Document TwoParagraphDocument(double pageHeight, int firstLines, int secondLines, Action<Paragraph>? configureSecond = null)
    {
        var document = new Document();
        var section = document.AddSection();
        section.PageHeight = pageHeight;
        section.TopMargin = section.BottomMargin = 20;

        var first = new Paragraph { DefaultStyle = Style };
        for (var i = 0; i < firstLines; i++)
        {
            if (i > 0) first.Add(new LineBreakInline());
            first.Add(new TextRun($"first {i}", Style));
        }
        section.Blocks.Add(first);

        var second = new Paragraph { DefaultStyle = Style };
        for (var i = 0; i < secondLines; i++)
        {
            if (i > 0) second.Add(new LineBreakInline());
            second.Add(new TextRun($"second {i}", Style));
        }
        configureSecond?.Invoke(second);
        section.Blocks.Add(second);
        return document;
    }

    static string PageText(Page page) => string.Join(" ", page.Body.OfType<TextOp>().Select(op => op.Text));

    [Fact]
    public void WidowControl_MovesAParagraphWhoseFirstLineWouldBeAlone()
    {
        var lineHeight = TextLayout.Layout(ParagraphOf("x"), 400, Measurer(), _ => string.Empty).Lines[0].Height;
        var pageHeight = 40 + lineHeight * 10.5;

        var pages = new LayoutEngine(FontRegistry.Current).Layout(TwoParagraphDocument(pageHeight, 9, 5));

        Assert.DoesNotContain("second", PageText(pages[0]));
        Assert.Contains("second 0", PageText(pages[1]));
    }

    [Fact]
    public void WidowControl_Off_LetsASingleLineStayOnThePage()
    {
        var lineHeight = TextLayout.Layout(ParagraphOf("x"), 400, Measurer(), _ => string.Empty).Lines[0].Height;
        var pageHeight = 40 + lineHeight * 10.5;

        var pages = new LayoutEngine(FontRegistry.Current).Layout(
            TwoParagraphDocument(pageHeight, 9, 5, p => p.WidowControl = false));

        Assert.Contains("second 0", PageText(pages[0]));
    }

    [Fact]
    public void KeepTogether_MovesTheWholeParagraph()
    {
        var lineHeight = TextLayout.Layout(ParagraphOf("x"), 400, Measurer(), _ => string.Empty).Lines[0].Height;
        var pageHeight = 40 + lineHeight * 10.5;

        var pages = new LayoutEngine(FontRegistry.Current).Layout(
            TwoParagraphDocument(pageHeight, 7, 5, p => p.KeepTogether = true));

        Assert.DoesNotContain("second", PageText(pages[0]));
        Assert.Contains("second 4", PageText(pages[1]));
    }

    [Fact]
    public void KeepWithNext_MovesTheHeadingToThePageOfItsParagraph()
    {
        var lineHeight = TextLayout.Layout(ParagraphOf("x"), 400, Measurer(), _ => string.Empty).Lines[0].Height;
        var document = new Document();
        var section = document.AddSection();
        section.PageHeight = 40 + lineHeight * 10.5;
        section.TopMargin = section.BottomMargin = 20;

        var filler = new Paragraph { DefaultStyle = Style };
        for (var i = 0; i < 9; i++)
        {
            if (i > 0) filler.Add(new LineBreakInline());
            filler.Add(new TextRun($"filler {i}", Style));
        }
        section.Blocks.Add(filler);
        var heading = ParagraphOf("Heading");
        heading.KeepWithNext = true;
        section.Blocks.Add(heading);
        section.Blocks.Add(ParagraphOf("Body text"));

        var pages = new LayoutEngine(FontRegistry.Current).Layout(document);

        Assert.DoesNotContain("Heading", PageText(pages[0]));
        Assert.StartsWith("Heading Body", PageText(pages[1]));
    }

    [Fact]
    public void Sections_UseTheirOwnPageSize()
    {
        var document = new Document();
        var portrait = document.AddSection();
        portrait.Blocks.Add(ParagraphOf("one"));
        var landscape = document.AddSection();
        landscape.PageWidth = 842;
        landscape.PageHeight = 595;
        landscape.Blocks.Add(ParagraphOf("two"));

        var pages = new LayoutEngine(FontRegistry.Current).Layout(document);

        Assert.Equal(2, pages.Count);
        Assert.True(pages[0].Height > pages[0].Width);
        Assert.True(pages[1].Width > pages[1].Height);
    }

    [Theory]
    [InlineData("dashed", "Dash")]
    [InlineData("mediumDashed", "Dash")]
    [InlineData("dotted", "Dotted")]
    [InlineData("hair", "Dotted")]
    [InlineData("dotDash", "DotDash")]
    [InlineData("mediumDashDotDot", "DotDotDash")]
    [InlineData("single", "Single")]
    [InlineData(null, "Single")]
    public void BorderPatterns_MapWordAndExcelStyleNames(string? style, string expected)
    {
        Assert.Equal(Enum.Parse<UnderlineStyle>(expected), BorderPatterns.From(style));
    }
}

file static class TextOpExtensions
{
    public static double Width(this TextOp op)
    {
        using var font = new SKFont(op.Face.Typeface, (float)op.Size);
        return font.MeasureText(op.Text);
    }
}
