using PDFConverter.Layout;
using SkiaSharp;
using Xunit;

namespace PDFConverter.Tests;

public class LayoutEngineTests
{
    static readonly TextStyle Style = new("Arial", 10);

    static LayoutEngineTests() => OpenXmlHelpers.EnsureFontResolverInitialized();

    static TextMeasurer Measurer() => new(FontRegistry.Current);

    static Paragraph ParagraphOf(string text, ParagraphAlignment alignment = ParagraphAlignment.Left)
    {
        var paragraph = new Paragraph { Alignment = alignment, DefaultStyle = Style };
        paragraph.Add(new TextRun(text, Style));
        return paragraph;
    }

    static IEnumerable<DrawOp> Flatten(IEnumerable<DrawOp> ops) =>
        ops.SelectMany(op => op is ClipOp clip ? Flatten(clip.Children) : [op]);

    static string TextOf(IEnumerable<DrawOp> ops) =>
        string.Join(" ", Flatten(ops).OfType<TextOp>().Select(op => op.Text));

    [Fact]
    public void Layout_LongText_WrapsWithinTheAvailableWidth()
    {
        var text = string.Join(" ", Enumerable.Repeat("lorem ipsum dolor", 20));
        var layout = TextLayout.Layout(ParagraphOf(text), 200, Measurer(), _ => string.Empty);

        Assert.True(layout.Lines.Count > 3);
        Assert.All(layout.Lines, line => Assert.True(line.Width <= 200.01, $"line width {line.Width}"));
    }

    [Fact]
    public void Layout_WordWiderThanTheLine_IsBrokenAcrossLines()
    {
        var word = new string('x', 80);
        var layout = TextLayout.Layout(ParagraphOf(word), 100, Measurer(), _ => string.Empty);

        Assert.True(layout.Lines.Count > 1);
        Assert.Equal(word, string.Concat(layout.Lines.SelectMany(l => l.Items).Select(i => i.Text)));
        Assert.All(layout.Lines, line => Assert.True(line.Width <= 100.01));
    }

    [Fact]
    public void Layout_Justified_StretchesAllButTheLastLine()
    {
        var text = string.Join(" ", Enumerable.Repeat("word", 60));
        var layout = TextLayout.Layout(ParagraphOf(text, ParagraphAlignment.Justify), 150, Measurer(), _ => string.Empty);

        Assert.True(layout.Lines.Count > 2);
        foreach (var line in layout.Lines.Take(layout.Lines.Count - 1))
            Assert.Equal(150, line.Width, 0.01);
        Assert.True(layout.Lines[^1].Width < 150);
    }

    [Fact]
    public void Layout_TabStop_MovesTheFollowingTextToTheStop()
    {
        var paragraph = new Paragraph { DefaultStyle = Style, TabStops = [new TabStop(100)] };
        paragraph.Add(new TextRun("a", Style)).Add(new TabInline()).Add(new TextRun("b", Style));

        var line = Assert.Single(TextLayout.Layout(paragraph, 300, Measurer(), _ => string.Empty).Lines);
        var b = line.Items.Last(i => i.Kind == ItemKind.Text);
        Assert.Equal(100, b.X, 0.01);
    }

    [Fact]
    public void Layout_LeadingSpacesAtParagraphStart_AreKept()
    {
        var layout = TextLayout.Layout(ParagraphOf("   indented"), 300, Measurer(), _ => string.Empty);

        var first = layout.Lines[0].Items[0];
        Assert.Equal(ItemKind.Space, first.Kind);
        Assert.True(layout.Lines[0].Items[1].X > 0);
    }

    [Fact]
    public void Layout_EmptyParagraph_HasOneLineWithFontHeight()
    {
        var layout = TextLayout.Layout(new Paragraph { DefaultStyle = Style }, 300, Measurer(), _ => string.Empty);

        var line = Assert.Single(layout.Lines);
        Assert.InRange(line.Height, 10, 14);
    }

    [Fact]
    public void Engine_ManyParagraphs_PaginateAndNumberPagesInTheFooter()
    {
        var document = new Document();
        var section = document.AddSection();
        section.PageHeight = 300;
        for (var i = 0; i < 40; i++) section.Blocks.Add(ParagraphOf($"Paragraph {i}"));

        var footer = new Paragraph { DefaultStyle = Style };
        footer.Add(new FieldRun(FieldKind.Page, Style)).Add(new TextRun("/", Style)).Add(new FieldRun(FieldKind.NumPages, Style));
        section.PrimaryFooter.Blocks.Add(footer);

        var pages = new LayoutEngine(FontRegistry.Current).Layout(document);

        Assert.True(pages.Count > 2);
        Assert.Equal($"2 / {pages.Count}", TextOf(pages[1].Overlay));
        Assert.StartsWith("Paragraph 0 ", TextOf(pages[0].Body));
        Assert.DoesNotContain("Paragraph 0 ", TextOf(pages[1].Body));
    }

    [Fact]
    public void Engine_SpaceBefore_IsDroppedAtANaturalPageTopButKeptAtTheStart()
    {
        var document = new Document();
        var section = document.AddSection();
        section.PageHeight = 200;
        section.TopMargin = section.BottomMargin = 20;
        for (var i = 0; i < 30; i++)
        {
            var paragraph = ParagraphOf($"p{i}");
            paragraph.SpaceBefore = 30;
            section.Blocks.Add(paragraph);
        }

        var pages = new LayoutEngine(FontRegistry.Current).Layout(document);

        var firstOnPage1 = Flatten(pages[0].Body).OfType<TextOp>().First();
        var firstOnPage2 = Flatten(pages[1].Body).OfType<TextOp>().First();
        Assert.True(firstOnPage1.Baseline > 20 + 30, $"baseline {firstOnPage1.Baseline}");
        Assert.True(firstOnPage2.Baseline < 20 + 30, $"baseline {firstOnPage2.Baseline}");
    }

    [Fact]
    public void Engine_HeaderRow_IsRepeatedOnEveryPage()
    {
        var document = new Document();
        var section = document.AddSection();
        section.PageHeight = 250;

        var table = section.AddTable();
        table.ColumnWidths.Add(200);
        var header = table.AddRow();
        header.IsHeader = true;
        header.Cell(0).Blocks.Add(ParagraphOf("HEADER"));
        for (var i = 0; i < 30; i++) table.AddRow().Cell(0).Blocks.Add(ParagraphOf($"row {i}"));

        var pages = new LayoutEngine(FontRegistry.Current).Layout(document);

        Assert.True(pages.Count > 1);
        foreach (var page in pages)
            Assert.Equal("HEADER", Flatten(page.Body).OfType<TextOp>().First().Text);
    }

    [Fact]
    public void Engine_ExactRowHeight_ClipsTheCellContent()
    {
        var document = new Document();
        var section = document.AddSection();
        var table = section.AddTable();
        table.ColumnWidths.Add(100);
        var row = table.AddRow();
        row.MinHeight = 8;
        row.ExactHeight = true;
        row.Cell(0).Blocks.Add(ParagraphOf("tall content that does not fit in eight points"));

        var pages = new LayoutEngine(FontRegistry.Current).Layout(document);

        var clip = Assert.Single(pages[0].Body.OfType<ClipOp>());
        Assert.Equal(8, clip.Height, 0.01);
    }

    [Fact]
    public void Engine_RowSpan_GrowsTheLastSpannedRowToFitTheContent()
    {
        var document = new Document();
        var section = document.AddSection();
        var table = section.AddTable();
        table.ColumnWidths.AddRange([100, 100]);
        table.DefaultBorders.SetAll(new BorderEdge(0.5, SKColors.Black));
        var first = table.AddRow();
        first.Cell(0).RowSpan = 2;
        first.Cell(0).Blocks.Add(ParagraphOf(string.Join(" ", Enumerable.Repeat("spanning text", 12))));
        first.Cell(1).Blocks.Add(ParagraphOf("a"));
        var second = table.AddRow();
        second.Cell(1).Blocks.Add(ParagraphOf("b"));

        var pages = new LayoutEngine(FontRegistry.Current).Layout(document);

        var spanningText = Flatten(pages[0].Body).OfType<TextOp>().Where(op => op.X < section.LeftMargin + 100).ToList();
        var lastBorder = pages[0].Body.OfType<LineOp>().Max(line => Math.Max(line.Y1, line.Y2));
        Assert.True(spanningText.Max(op => op.Baseline) <= lastBorder, "spanning content overflowed the table");
    }

    [Fact]
    public void Engine_BookmarkLink_WritesADestinationAndAnInternalLink()
    {
        var document = new Document();
        var section = document.AddSection();
        var target = new Paragraph { DefaultStyle = Style };
        target.Add(new BookmarkInline("top")).Add(new TextRun("Top", Style));
        section.Blocks.Add(target);
        var link = new Paragraph { DefaultStyle = Style };
        link.Add(new TextRun("go to top", Style, "#top"));
        section.Blocks.Add(link);

        var pages = new LayoutEngine(FontRegistry.Current).Layout(document);
        Assert.Contains(pages[0].Body, op => op is DestinationOp { Name: "top" });
        Assert.Contains(pages[0].Body, op => op is LinkOp { Url: "#top" });

        using var output = new MemoryStream();
        SkiaPdfWriter.Write(pages, output, null);
        Assert.True(output.Length > 0);
    }

    [Fact]
    public void Engine_NoWrapParagraph_OverflowsInsteadOfBreaking()
    {
        var paragraph = ParagraphOf(string.Join(" ", Enumerable.Repeat("overflowing", 10)));
        paragraph.NoWrap = true;

        var layout = TextLayout.Layout(paragraph, 50, Measurer(), _ => string.Empty);

        var line = Assert.Single(layout.Lines);
        Assert.True(line.Width > 50);
    }

    [Fact]
    public void Engine_FontFallback_SplitsRunsByFace()
    {
        var paragraph = ParagraphOf("plain 😀 text");
        var line = Assert.Single(TextLayout.Layout(paragraph, 300, Measurer(), _ => string.Empty).Lines);

        var faces = line.Items.Where(i => i.Kind == ItemKind.Text).Select(i => i.Face!.Family).Distinct().ToList();
        Assert.True(faces.Count >= 2, string.Join(", ", faces));
    }

    [Fact]
    public void Writer_SubsetsFontsWhenEmbedding()
    {
        var document = new Document();
        var section = document.AddSection();
        section.Blocks.Add(ParagraphOf("Hello"));
        var pages = new LayoutEngine(FontRegistry.Current).Layout(document);

        using var output = new MemoryStream();
        SkiaPdfWriter.Write(pages, output, null);

        Assert.InRange(output.Length, 1, 120_000);
        Assert.Equal(["Hello"], PdfProbe.Words(output.ToArray()));
    }
}
