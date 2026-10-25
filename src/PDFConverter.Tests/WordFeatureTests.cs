using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PDFConverter.Layout;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter.Tests;

public class WordFeatureTests
{
    static byte[] BuildDocument(params OpenXmlElement[] bodyChildren)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body(bodyChildren));
        }
        return stream.ToArray();
    }

    static W.Paragraph ParagraphOf(string text, W.RunProperties? runProperties = null, W.ParagraphProperties? paragraphProperties = null)
    {
        var run = new W.Run(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve });
        if (runProperties != null) run.PrependChild(runProperties);
        var paragraph = new W.Paragraph(run);
        if (paragraphProperties != null) paragraph.PrependChild(paragraphProperties);
        return paragraph;
    }

    [Fact]
    public void ResolveRunFormatting_ReadsStrikeHighlightSpacingAndScale()
    {
        var run = new W.Run(new W.RunProperties(
            new W.Strike(),
            new W.Highlight { Val = W.HighlightColorValues.Yellow },
            new W.Spacing { Val = 40 },
            new W.CharacterScale { Val = 50 }));

        var format = WordHelpers.ResolveRunFormatting(null, run, null);

        Assert.Equal(StrikeStyle.Single, format.Strike);
        Assert.Equal("FFFF00", format.Highlight);
        Assert.Equal(2, format.LetterSpacing);
        Assert.Equal(0.5, format.WidthScale);
        Assert.False(format.Hidden);
    }

    [Fact]
    public void ResolveRunFormatting_ReadsDoubleStrikeAndHidden()
    {
        var run = new W.Run(new W.RunProperties(new W.DoubleStrike(), new W.Vanish()));

        var format = WordHelpers.ResolveRunFormatting(null, run, null);

        Assert.Equal(StrikeStyle.Double, format.Strike);
        Assert.True(format.Hidden);
    }

    [Fact]
    public void GetParagraphFormatting_ReadsBordersAndKeepFlags()
    {
        var pPr = new W.ParagraphProperties(
            new W.KeepNext(),
            new W.KeepLines(),
            new W.WidowControl { Val = false },
            new W.ParagraphBorders(new W.BottomBorder { Val = W.BorderValues.Dashed, Size = 8 }));

        var format = WordHelpers.GetParagraphFormatting(pPr);

        Assert.True(format.KeepNext);
        Assert.True(format.KeepLines);
        Assert.False(format.WidowControl);
        Assert.NotNull(format.Borders);
        Assert.Equal(1, format.Borders.BottomWidth);
        Assert.Equal("dashed", format.Borders.BottomStyle);
    }

    [Fact]
    public void HiddenRuns_AreNotRendered()
    {
        var docx = BuildDocument(
            ParagraphOf("Visible"),
            ParagraphOf("Invisible", new W.RunProperties(new W.Vanish())));

        var words = PdfProbe.Words(Converters.DocxToPdfBytes(docx));

        Assert.Equal(["Visible"], words);
    }

    [Fact]
    public void SectionBreaks_ProduceSectionsWithTheirOwnPageSize()
    {
        var landscape = new W.SectionProperties(
            new W.PageSize { Width = 16838, Height = 11906, Orient = W.PageOrientationValues.Landscape });
        var portrait = new W.SectionProperties(new W.PageSize { Width = 11906, Height = 16838 });

        var docx = BuildDocument(
            ParagraphOf("wide", paragraphProperties: new W.ParagraphProperties(landscape)),
            ParagraphOf("tall"),
            portrait);

        var pdf = Converters.DocxToPdfBytes(docx);

        Assert.Equal(2, PdfProbe.PageCount(pdf));
        var (width1, height1) = PdfProbe.PageSize(pdf, 1);
        var (width2, height2) = PdfProbe.PageSize(pdf, 2);
        Assert.True(width1 > height1, "first section should be landscape");
        Assert.True(height2 > width2, "second section should be portrait");
    }

    [Fact]
    public void SectionBreaks_InheritHeadersFromThePreviousSection()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var header = main.AddNewPart<HeaderPart>("rIdHeader");
            header.Header = new W.Header(ParagraphOf("Inherited header"));

            var first = new W.SectionProperties(new W.HeaderReference { Id = "rIdHeader", Type = W.HeaderFooterValues.Default });
            var last = new W.SectionProperties();
            main.Document = new W.Document(new W.Body(
                ParagraphOf("one", paragraphProperties: new W.ParagraphProperties(first)),
                ParagraphOf("two"),
                last));
        }

        var pdf = Converters.DocxToPdfBytes(stream.ToArray());

        var texts = PdfProbe.Texts(pdf);
        Assert.Equal(2, PdfProbe.PageCount(pdf));
        Assert.Equal(2, texts.Count(t => t.Content.Contains("Inherited header")));
    }

    [Fact]
    public void Metadata_IsCopiedFromTheDocumentProperties()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body(ParagraphOf("text")));
            document.PackageProperties.Title = "The title";
            document.PackageProperties.Creator = "The author";
        }

        var pdf = Converters.DocxToPdfBytes(stream.ToArray());

        using var reader = UglyToad.PdfPig.PdfDocument.Open(pdf);
        Assert.Equal("The title", reader.Information.Title);
        Assert.Equal("The author", reader.Information.Author);
    }
}
