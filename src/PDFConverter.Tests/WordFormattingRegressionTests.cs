using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;
using PdfSharp.Pdf.IO;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace PDFConverter.Tests;

public class WordFormattingRegressionTests
{
    static W.Style BoldStyle(string id) => new(
        new W.StyleName { Val = id },
        new W.StyleRunProperties(new W.Bold()))
    { StyleId = id, Type = W.StyleValues.Paragraph };

    static (WordprocessingDocument Document, MemoryStream Stream) NewDocument(params W.Style[] styles)
    {
        var stream = new MemoryStream();
        var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        if (styles.Length > 0)
        {
            var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var definitions = new W.Styles();
            foreach (var style in styles) definitions.AppendChild(style);
            stylePart.Styles = definitions;
        }
        return (document, stream);
    }

    [Fact]
    public void ResolveRunFormatting_ExplicitBoldOff_OverridesBoldStyle()
    {
        var (document, stream) = NewDocument(BoldStyle("Strong"));
        using (stream)
        using (document)
        {
            var paragraph = new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Strong" }));
            var run = new W.Run(new W.RunProperties(new W.Bold { Val = false }), new W.Text("plain"));
            paragraph.AppendChild(run);

            var format = WordHelpers.ResolveRunFormatting(document.MainDocumentPart, run, paragraph);

            Assert.False(format.Bold);
        }
    }

    [Fact]
    public void ResolveRunFormatting_StyleBold_AppliesWhenRunIsSilent()
    {
        var (document, stream) = NewDocument(BoldStyle("Strong"));
        using (stream)
        using (document)
        {
            var paragraph = new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Strong" }));
            var run = new W.Run(new W.Text("bold"));
            paragraph.AppendChild(run);

            Assert.True(WordHelpers.ResolveRunFormatting(document.MainDocumentPart, run, paragraph).Bold);
        }
    }

    [Fact]
    public void ResolveRunFormatting_ReadsSuperscriptAndCaps()
    {
        var (document, stream) = NewDocument();
        using (stream)
        using (document)
        {
            var run = new W.Run(
                new W.RunProperties(
                    new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Superscript },
                    new W.Caps()),
                new W.Text("ref"));
            var paragraph = new W.Paragraph(run);

            var format = WordHelpers.ResolveRunFormatting(document.MainDocumentPart, run, paragraph);

            Assert.Equal(RunVerticalAlignment.Superscript, format.VerticalAlignment);
            Assert.True(format.AllCaps);
            Assert.Equal("REF", format.TransformText("ref"));
        }
    }

    [Fact]
    public void GetTableGridColumnWidths_PercentageCellWidths_ResolveAgainstAvailableWidth()
    {
        var table = new W.Table(
            new W.TableRow(
                new W.TableCell(
                    new W.TableCellProperties(
                        new W.TableCellWidth { Width = "2500", Type = W.TableWidthUnitValues.Pct }),
                    new W.Paragraph()),
                new W.TableCell(
                    new W.TableCellProperties(
                        new W.TableCellWidth { Width = "2500", Type = W.TableWidthUnitValues.Pct }),
                    new W.Paragraph())));

        var widths = WordHelpers.GetTableGridColumnWidths(table, 400);

        Assert.Equal(200.0, widths[0], 1);
        Assert.Equal(200.0, widths[1], 1);
    }

    [Fact]
    public void GetTableWidth_PercentageTable_ResolvesAgainstAvailableWidth()
    {
        var tblPr = new W.TableProperties(
            new W.TableWidth { Width = "5000", Type = W.TableWidthUnitValues.Pct });

        Assert.Equal(400.0, WordHelpers.GetTableWidth(tblPr, 400)!.Value, 1);
    }

    [Fact]
    public void GetTableWidth_AutoTable_ReturnsNull()
    {
        var tblPr = new W.TableProperties(
            new W.TableWidth { Width = "0", Type = W.TableWidthUnitValues.Auto });

        Assert.Null(WordHelpers.GetTableWidth(tblPr, 400));
    }

    [Fact]
    public void GetParagraphFormatting_ReadsPageBreakBeforeAndShading()
    {
        var pPr = new W.ParagraphProperties(
            new W.PageBreakBefore(),
            new W.Shading { Fill = "FFFF00" });

        var format = WordHelpers.GetParagraphFormatting(pPr);

        Assert.True(format.PageBreakBefore);
        Assert.Equal("FFFF00", format.ShadingColor);
    }

    [Fact]
    public void DocxToPdfBytes_PageField_RendersThePageNumber()
    {
        var texts = ExtractTexts(Converters.DocxToPdfBytes(BuildDocumentWithFooterPageField()));

        Assert.Contains("1", texts);
    }

    [Fact]
    public void DocxToPdfBytes_ManualPageBreak_StartsASecondPage()
    {
        using var ms = new MemoryStream(Converters.DocxToPdfBytes(BuildDocumentWithPageBreak()));
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);

        Assert.Equal(2, pdf.PageCount);
    }

    [Fact]
    public void GetTableCellMargins_NoOverride_UsesWordDefault()
    {
        var margins = WordHelpers.GetTableCellMargins(null, new W.TableProperties());

        Assert.Equal(5.4, margins.Left, 2);
        Assert.Equal(5.4, margins.Right, 2);
        Assert.Equal(0, margins.Top, 2);
    }

    [Fact]
    public void GetTableCellMargins_ZeroOverride_IsHonoured()
    {
        var tblPr = new W.TableProperties(new W.TableCellMarginDefault(
            new W.TableCellLeftMargin { Width = 0, Type = W.TableWidthValues.Dxa },
            new W.TableCellRightMargin { Width = 0, Type = W.TableWidthValues.Dxa }));

        var margins = WordHelpers.GetTableCellMargins(null, tblPr);

        Assert.Equal(0, margins.Left, 2);
        Assert.Equal(0, margins.Right, 2);
    }

    [Fact]
    public void GetTableCellMargins_ExplicitWidths_AreConvertedFromTwips()
    {
        var tblPr = new W.TableProperties(new W.TableCellMarginDefault(
            new W.TableCellLeftMargin { Width = 240, Type = W.TableWidthValues.Dxa },
            new W.TopMargin { Width = "120", Type = W.TableWidthUnitValues.Dxa }));

        var margins = WordHelpers.GetTableCellMargins(null, tblPr);

        Assert.Equal(12.0, margins.Left, 2);
        Assert.Equal(6.0, margins.Top, 2);
        Assert.Equal(5.4, margins.Right, 2);
    }

    [Fact]
    public void DocxToPdfBytes_ParagraphsWithNoSpacing_DoNotGainAnEightPointGap()
    {
        using var ms = new MemoryStream(Converters.DocxToPdfBytes(BuildDocumentWithEmptyParagraphs(30)));
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);

        Assert.Equal(1, pdf.PageCount);
    }

    [Fact]
    public void DocxToPdfBytes_ImageExtent_OverridesTheFilesAspectRatio()
    {
        var placements = ExtractImagePlacements(Converters.DocxToPdfBytes(BuildDocumentWithStretchedImage()));

        var (width, height) = Assert.Single(placements);
        Assert.Equal(144.0, width, 1);
        Assert.Equal(36.0, height, 1);
    }

    static byte[] BuildDocumentWithEmptyParagraphs(int count)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            var body = new W.Body();
            for (var i = 0; i < count; i++) body.AppendChild(new W.Paragraph());
            body.AppendChild(new W.SectionProperties(new W.PageSize { Width = 12240, Height = 15840 }));
            mainPart.Document = new W.Document(body);
        }
        return ms.ToArray();
    }

    static byte[] BuildDocumentWithStretchedImage()
    {
        const long widthEmu = 1828800;
        const long heightEmu = 457200;

        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var png = new MemoryStream(SquarePng())) imagePart.FeedData(png);

            var drawing = new W.Drawing(new DW.Inline(
                new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                new DW.DocProperties { Id = 1, Name = "square" },
                new A.Graphic(new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0, Name = "square" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = mainPart.GetIdOfPart(imagePart) },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));

            mainPart.Document = new W.Document(new W.Body(
                new W.Paragraph(new W.Run(drawing)),
                new W.SectionProperties(new W.PageSize { Width = 12240, Height = 15840 })));
        }
        return ms.ToArray();
    }

    static byte[] SquarePng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    static List<(double Width, double Height)> ExtractImagePlacements(byte[] pdfBytes)
    {
        var placements = new List<(double, double)>();
        using var ms = new MemoryStream(pdfBytes);
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        foreach (var page in pdf.Pages) WalkImages(ContentReader.ReadContent(page), placements, new double[2]);
        return placements;
    }

    static void WalkImages(CSequence sequence, List<(double, double)> placements, double[] lastTransform)
    {
        foreach (var item in sequence)
        {
            if (item is CSequence nested) { WalkImages(nested, placements, lastTransform); continue; }
            if (item is not COperator op) continue;

            if (op.OpCode.Name == "cm" && op.Operands.Count >= 6)
            {
                lastTransform[0] = Operand(op.Operands[0]);
                lastTransform[1] = Operand(op.Operands[3]);
            }
            else if (op.OpCode.Name == "Do")
            {
                placements.Add((lastTransform[0], Math.Abs(lastTransform[1])));
            }
        }
    }

    static double Operand(CObject value) => value switch
    {
        CReal real => real.Value,
        CInteger integer => integer.Value,
        _ => 0,
    };

    [Fact]
    public void DocDefaults_AsWrittenByWord_AreRead()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylePart.Styles = new W.Styles(new W.DocDefaults(
            new W.RunPropertiesDefault(
                new W.RunPropertiesBaseStyle(new W.FontSize { Val = "22" })),
            new W.ParagraphPropertiesDefault(
                new W.ParagraphPropertiesBaseStyle(
                    new W.SpacingBetweenLines { After = "160" }))));

        Assert.Equal("22",
            WordHelpers.GetDocDefaultsRunProperties(mainPart)!.FontSize!.Val!.Value);
        Assert.Equal("160",
            WordHelpers.GetDocDefaultsParagraphProperties(mainPart)!.SpacingBetweenLines!.After!.Value);
    }

    static byte[] BuildDocumentWithFooterPageField()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new W.Footer(new W.Paragraph(
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
                new W.Run(new W.FieldCode(" PAGE ")),
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
                new W.Run(new W.Text("stale")),
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End })));

            mainPart.Document = new W.Document(new W.Body(
                new W.Paragraph(new W.Run(new W.Text("body"))),
                new W.SectionProperties(
                    new W.FooterReference
                    {
                        Type = W.HeaderFooterValues.Default,
                        Id = mainPart.GetIdOfPart(footerPart),
                    },
                    new W.PageSize { Width = 12240, Height = 15840 })));
        }
        return ms.ToArray();
    }

    static byte[] BuildDocumentWithPageBreak()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new W.Document(new W.Body(
                new W.Paragraph(new W.Run(new W.Text("first"))),
                new W.Paragraph(new W.Run(new W.Break { Type = W.BreakValues.Page })),
                new W.Paragraph(new W.Run(new W.Text("second"))),
                new W.SectionProperties(new W.PageSize { Width = 12240, Height = 15840 })));
        }
        return ms.ToArray();
    }

    static List<string> ExtractTexts(byte[] pdfBytes)
    {
        var texts = new List<string>();
        using var ms = new MemoryStream(pdfBytes);
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        foreach (var page in pdf.Pages) Walk(ContentReader.ReadContent(page), texts);
        return texts;
    }

    static void Walk(CSequence sequence, List<string> texts)
    {
        foreach (var item in sequence)
        {
            if (item is CSequence nested) { Walk(nested, texts); continue; }
            if (item is not COperator op || op.OpCode.Name is not ("Tj" or "TJ")) continue;
            if (op.Operands.Count == 0) continue;

            var text = string.Empty;
            if (op.Operands[0] is CString s) text = s.Value;
            else if (op.Operands[0] is CArray array)
                foreach (var element in array)
                    if (element is CString cs) text += cs.Value;

            if (!string.IsNullOrWhiteSpace(text)) texts.Add(text.Trim());
        }
    }
}
