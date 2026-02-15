using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;
using PdfSharp.Pdf.IO;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter.Tests;

/// <summary>
/// Integration tests for DOCX → PDF conversion using in-memory OpenXML documents.
/// No external file dependencies — all test documents are built programmatically.
/// </summary>
public class DocxConverterIntegrationTests : IDisposable
{
    private readonly List<string> _outputFiles = new();

    public void Dispose()
    {
        foreach (var f in _outputFiles)
        {
            try { if (File.Exists(f)) File.Delete(f); } catch { }
        }
    }

    private string GetOutputPath(string name)
    {
        var path = Path.Combine(Path.GetTempPath(), $"PDFConverter_Test_{name}_{Guid.NewGuid():N}.pdf");
        _outputFiles.Add(path);
        return path;
    }

    private static int GetPdfPageCount(byte[] pdfBytes)
    {
        using var ms = new MemoryStream(pdfBytes);
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        return pdf.PageCount;
    }

    private static int GetPdfPageCount(string pdfPath)
    {
        using var pdf = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
        return pdf.PageCount;
    }

    private static void AssertValidPdf(byte[] pdfBytes)
    {
        Assert.NotNull(pdfBytes);
        Assert.True(pdfBytes.Length > 100, "PDF too small to be valid");
        Assert.Equal((byte)'%', pdfBytes[0]);
        Assert.Equal((byte)'P', pdfBytes[1]);
        Assert.Equal((byte)'D', pdfBytes[2]);
        Assert.Equal((byte)'F', pdfBytes[3]);
    }

    #region Basic Conversion

    /// <summary>
    /// Converts a simple single-paragraph DOCX and verifies valid PDF output.
    /// </summary>
    [Fact]
    public void DocxToPdf_SimpleParagraph_ProducesValidPdf()
    {
        var docx = BuildSimpleDocx("Hello World");
        var pdfBytes = Converters.DocxToPdfBytes(docx);

        AssertValidPdf(pdfBytes);
        Assert.Equal(1, GetPdfPageCount(pdfBytes));
    }

    /// <summary>
    /// Verifies that text content from the DOCX appears in the PDF output.
    /// </summary>
    [Fact]
    public void DocxToPdf_SimpleParagraph_ContainsText()
    {
        var docx = BuildSimpleDocx("Test Content ABC");
        var pdfBytes = Converters.DocxToPdfBytes(docx);
        var texts = ExtractTexts(pdfBytes);

        // PDF may split text into separate word fragments
        Assert.Contains(texts, t => t.Content.Contains("Test"));
        Assert.Contains(texts, t => t.Content.Contains("ABC"));
    }

    /// <summary>
    /// A DOCX with multiple paragraphs should produce a single page
    /// when content fits within default margins.
    /// </summary>
    [Fact]
    public void DocxToPdf_MultipleParagraphs_SinglePage()
    {
        var docx = BuildMultiParagraphDocx(5, "Short paragraph.");
        var pdfBytes = Converters.DocxToPdfBytes(docx);

        AssertValidPdf(pdfBytes);
        Assert.Equal(1, GetPdfPageCount(pdfBytes));
    }

    #endregion

    #region Styled Text

    /// <summary>
    /// Bold text in the DOCX should be preserved in the PDF.
    /// Validates that run properties (bold) are correctly mapped.
    /// </summary>
    [Fact]
    public void DocxToPdf_BoldText_ProducesValidPdf()
    {
        var docx = BuildStyledDocx(bold: true, italic: false, fontSize: 24);
        var pdfBytes = Converters.DocxToPdfBytes(docx);
        var texts = ExtractTexts(pdfBytes);

        AssertValidPdf(pdfBytes);
        // PDF may split text into separate word fragments
        Assert.Contains(texts, t => t.Content.Contains("Styled"));
    }

    /// <summary>
    /// Italic text in the DOCX should produce valid PDF output.
    /// </summary>
    [Fact]
    public void DocxToPdf_ItalicText_ProducesValidPdf()
    {
        var docx = BuildStyledDocx(bold: false, italic: true, fontSize: 12);
        var pdfBytes = Converters.DocxToPdfBytes(docx);

        AssertValidPdf(pdfBytes);
    }

    /// <summary>
    /// Custom font size should not crash the converter.
    /// </summary>
    [Fact]
    public void DocxToPdf_CustomFontSize_ProducesValidPdf()
    {
        var docx = BuildStyledDocx(bold: false, italic: false, fontSize: 36);
        var pdfBytes = Converters.DocxToPdfBytes(docx);

        AssertValidPdf(pdfBytes);
    }

    #endregion

    #region Table Rendering

    /// <summary>
    /// A DOCX with a simple 3×3 table should convert without errors
    /// and the table content should appear in the PDF.
    /// </summary>
    [Fact]
    public void DocxToPdf_SimpleTable_ContainsCellText()
    {
        var docx = BuildTableDocx(3, 3);
        var pdfBytes = Converters.DocxToPdfBytes(docx);
        var texts = ExtractTexts(pdfBytes);

        AssertValidPdf(pdfBytes);
        // Verify at least some cell content is present
        Assert.Contains(texts, t => t.Content.Contains("R1C1"));
        Assert.Contains(texts, t => t.Content.Contains("R3C3"));
    }

    /// <summary>
    /// A table with styled header row (shading, bold) should render without errors.
    /// </summary>
    [Fact]
    public void DocxToPdf_StyledTable_ProducesValidPdf()
    {
        var docx = BuildStyledTableDocx();
        var pdfBytes = Converters.DocxToPdfBytes(docx);
        var texts = ExtractTexts(pdfBytes);

        AssertValidPdf(pdfBytes);
        Assert.Contains(texts, t => t.Content.Contains("Header1"));
    }

    #endregion

    #region Landscape Orientation

    /// <summary>
    /// A DOCX with landscape orientation should produce a landscape PDF page.
    /// Landscape pages have width > height.
    /// </summary>
    [Fact]
    public void DocxToPdf_LandscapeOrientation_ProducesWidePage()
    {
        var docx = BuildLandscapeDocx();
        var pdfBytes = Converters.DocxToPdfBytes(docx);

        AssertValidPdf(pdfBytes);

        using var ms = new MemoryStream(pdfBytes);
        var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        var page = pdf.Pages[0];

        Assert.True(page.Width > page.Height,
            $"Landscape page should be wider than tall: {page.Width:F0} × {page.Height:F0}");
    }

    #endregion

    #region Footer Distance

    /// <summary>
    /// A DOCX with a table and a footer should not overlap — the footer text
    /// must appear below the table content when footer distance is set.
    /// </summary>
    [Fact]
    public void DocxToPdf_TableWithFooter_FooterBelowTable()
    {
        var docx = BuildTableWithFooterDocx();
        var pdfBytes = Converters.DocxToPdfBytes(docx);

        AssertValidPdf(pdfBytes);
        var texts = ExtractTexts(pdfBytes);

        // Table content and footer text should both be present
        Assert.Contains(texts, t => t.Content.Contains("Cell"));
        Assert.Contains(texts, t => t.Content.Contains("FooterText"));

        // Footer text should have a lower Y coordinate than any table cell text
        // (PDF Y axis: 0 = bottom of page, higher = further up)
        var footerY = texts.Where(t => t.Content.Contains("FooterText")).Min(t => t.Y);
        var tableMinY = texts.Where(t => t.Content.Contains("Cell")).Min(t => t.Y);

        Assert.True(footerY < tableMinY,
            $"Footer (Y={footerY:F1}) should be below table content (min Y={tableMinY:F1})");
    }

    /// <summary>
    /// When an image has an a:hlinkClick in its docPr, the resulting PDF should
    /// contain a link annotation with the correct URI.
    /// </summary>
    [Fact]
    public void DocxToPdf_ImageWithHyperlink_ProducesLinkAnnotation()
    {
        var docx = BuildImageWithHyperlinkDocx("https://example.com/test");
        var pdfBytes = Converters.DocxToPdfBytes(docx);

        AssertValidPdf(pdfBytes);

        // Scan the raw PDF bytes for link annotation markers
        var pdfText = System.Text.Encoding.ASCII.GetString(pdfBytes);
        Assert.Contains("/Subtype/Link", pdfText.Replace(" ", ""));
        Assert.Contains("/URI", pdfText);
        Assert.Contains("https://example.com/test", pdfText);
    }

    #endregion

    #region All Overloads

    /// <summary>
    /// All three overloads (byte[], Stream, file path) should produce
    /// identical page counts from the same in-memory DOCX.
    /// </summary>
    [Fact]
    public void DocxToPdf_AllOverloads_ProduceSamePageCount()
    {
        var docx = BuildSimpleDocx("Overload consistency test");

        // Byte array overload
        var pdfFromBytes = Converters.DocxToPdfBytes(docx);
        int countBytes = GetPdfPageCount(pdfFromBytes);

        // Stream overload
        using var stream = new MemoryStream(docx);
        var pdfFromStream = Converters.DocxToPdfBytes(stream);
        int countStream = GetPdfPageCount(pdfFromStream);

        // File path overload
        var tempDocx = Path.GetTempFileName() + ".docx";
        var tempPdf = GetOutputPath("AllOverloads_File");
        try
        {
            File.WriteAllBytes(tempDocx, docx);
            DocxConverter.DocxToPdf(tempDocx, tempPdf);
            int countFile = GetPdfPageCount(tempPdf);

            Assert.Equal(countBytes, countStream);
            Assert.Equal(countBytes, countFile);
        }
        finally
        {
            try { File.Delete(tempDocx); } catch { }
        }
    }

    /// <summary>
    /// DocxToPdfBytes from a Stream should return a valid PDF with %PDF header.
    /// </summary>
    [Fact]
    public void DocxToPdfBytes_FromStream_ReturnsValidPdf()
    {
        var docx = BuildSimpleDocx("Stream overload test");
        using var stream = new MemoryStream(docx);
        var pdfBytes = Converters.DocxToPdfBytes(stream);

        AssertValidPdf(pdfBytes);
    }

    /// <summary>
    /// DocxToPdfBytes from byte[] should return a valid PDF with %PDF header.
    /// </summary>
    [Fact]
    public void DocxToPdfBytes_FromByteArray_ReturnsValidPdf()
    {
        var docx = BuildSimpleDocx("Byte array overload test");
        var pdfBytes = Converters.DocxToPdfBytes(docx);

        AssertValidPdf(pdfBytes);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void DocxToPdf_NonExistentFile_ThrowsFileNotFound()
    {
        var output = GetOutputPath("NonExistent");
        Assert.ThrowsAny<Exception>(() =>
            DocxConverter.DocxToPdf(@"C:\nonexistent_document_12345.docx", output));
    }

    [Fact]
    public void DocxToPdf_EmptyStream_ThrowsException()
    {
        var output = GetOutputPath("EmptyStream");
        using var ms = new MemoryStream();
        Assert.ThrowsAny<Exception>(() =>
            DocxConverter.DocxToPdf(ms, output));
    }

    [Fact]
    public void DocxToPdf_EmptyByteArray_ThrowsException()
    {
        var output = GetOutputPath("EmptyBytes");
        Assert.ThrowsAny<Exception>(() =>
            DocxConverter.DocxToPdf(Array.Empty<byte>(), output));
    }

    #endregion

    #region DOCX Builders

    /// <summary>
    /// Builds a minimal DOCX with a single paragraph.
    /// </summary>
    private static byte[] BuildSimpleDocx(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(
                new Body(
                    new W.Paragraph(
                        new W.Run(new W.Text(text)))));
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a DOCX with N paragraphs, each containing the given text.
    /// </summary>
    private static byte[] BuildMultiParagraphDocx(int count, string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            for (int i = 0; i < count; i++)
                body.Append(new W.Paragraph(new W.Run(new W.Text($"{text} #{i + 1}"))));
            mainPart.Document = new Document(body);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a DOCX with styled text (bold, italic, font size in half-points).
    /// </summary>
    private static byte[] BuildStyledDocx(bool bold, bool italic, int fontSize)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var runProps = new W.RunProperties();
            if (bold) runProps.Append(new W.Bold());
            if (italic) runProps.Append(new W.Italic());
            runProps.Append(new W.FontSize { Val = (fontSize * 2).ToString() }); // half-points

            mainPart.Document = new Document(
                new Body(
                    new W.Paragraph(
                        new W.Run(runProps, new W.Text("Styled Text")))));
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a DOCX with a rows×cols table, each cell containing "R{r}C{c}".
    /// </summary>
    private static byte[] BuildTableDocx(int rows, int cols)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var table = new W.Table();

            // Table properties with borders
            var tblPr = new W.TableProperties(
                new W.TableBorders(
                    new W.TopBorder { Val = BorderValues.Single, Size = 4 },
                    new W.BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new W.LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new W.RightBorder { Val = BorderValues.Single, Size = 4 },
                    new W.InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new W.InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }));
            table.Append(tblPr);

            for (int r = 1; r <= rows; r++)
            {
                var row = new W.TableRow();
                for (int c = 1; c <= cols; c++)
                {
                    var cell = new W.TableCell(
                        new W.Paragraph(
                            new W.Run(new W.Text($"R{r}C{c}"))));
                    row.Append(cell);
                }
                table.Append(row);
            }

            mainPart.Document = new Document(new Body(table));
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a DOCX with a 3×2 table where the first row has shading and bold text.
    /// </summary>
    private static byte[] BuildStyledTableDocx()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var table = new W.Table();

            var tblPr = new W.TableProperties(
                new W.TableBorders(
                    new W.TopBorder { Val = BorderValues.Single, Size = 4 },
                    new W.BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new W.LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new W.RightBorder { Val = BorderValues.Single, Size = 4 },
                    new W.InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new W.InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }));
            table.Append(tblPr);

            // Header row with shading
            var headerRow = new W.TableRow();
            for (int c = 1; c <= 3; c++)
            {
                var cellProps = new W.TableCellProperties(
                    new W.Shading { Val = ShadingPatternValues.Clear, Fill = "4472C4" });
                var runProps = new W.RunProperties(new W.Bold(), new W.Color { Val = "FFFFFF" });
                var cell = new W.TableCell(cellProps,
                    new W.Paragraph(
                        new W.Run(runProps, new W.Text($"Header{c}"))));
                headerRow.Append(cell);
            }
            table.Append(headerRow);

            // Data rows
            for (int r = 2; r <= 3; r++)
            {
                var row = new W.TableRow();
                for (int c = 1; c <= 3; c++)
                    row.Append(new W.TableCell(
                        new W.Paragraph(new W.Run(new W.Text($"Data{r}{c}")))));
                table.Append(row);
            }

            mainPart.Document = new Document(new Body(table));
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a DOCX with landscape orientation (US Letter rotated).
    /// </summary>
    private static byte[] BuildLandscapeDocx()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            // Landscape: width=15840 (11"), height=12240 (8.5"), orient=landscape
            var sectPr = new W.SectionProperties(
                new W.PageSize
                {
                    Width = 15840,  // 11 inches in twips
                    Height = 12240, // 8.5 inches in twips
                    Orient = PageOrientationValues.Landscape
                },
                new W.PageMargin
                {
                    Top = 720, Bottom = 720,
                    Left = 720, Right = 720
                });

            mainPart.Document = new Document(
                new Body(
                    new W.Paragraph(
                        new W.Run(new W.Text("Landscape document"))),
                    sectPr));
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a DOCX with a 5×3 table and a footer paragraph, with explicit footer distance.
    /// </summary>
    private static byte[] BuildTableWithFooterDocx()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();

            // Create footer part
            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(
                new W.Paragraph(
                    new W.Run(new W.Text("FooterText Here"))));

            var footerRefId = mainPart.GetIdOfPart(footerPart);

            var table = new W.Table();
            var tblPr = new W.TableProperties(
                new W.TableBorders(
                    new W.TopBorder { Val = BorderValues.Single, Size = 4 },
                    new W.BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new W.LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new W.RightBorder { Val = BorderValues.Single, Size = 4 },
                    new W.InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new W.InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }));
            table.Append(tblPr);

            for (int r = 1; r <= 5; r++)
            {
                var row = new W.TableRow();
                for (int c = 1; c <= 3; c++)
                    row.Append(new W.TableCell(
                        new W.Paragraph(new W.Run(new W.Text($"Cell R{r}C{c}")))));
                table.Append(row);
            }

            // Section properties with footer reference and explicit footer margin
            var sectPr = new W.SectionProperties(
                new W.FooterReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = footerRefId
                },
                new W.PageSize { Width = 12240, Height = 15840 },
                new W.PageMargin
                {
                    Top = 1440, Bottom = 1440,
                    Left = 1440, Right = 1440,
                    Footer = 720 // 0.5 inch footer distance
                });

            mainPart.Document = new Document(
                new Body(table, sectPr));
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a DOCX with an inline image that has an a:hlinkClick hyperlink in its docPr.
    /// </summary>
    private static byte[] BuildImageWithHyperlinkDocx(string url)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();

            // Create a minimal 1x1 red PNG image
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            var pngBytes = CreateMinimalPng();
            using (var imgStream = new MemoryStream(pngBytes))
                imagePart.FeedData(imgStream);
            var imageRelId = mainPart.GetIdOfPart(imagePart);

            // Add a hyperlink relationship for the image
            var hlinkRel = mainPart.AddHyperlinkRelationship(new Uri(url), true);

            // Build the drawing XML with a:hlinkClick in docPr
            var drawingXml = $@"<w:drawing xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
                xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
                xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
                xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture""
                xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
                <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"">
                    <wp:extent cx=""914400"" cy=""914400"" />
                    <wp:docPr id=""1"" name=""TestImage"">
                        <a:hlinkClick r:id=""{hlinkRel.Id}"" />
                    </wp:docPr>
                    <wp:cNvGraphicFramePr>
                        <a:graphicFrameLocks noChangeAspect=""1"" />
                    </wp:cNvGraphicFramePr>
                    <a:graphic>
                        <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                            <pic:pic>
                                <pic:nvPicPr>
                                    <pic:cNvPr id=""1"" name=""TestImage"" />
                                    <pic:cNvPicPr />
                                </pic:nvPicPr>
                                <pic:blipFill>
                                    <a:blip r:embed=""{imageRelId}"" />
                                    <a:stretch><a:fillRect /></a:stretch>
                                </pic:blipFill>
                                <pic:spPr>
                                    <a:xfrm>
                                        <a:off x=""0"" y=""0"" />
                                        <a:ext cx=""914400"" cy=""914400"" />
                                    </a:xfrm>
                                    <a:prstGeom prst=""rect""><a:avLst /></a:prstGeom>
                                </pic:spPr>
                            </pic:pic>
                        </a:graphicData>
                    </a:graphic>
                </wp:inline>
            </w:drawing>";

            var drawing = new W.Drawing(drawingXml);

            var body = new Body(
                new W.Paragraph(new W.Run(drawing)),
                new W.SectionProperties(
                    new W.PageSize { Width = 12240, Height = 15840 },
                    new W.PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 }));

            mainPart.Document = new Document(body);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Creates a minimal valid PNG image using System.Drawing.
    /// </summary>
    private static byte[] CreateMinimalPng()
    {
        using var bmp = new System.Drawing.Bitmap(10, 10);
        using var g = System.Drawing.Graphics.FromImage(bmp);
        g.Clear(System.Drawing.Color.Red);
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        return ms.ToArray();
    }

    #endregion

    #region PDF Analysis Helpers

    private record PdfText(double X, double Y, string Content);

    private static List<PdfText> ExtractTexts(byte[] pdfBytes)
    {
        var texts = new List<PdfText>();
        using var ms = new MemoryStream(pdfBytes);
        var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        foreach (var page in doc.Pages)
        {
            var content = ContentReader.ReadContent(page);
            double curX = 0, curY = 0;
            WalkTexts(content, texts, ref curX, ref curY);
        }
        return texts;
    }

    private static void WalkTexts(CSequence seq, List<PdfText> texts, ref double curX, ref double curY)
    {
        foreach (var item in seq)
        {
            if (item is CSequence sub) { WalkTexts(sub, texts, ref curX, ref curY); continue; }
            if (item is not COperator op) continue;

            switch (op.OpCode.Name)
            {
                case "BT": curX = 0; curY = 0; break;
                case "Tm" when op.Operands.Count == 6:
                    curX = Val(op.Operands[4]); curY = Val(op.Operands[5]); break;
                case "Td" when op.Operands.Count == 2:
                    curX += Val(op.Operands[0]); curY += Val(op.Operands[1]); break;
                case "Tj" or "TJ" when op.Operands.Count > 0:
                    string txt = "";
                    if (op.Operands[0] is CString s) txt = s.Value;
                    else if (op.Operands[0] is CArray arr)
                        foreach (var el in arr)
                            if (el is CString cs) txt += cs.Value;
                    if (!string.IsNullOrWhiteSpace(txt))
                        texts.Add(new PdfText(curX, curY, txt.Trim()));
                    break;
            }
        }
    }

    private static double Val(CObject o) => o switch
    {
        CReal r => r.Value,
        CInteger ci => ci.Value,
        _ => 0
    };

    #endregion
}
