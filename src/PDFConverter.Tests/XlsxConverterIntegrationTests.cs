using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;
using PdfSharp.Pdf.IO;
using Xunit;

namespace PDFConverter.Tests;

/// <summary>
/// Integration tests for XLSX → PDF conversion using in-memory OpenXML documents.
/// No external file dependencies — all test documents are built programmatically.
/// </summary>
public class XlsxConverterIntegrationTests : IDisposable
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

    private record PdfImage(double Width, double Height, double X, double Y);
    private record PdfText(double X, double Y, string Content);

    private static void AssertValidPdf(byte[] pdfBytes)
    {
        Assert.NotNull(pdfBytes);
        Assert.True(pdfBytes.Length > 100, "PDF too small to be valid");
        Assert.Equal((byte)'%', pdfBytes[0]);
        Assert.Equal((byte)'P', pdfBytes[1]);
        Assert.Equal((byte)'D', pdfBytes[2]);
        Assert.Equal((byte)'F', pdfBytes[3]);
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

    #region Basic Conversion

    /// <summary>
    /// Converts a simple 3×3 XLSX spreadsheet to PDF and verifies output is valid.
    /// </summary>
    [Fact]
    public void XlsxToPdf_SimpleSheet_ProducesValidPdf()
    {
        var xlsx = BuildSimpleSheet(3, 3);
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);

        AssertValidPdf(pdfBytes);
        Assert.Equal(1, GetPdfPageCount(pdfBytes));
    }

    /// <summary>
    /// Cell content from the spreadsheet should appear in the PDF output.
    /// </summary>
    [Fact]
    public void XlsxToPdf_SimpleSheet_ContainsCellText()
    {
        var xlsx = BuildSimpleSheet(3, 3);
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var texts = ExtractTexts(pdfBytes);

        // Check for cell content — at least one cell should be rendered
        Assert.Contains(texts, t => t.Content.Contains("R1C1"));
        Assert.Contains(texts, t => t.Content.Contains("R3C3"));
    }

    /// <summary>
    /// An empty spreadsheet should still produce valid PDF.
    /// </summary>
    [Fact]
    public void XlsxToPdf_EmptySheet_ProducesValidPdf()
    {
        var xlsx = BuildSimpleSheet(0, 0);
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);

        AssertValidPdf(pdfBytes);
    }

    #endregion

    #region Styled Cells

    /// <summary>
    /// A spreadsheet with cell borders should produce a PDF containing
    /// line segments (border rendering).
    /// </summary>
    [Fact]
    public void XlsxToPdf_CellBorders_ProducesLinesInPdf()
    {
        var xlsx = BuildBorderedSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var lines = ExtractLines(pdfBytes);

        AssertValidPdf(pdfBytes);
        Assert.True(lines.Count > 0, "Expected border lines in PDF output");
    }

    /// <summary>
    /// A spreadsheet with fill colors should convert without errors.
    /// </summary>
    [Fact]
    public void XlsxToPdf_CellFills_ProducesValidPdf()
    {
        var xlsx = BuildFilledSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);

        AssertValidPdf(pdfBytes);
    }

    /// <summary>
    /// Different font sizes should render text correctly.
    /// </summary>
    [Fact]
    public void XlsxToPdf_MultipleFontSizes_ProducesValidPdf()
    {
        var xlsx = BuildMultiFontSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var texts = ExtractTexts(pdfBytes);

        AssertValidPdf(pdfBytes);
        Assert.Contains(texts, t => t.Content.Contains("Small"));
        Assert.Contains(texts, t => t.Content.Contains("Large"));
    }

    #endregion

    #region Merged Cells

    /// <summary>
    /// A spreadsheet with horizontally merged cells should render correctly
    /// with the merged text appearing in the PDF.
    /// </summary>
    [Fact]
    public void XlsxToPdf_MergedCells_TextIsPresent()
    {
        var xlsx = BuildMergedSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var texts = ExtractTexts(pdfBytes);

        AssertValidPdf(pdfBytes);
        // PDF may split multi-word text into separate fragments
        Assert.Contains(texts, t => t.Content.Contains("Merged"));
    }

    /// <summary>
    /// Vertical merge (across rows) should not crash and text should render.
    /// </summary>
    [Fact]
    public void XlsxToPdf_VerticalMerge_ProducesValidPdf()
    {
        var xlsx = BuildVerticalMergeSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var texts = ExtractTexts(pdfBytes);

        AssertValidPdf(pdfBytes);
        // PDF may split multi-word text into separate fragments
        Assert.Contains(texts, t => t.Content.Contains("Tall"));
    }

    #endregion

    #region Image Handling

    /// <summary>
    /// A spreadsheet with an embedded image should render the image in the PDF.
    /// </summary>
    [Fact]
    public void XlsxToPdf_WithImage_ContainsImageInPdf()
    {
        var xlsx = BuildSheetWithImage();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var images = ExtractImages(pdfBytes);

        AssertValidPdf(pdfBytes);
        Assert.True(images.Count >= 1,
            $"Expected at least 1 image in PDF, found {images.Count}");
    }

    /// <summary>
    /// Image dimensions in the PDF should be non-zero, confirming the image was rendered.
    /// The actual size depends on cell span calculations, not just EMU extents.
    /// </summary>
    [Fact]
    public void XlsxToPdf_WithImage_ImageHasReasonableDimensions()
    {
        var xlsx = BuildSheetWithImage();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var images = ExtractImages(pdfBytes);

        Assert.True(images.Count >= 1);
        var img = images[0];
        // Image should have non-zero dimensions
        Assert.True(img.Width > 1, $"Image width too small: {img.Width:F1}");
        Assert.True(img.Height > 1, $"Image height too small: {img.Height:F1}");
    }

    #endregion

    #region All Overloads

    /// <summary>
    /// All three overloads (byte[], Stream, file path) should produce
    /// identical page counts from the same in-memory XLSX.
    /// </summary>
    [Fact]
    public void XlsxToPdf_AllOverloads_ProduceSamePageCount()
    {
        var xlsx = BuildSimpleSheet(3, 3);

        // Byte array overload
        var pdfFromBytes = Converters.XlsxToPdfBytes(xlsx);
        int countBytes = GetPdfPageCount(pdfFromBytes);

        // Stream overload
        using var stream = new MemoryStream(xlsx);
        var pdfFromStream = Converters.XlsxToPdfBytes(stream);
        int countStream = GetPdfPageCount(pdfFromStream);

        // File path overload
        var tempXlsx = Path.GetTempFileName() + ".xlsx";
        var tempPdf = GetOutputPath("AllOverloads_File");
        try
        {
            File.WriteAllBytes(tempXlsx, xlsx);
            XlsxConverter.XlsxToPdf(tempXlsx, tempPdf);
            int countFile = GetPdfPageCount(tempPdf);

            Assert.Equal(countBytes, countStream);
            Assert.Equal(countBytes, countFile);
        }
        finally
        {
            try { File.Delete(tempXlsx); } catch { }
        }
    }

    /// <summary>
    /// XlsxToPdfBytes from a Stream should return a valid PDF with %PDF header.
    /// </summary>
    [Fact]
    public void XlsxToPdfBytes_FromStream_ReturnsValidPdf()
    {
        var xlsx = BuildSimpleSheet(2, 2);
        using var stream = new MemoryStream(xlsx);
        var pdfBytes = Converters.XlsxToPdfBytes(stream);

        AssertValidPdf(pdfBytes);
    }

    /// <summary>
    /// XlsxToPdfBytes from byte[] should return a valid PDF with %PDF header.
    /// </summary>
    [Fact]
    public void XlsxToPdfBytes_FromByteArray_ReturnsValidPdf()
    {
        var xlsx = BuildSimpleSheet(2, 2);
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);

        AssertValidPdf(pdfBytes);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void XlsxToPdf_NonExistentFile_ThrowsException()
    {
        var output = GetOutputPath("NonExistent");
        Assert.ThrowsAny<Exception>(() =>
            XlsxConverter.XlsxToPdf(@"C:\nonexistent_xlsx_12345.xlsx", output));
    }

    #endregion

    #region XLSX Builders

    /// <summary>
    /// Builds a simple rows×cols spreadsheet with cell values "R{r}C{c}".
    /// </summary>
    private static byte[] BuildSimpleSheet(int rows, int cols)
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");

            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateMinimalStylesheet();

            var sheetData = new SheetData();
            for (int r = 1; r <= rows; r++)
            {
                var row = new Row { RowIndex = (uint)r };
                for (int c = 1; c <= cols; c++)
                {
                    row.Append(new Cell
                    {
                        CellReference = $"{(char)('A' + c - 1)}{r}",
                        DataType = CellValues.String,
                        CellValue = new CellValue($"R{r}C{c}")
                    });
                }
                sheetData.Append(row);
            }

            wsPart.Worksheet = new Worksheet(sheetData);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a 3×3 sheet with all-around borders on every cell.
    /// </summary>
    private static byte[] BuildBorderedSheet()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");

            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(new Font()),
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(
                    new Border(), // 0: no borders
                    new Border(   // 1: all borders
                        new LeftBorder(new Color { Indexed = 64 }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color { Indexed = 64 }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color { Indexed = 64 }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color { Indexed = 64 }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())),
                new CellFormats(
                    new CellFormat(),
                    new CellFormat { BorderId = 1, ApplyBorder = true }));

            var sheetData = new SheetData();
            for (int r = 1; r <= 3; r++)
            {
                var row = new Row { RowIndex = (uint)r };
                for (int c = 0; c < 3; c++)
                {
                    row.Append(new Cell
                    {
                        CellReference = $"{(char)('A' + c)}{r}",
                        DataType = CellValues.String,
                        CellValue = new CellValue($"B{r}{c}"),
                        StyleIndex = 1
                    });
                }
                sheetData.Append(row);
            }

            wsPart.Worksheet = new Worksheet(sheetData);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a 2×3 sheet where the first row has a blue background fill.
    /// </summary>
    private static byte[] BuildFilledSheet()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");

            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(new Font()),
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                    new Fill(new PatternFill(new ForegroundColor { Rgb = "FF4472C4" })
                        { PatternType = PatternValues.Solid })),
                new Borders(new Border()),
                new CellFormats(
                    new CellFormat(),
                    new CellFormat { FillId = 2, ApplyFill = true }));

            var sheetData = new SheetData();

            // Header row with fill
            var headerRow = new Row { RowIndex = 1 };
            for (int c = 0; c < 3; c++)
                headerRow.Append(new Cell
                {
                    CellReference = $"{(char)('A' + c)}1",
                    DataType = CellValues.String,
                    CellValue = new CellValue($"Header{c + 1}"),
                    StyleIndex = 1
                });
            sheetData.Append(headerRow);

            // Data row
            var dataRow = new Row { RowIndex = 2 };
            for (int c = 0; c < 3; c++)
                dataRow.Append(new Cell
                {
                    CellReference = $"{(char)('A' + c)}2",
                    DataType = CellValues.String,
                    CellValue = new CellValue($"Data{c + 1}")
                });
            sheetData.Append(dataRow);

            wsPart.Worksheet = new Worksheet(sheetData);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a 3×2 sheet with different font sizes (8pt and 16pt).
    /// </summary>
    private static byte[] BuildMultiFontSheet()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");

            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(
                    new Font(), // 0: default
                    new Font(new FontSize { Val = 8 }),  // 1: small
                    new Font(new FontSize { Val = 16 })), // 2: large
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(new Border()),
                new CellFormats(
                    new CellFormat(),
                    new CellFormat { FontId = 1, ApplyFont = true },
                    new CellFormat { FontId = 2, ApplyFont = true }));

            var sheetData = new SheetData();

            var row1 = new Row { RowIndex = 1 };
            row1.Append(new Cell { CellReference = "A1", DataType = CellValues.String,
                CellValue = new CellValue("Small"), StyleIndex = 1 });
            sheetData.Append(row1);

            var row2 = new Row { RowIndex = 2 };
            row2.Append(new Cell { CellReference = "A2", DataType = CellValues.String,
                CellValue = new CellValue("Large"), StyleIndex = 2 });
            sheetData.Append(row2);

            wsPart.Worksheet = new Worksheet(sheetData);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a 3×3 sheet with a horizontal merge at A1:C1 containing "Merged Title".
    /// </summary>
    private static byte[] BuildMergedSheet()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");

            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateMinimalStylesheet();

            var sheetData = new SheetData();

            // Row 1: merged A1:C1
            var row1 = new Row { RowIndex = 1 };
            row1.Append(new Cell { CellReference = "A1", DataType = CellValues.String,
                CellValue = new CellValue("Merged Title") });
            sheetData.Append(row1);

            // Data rows
            for (int r = 2; r <= 3; r++)
            {
                var row = new Row { RowIndex = (uint)r };
                for (int c = 0; c < 3; c++)
                    row.Append(new Cell { CellReference = $"{(char)('A' + c)}{r}",
                        DataType = CellValues.String, CellValue = new CellValue($"D{r}{c}") });
                sheetData.Append(row);
            }

            var mergeCells = new MergeCells(new MergeCell { Reference = "A1:C1" });
            wsPart.Worksheet = new Worksheet(sheetData, mergeCells);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a 4×2 sheet with a vertical merge at A1:A3 containing "Tall Cell".
    /// </summary>
    private static byte[] BuildVerticalMergeSheet()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");

            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateMinimalStylesheet();

            var sheetData = new SheetData();

            var row1 = new Row { RowIndex = 1 };
            row1.Append(
                new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("Tall Cell") },
                new Cell { CellReference = "B1", DataType = CellValues.String, CellValue = new CellValue("Right1") });
            sheetData.Append(row1);

            var row2 = new Row { RowIndex = 2 };
            row2.Append(
                new Cell { CellReference = "B2", DataType = CellValues.String, CellValue = new CellValue("Right2") });
            sheetData.Append(row2);

            var row3 = new Row { RowIndex = 3 };
            row3.Append(
                new Cell { CellReference = "B3", DataType = CellValues.String, CellValue = new CellValue("Right3") });
            sheetData.Append(row3);

            var row4 = new Row { RowIndex = 4 };
            row4.Append(
                new Cell { CellReference = "A4", DataType = CellValues.String, CellValue = new CellValue("Below") },
                new Cell { CellReference = "B4", DataType = CellValues.String, CellValue = new CellValue("Right4") });
            sheetData.Append(row4);

            var mergeCells = new MergeCells(new MergeCell { Reference = "A1:A3" });
            wsPart.Worksheet = new Worksheet(sheetData, mergeCells);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a 3×3 sheet with an embedded PNG image at B2 (TwoCellAnchor).
    /// </summary>
    private static byte[] BuildSheetWithImage()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");

            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateMinimalStylesheet();

            var sheetData = new SheetData();
            for (int r = 1; r <= 3; r++)
            {
                var row = new Row { RowIndex = (uint)r };
                for (int c = 0; c < 3; c++)
                    row.Append(new Cell { CellReference = $"{(char)('A' + c)}{r}",
                        DataType = CellValues.String, CellValue = new CellValue($"C{r}{c}") });
                sheetData.Append(row);
            }

            // Add image
            var drawingsPart = wsPart.AddNewPart<DrawingsPart>();
            var imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            imagePart.FeedData(new MemoryStream(CreateMinimalPng()));
            string imageRelId = drawingsPart.GetIdOfPart(imagePart);

            drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor(
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("1"),
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("1"),
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")),
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("2"),
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("2"),
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")),
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture(
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties(
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2, Name = "TestImg" },
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties()),
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill(
                            new DocumentFormat.OpenXml.Drawing.Blip { Embed = imageRelId },
                            new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties(
                            new DocumentFormat.OpenXml.Drawing.Transform2D(
                                new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                                new DocumentFormat.OpenXml.Drawing.Extents { Cx = 914400, Cy = 914400 }),
                            new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                new DocumentFormat.OpenXml.Drawing.AdjustValueList()) { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle })),
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData()));

            var drawingRelId = wsPart.GetIdOfPart(drawingsPart);
            wsPart.Worksheet = new Worksheet(sheetData, new Drawing { Id = drawingRelId });
        }
        return ms.ToArray();
    }

    #endregion

    #region Shared Helpers

    private static Stylesheet CreateMinimalStylesheet() => new(
        new Fonts(new Font()),
        new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
        new Borders(new Border()),
        new CellFormats(new CellFormat()));

    /// <summary>
    /// Valid 2×2 red PNG image (124 bytes). Pre-generated for cross-platform compatibility.
    /// </summary>
    private static byte[] CreateMinimalPng() =>
    [
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
        0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x02,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x72, 0xB6, 0x0D, 0x24, 0x00, 0x00, 0x00,
        0x01, 0x73, 0x52, 0x47, 0x42, 0x00, 0xAE, 0xCE, 0x1C, 0xE9, 0x00, 0x00,
        0x00, 0x04, 0x67, 0x41, 0x4D, 0x41, 0x00, 0x00, 0xB1, 0x8F, 0x0B, 0xFC,
        0x61, 0x05, 0x00, 0x00, 0x00, 0x09, 0x70, 0x48, 0x59, 0x73, 0x00, 0x00,
        0x0E, 0xC3, 0x00, 0x00, 0x0E, 0xC3, 0x01, 0xC7, 0x6F, 0xA8, 0x64, 0x00,
        0x00, 0x00, 0x11, 0x49, 0x44, 0x41, 0x54, 0x18, 0x57, 0x63, 0xF8, 0xCF,
        0xC0, 0xF0, 0x1F, 0x84, 0x19, 0x60, 0x0C, 0x00, 0x47, 0xCA, 0x07, 0xF9,
        0x3C, 0xE4, 0x09, 0x0A, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
        0xAE, 0x42, 0x60, 0x82
    ];

    private (List<PdfImage> Images, List<PdfText> Texts) AnalyzePdf(byte[] pdfBytes)
    {
        var images = new List<PdfImage>();
        var texts = new List<PdfText>();

        using var ms = new MemoryStream(pdfBytes);
        var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        var content = ContentReader.ReadContent(doc.Pages[0]);

        double curX = 0, curY = 0;
        WalkContent(content, images, texts, ref curX, ref curY);
        return (images, texts);
    }

    private static List<PdfText> ExtractTexts(byte[] pdfBytes)
    {
        var texts = new List<PdfText>();
        using var ms = new MemoryStream(pdfBytes);
        var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        foreach (var page in doc.Pages)
        {
            var content = ContentReader.ReadContent(page);
            double curX = 0, curY = 0;
            WalkForTexts(content, texts, ref curX, ref curY);
        }
        return texts;
    }

    private static List<PdfImage> ExtractImages(byte[] pdfBytes)
    {
        var images = new List<PdfImage>();
        using var ms = new MemoryStream(pdfBytes);
        var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        foreach (var page in doc.Pages)
        {
            var content = ContentReader.ReadContent(page);
            WalkForImages(content, images);
        }
        return images;
    }

    private static List<PdfLine> ExtractLines(byte[] pdfBytes)
    {
        var lines = new List<PdfLine>();
        using var ms = new MemoryStream(pdfBytes);
        var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        foreach (var page in doc.Pages)
        {
            var content = ContentReader.ReadContent(page);
            double lineW = 0, mx = 0, my = 0;
            WalkForLines(content, lines, ref lineW, ref mx, ref my);
        }
        return lines;
    }

    private record PdfLine(double X1, double Y1, double X2, double Y2, double Width);

    private void WalkContent(CSequence seq, List<PdfImage> images, List<PdfText> texts,
        ref double curX, ref double curY)
    {
        for (int i = 0; i < seq.Count; i++)
        {
            if (seq[i] is CSequence sub) { WalkContent(sub, images, texts, ref curX, ref curY); continue; }
            if (seq[i] is not COperator op) continue;

            if (op.OpCode.Name == "cm" && op.Operands.Count == 6)
            {
                double a = Val(op.Operands[0]), d = Val(op.Operands[3]);
                double tx = Val(op.Operands[4]), ty = Val(op.Operands[5]);
                if (a > 10 || d > 10)
                    images.Add(new PdfImage(a, d, tx, ty));
            }
            else if (op.OpCode.Name == "BT") { curX = 0; curY = 0; }
            else if (op.OpCode.Name == "Tm" && op.Operands.Count == 6)
            { curX = Val(op.Operands[4]); curY = Val(op.Operands[5]); }
            else if (op.OpCode.Name == "Td" && op.Operands.Count == 2)
            { curX += Val(op.Operands[0]); curY += Val(op.Operands[1]); }
            else if ((op.OpCode.Name == "Tj" || op.OpCode.Name == "TJ") && op.Operands.Count > 0)
            {
                string txt = "";
                if (op.Operands[0] is CString s) txt = s.Value;
                else if (op.Operands[0] is CArray arr)
                    foreach (var el in arr)
                        if (el is CString cs) txt += cs.Value;
                if (!string.IsNullOrWhiteSpace(txt))
                    texts.Add(new PdfText(curX, curY, txt.Trim()));
            }
        }
    }

    private static void WalkForTexts(CSequence seq, List<PdfText> texts, ref double curX, ref double curY)
    {
        foreach (var item in seq)
        {
            if (item is CSequence sub) { WalkForTexts(sub, texts, ref curX, ref curY); continue; }
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
                        foreach (var el in arr) if (el is CString cs) txt += cs.Value;
                    if (!string.IsNullOrWhiteSpace(txt))
                        texts.Add(new PdfText(curX, curY, txt.Trim()));
                    break;
            }
        }
    }

    private static void WalkForImages(CSequence seq, List<PdfImage> images)
    {
        foreach (var item in seq)
        {
            if (item is CSequence sub) { WalkForImages(sub, images); continue; }
            if (item is COperator op && op.OpCode.Name == "cm" && op.Operands.Count == 6)
            {
                double a = Val(op.Operands[0]), d = Val(op.Operands[3]);
                double tx = Val(op.Operands[4]), ty = Val(op.Operands[5]);
                if (a > 10 || d > 10) images.Add(new PdfImage(a, d, tx, ty));
            }
        }
    }

    private static void WalkForLines(CSequence seq, List<PdfLine> lines,
        ref double lineW, ref double mx, ref double my)
    {
        foreach (var item in seq)
        {
            if (item is CSequence sub) { WalkForLines(sub, lines, ref lineW, ref mx, ref my); continue; }
            if (item is not COperator op) continue;
            switch (op.OpCode.Name)
            {
                case "w" when op.Operands.Count >= 1: lineW = Val(op.Operands[0]); break;
                case "m" when op.Operands.Count >= 2:
                    mx = Val(op.Operands[0]); my = Val(op.Operands[1]); break;
                case "l" when op.Operands.Count >= 2:
                    lines.Add(new PdfLine(mx, my, Val(op.Operands[0]), Val(op.Operands[1]), lineW)); break;
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
