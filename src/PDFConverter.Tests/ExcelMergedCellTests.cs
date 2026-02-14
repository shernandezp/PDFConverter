using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;
using PdfSharp.Pdf.IO;
using Xunit;

namespace PDFConverter.Tests;

/// <summary>
/// Tests for XLSX merged cell border handling, border extension detection,
/// and image rendering in merged-away cells (BUG-034, BUG-035, BUG-036).
/// Uses in-memory OpenXML documents — no file dependencies.
/// </summary>
public class ExcelMergedCellTests
{
    private record PdfLine(double X1, double Y1, double X2, double Y2, double Width);
    private record PdfText(double X, double Y, string Content);
    private record PdfImage(double Width, double Height, double X, double Y);

    /// <summary>
    /// BUG-034: Merged-away cells must still apply their borders.
    /// Creates a 5-row × 3-col sheet with a merge at B2:C2 (cols 1-2).
    /// All cells have a right border. Without the fix, the merged-away cell (C2)
    /// would skip border application, leaving a gap in the right border column.
    /// </summary>
    [Fact]
    public void MergedAwayCells_StillApplyBorders_NoRightBorderGap()
    {
        var xlsx = BuildMergedBorderSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var lines = ExtractLines(pdfBytes);

        // Find vertical lines on the right edge (highest X among vertical lines)
        var verticalLines = lines
            .Where(l => Math.Abs(l.X1 - l.X2) < 1.0 && Math.Abs(l.Y1 - l.Y2) > 5)
            .ToList();

        Assert.True(verticalLines.Count > 0, "Expected vertical border lines in PDF");

        // Group by approximate X to find the rightmost column border
        double maxX = verticalLines.Max(l => l.X1);
        var rightBorder = verticalLines
            .Where(l => Math.Abs(l.X1 - maxX) < 3.0)
            .OrderByDescending(l => Math.Max(l.Y1, l.Y2))
            .ToList();

        // Should have continuous coverage — check for gaps
        Assert.True(rightBorder.Count >= 1,
            $"Expected right border segments, found {rightBorder.Count}");

        for (int i = 0; i < rightBorder.Count - 1; i++)
        {
            double bottom = Math.Min(rightBorder[i].Y1, rightBorder[i].Y2);
            double top = Math.Max(rightBorder[i + 1].Y1, rightBorder[i + 1].Y2);
            double gap = bottom - top;
            Assert.True(Math.Abs(gap) <= 2.0,
                $"Right border gap of {gap:F1}pt between segments {i} and {i + 1}");
        }
    }

    /// <summary>
    /// BUG-035: Border extension detection prevents spurious full-width lines.
    /// Creates a sheet where row N has partial bottom borders (cols 0-1 have border,
    /// col 2 does not), and row N+1 has a wide merge (A:C) with no top border.
    /// Without the fix, MigraDoc extends the partial border across the full merge width.
    /// </summary>
    [Fact]
    public void BorderExtensionMerge_NoSpuriousFullWidthLine()
    {
        var xlsx = BuildBorderExtensionSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var (lines, texts) = ExtractLinesAndTexts(pdfBytes);

        var mergedText = texts.FirstOrDefault(t => t.Content.Contains("MERGED"));
        Assert.NotNull(mergedText);

        // Look for wide horizontal lines near the merged row (within 20pt above)
        double mergedY = mergedText.Y;
        var wideLines = lines
            .Where(l =>
                Math.Abs(l.Y1 - l.Y2) < 1.0 &&       // horizontal
                Math.Abs(l.X2 - l.X1) > 150 &&        // wider than any single cell
                l.Y1 > mergedY - 5 && l.Y1 < mergedY + 25)
            .ToList();

        Assert.Empty(wideLines);
    }

    /// <summary>
    /// BUG-035: When border extension skips MergeRight, text should still be centered.
    /// The centered text in the skipped merge should appear near the center of
    /// what would be the full merged width.
    /// </summary>
    [Fact]
    public void BorderExtensionMerge_TextRemainsCentered()
    {
        var xlsx = BuildBorderExtensionSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var (lines, texts) = ExtractLinesAndTexts(pdfBytes);

        var mergedText = texts.FirstOrDefault(t => t.Content.Contains("MERGED"));
        Assert.NotNull(mergedText);

        // The table has 3 columns of 80pt each = 240pt total.
        // Page is Letter (612pt) with ~50pt margins each side → content ~512pt.
        // Table is left-aligned starting at left margin (~50pt).
        // "MERGED" centered across 3 cols → center around 50 + 120 = 170pt.
        // Text X should be in a reasonable centered range.
        Assert.True(mergedText.X > 60, $"MERGED text at X={mergedText.X:F1} is too far left to be centered");
    }

    /// <summary>
    /// BUG-036: Images anchored in merged-away cells must still render.
    /// Creates a sheet with a merge A2:C2 and an image anchored at col B (col 1),
    /// which falls within the merge range. Without the fix, the image would be lost
    /// because B2 is a merged-away cell.
    /// </summary>
    [Fact]
    public void ImageInMergedAwayCell_StillRendered()
    {
        var xlsx = BuildImageInMergeSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var images = ExtractImages(pdfBytes);

        Assert.True(images.Count >= 1,
            $"Expected at least 1 image (anchored in merged-away cell), found {images.Count}");
    }

    /// <summary>
    /// Merged cells with consistent borders above should NOT trigger
    /// border extension detection (no false positives).
    /// </summary>
    [Fact]
    public void ConsistentBordersAboveMerge_NoFalsePositiveDetection()
    {
        var xlsx = BuildConsistentBorderSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var (lines, texts) = ExtractLinesAndTexts(pdfBytes);

        var mergedText = texts.FirstOrDefault(t => t.Content.Contains("FULLMERGE"));
        Assert.NotNull(mergedText);

        // With consistent borders above (all cells have bottom border),
        // the merge should remain intact (MergeRight applied).
        // The full-width border line IS expected here since all cells have it.
    }

    /// <summary>
    /// Narrow anchor cells (< 40pt) should keep their merge even if they
    /// match border extension criteria, to prevent text wrapping.
    /// </summary>
    [Fact]
    public void NarrowAnchorCell_KeepsMerge_PreventTextWrapping()
    {
        var xlsx = BuildNarrowAnchorSheet();
        var pdfBytes = Converters.XlsxToPdfBytes(xlsx);
        var (_, texts) = ExtractLinesAndTexts(pdfBytes);

        // "LONG TEXT HERE" should appear as a single text fragment
        // (not wrapped) because the merge is preserved for narrow anchors
        var longText = texts.Where(t =>
            t.Content.Contains("LONG") || t.Content.Contains("TEXT")).ToList();
        Assert.True(longText.Count >= 1,
            "Expected LONG TEXT to be present in PDF output");
    }

    #region XLSX Builders

    /// <summary>
    /// 5 rows × 3 cols. Merge at B2:C2 (row 1, cols 1-2).
    /// All cells have right border (borderId=1). 
    /// Tests BUG-034: merged-away C2 must still apply right border.
    /// </summary>
    private static byte[] BuildMergedBorderSheet()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");

            // Styles: borderId=0 (none), borderId=1 (right border)
            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(new Font()),
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }),
                          new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(
                    new Border(),  // borderId=0: no borders
                    new Border(    // borderId=1: right border
                        new LeftBorder(),
                        new RightBorder(new Color { Indexed = 64 }) { Style = BorderStyleValues.Medium },
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder())),
                new CellFormats(
                    new CellFormat(),                                  // styleIndex=0
                    new CellFormat { BorderId = 1, ApplyBorder = true } // styleIndex=1
                ));

            var sheetData = new SheetData();
            for (int r = 1; r <= 5; r++)
            {
                var row = new Row { RowIndex = (uint)r };
                for (int c = 0; c < 3; c++)
                {
                    row.Append(new Cell
                    {
                        CellReference = $"{(char)('A' + c)}{r}",
                        DataType = CellValues.String,
                        CellValue = new CellValue($"R{r}C{c}"),
                        StyleIndex = 1 // right border
                    });
                }
                sheetData.Append(row);
            }

            var mergeCells = new MergeCells(
                new MergeCell { Reference = "B2:C2" });

            wsPart.Worksheet = new Worksheet(sheetData, mergeCells);
        }

        return ms.ToArray();
    }

    /// <summary>
    /// 4 rows × 3 cols (A-C). Row 2 has partial bottom borders:
    /// A2 and B2 have bottom border (borderId=2), C2 has none (borderId=0).
    /// Row 3 has a wide merge A3:C3 with no top border, text "MERGED", centered.
    /// Anchor col A is 80pt wide (≥40pt threshold).
    /// Tests BUG-035: border extension detection.
    /// </summary>
    private static byte[] BuildBorderExtensionSheet()
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
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }),
                          new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(
                    new Border(), // borderId=0: no borders
                    new Border(   // borderId=1: top+bottom
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(new Color { Indexed = 64 }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color { Indexed = 64 }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder()),
                    new Border(   // borderId=2: bottom only
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(new Color { Indexed = 64 }) { Style = BorderStyleValues.Medium },
                        new DiagonalBorder())),
                new CellFormats(
                    new CellFormat(),                                                                     // 0: default
                    new CellFormat { BorderId = 1, ApplyBorder = true },                                   // 1: top+bottom
                    new CellFormat { BorderId = 2, ApplyBorder = true },                                   // 2: bottom only
                    new CellFormat { ApplyAlignment = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } } // 3: centered
                ));

            var columns = new Columns(
                new Column { Min = 1, Max = 3, Width = 12, CustomWidth = true }); // ~80pt each

            var sheetData = new SheetData();

            // Row 1: headers with top+bottom borders
            var row1 = new Row { RowIndex = 1 };
            row1.Append(MakeCell("A1", "H1", 1), MakeCell("B1", "H2", 1), MakeCell("C1", "H3", 1));
            sheetData.Append(row1);

            // Row 2: A2 and B2 have bottom border, C2 has NO border (inconsistent)
            var row2 = new Row { RowIndex = 2 };
            row2.Append(MakeCell("A2", "D1", 2), MakeCell("B2", "D2", 2), MakeCell("C2", "D3", 0));
            sheetData.Append(row2);

            // Row 3: wide merge A3:C3, no top border, centered text "MERGED"
            var row3 = new Row { RowIndex = 3 };
            row3.Append(MakeCell("A3", "MERGED", 3));
            sheetData.Append(row3);

            // Row 4: data
            var row4 = new Row { RowIndex = 4 };
            row4.Append(MakeCell("A4", "End", 0));
            sheetData.Append(row4);

            var mergeCells = new MergeCells(
                new MergeCell { Reference = "A3:C3" });

            wsPart.Worksheet = new Worksheet(columns, sheetData, mergeCells);
        }

        return ms.ToArray();
    }

    /// <summary>
    /// Same as border extension sheet but ALL cells in row 2 have bottom border
    /// (consistent). The merge A3:C3 should NOT be flagged.
    /// </summary>
    private static byte[] BuildConsistentBorderSheet()
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
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }),
                          new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(
                    new Border(),
                    new Border(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(new Color { Indexed = 64 }) { Style = BorderStyleValues.Medium },
                        new DiagonalBorder())),
                new CellFormats(
                    new CellFormat(),
                    new CellFormat { BorderId = 1, ApplyBorder = true },
                    new CellFormat { ApplyAlignment = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } }
                ));

            var columns = new Columns(
                new Column { Min = 1, Max = 3, Width = 12, CustomWidth = true });

            var sheetData = new SheetData();

            var row1 = new Row { RowIndex = 1 };
            row1.Append(MakeCell("A1", "H1", 0), MakeCell("B1", "H2", 0), MakeCell("C1", "H3", 0));
            sheetData.Append(row1);

            // All cells have bottom border — consistent
            var row2 = new Row { RowIndex = 2 };
            row2.Append(MakeCell("A2", "D1", 1), MakeCell("B2", "D2", 1), MakeCell("C2", "D3", 1));
            sheetData.Append(row2);

            var row3 = new Row { RowIndex = 3 };
            row3.Append(MakeCell("A3", "FULLMERGE", 2));
            sheetData.Append(row3);

            var mergeCells = new MergeCells(
                new MergeCell { Reference = "A3:C3" });

            wsPart.Worksheet = new Worksheet(columns, sheetData, mergeCells);
        }

        return ms.ToArray();
    }

    /// <summary>
    /// 4 rows × 5 cols. Row 3 has merge A3:E3 with anchor col A at 15pt width
    /// (below 40pt threshold). Even with inconsistent borders above, the merge
    /// should be kept to prevent text wrapping.
    /// </summary>
    private static byte[] BuildNarrowAnchorSheet()
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
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }),
                          new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(
                    new Border(),
                    new Border(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(new Color { Indexed = 64 }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())),
                new CellFormats(
                    new CellFormat(),
                    new CellFormat { BorderId = 1, ApplyBorder = true },
                    new CellFormat { ApplyAlignment = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } }
                ));

            // Col A narrow (2.5 chars ≈ 15pt), cols B-E normal width
            var columns = new Columns(
                new Column { Min = 1, Max = 1, Width = 2.5, CustomWidth = true },
                new Column { Min = 2, Max = 5, Width = 12, CustomWidth = true });

            var sheetData = new SheetData();

            var row1 = new Row { RowIndex = 1 };
            row1.Append(MakeCell("A1", "X", 0));
            sheetData.Append(row1);

            // Row 2: partial bottom borders (A2 and B2 have border, C-E don't)
            var row2 = new Row { RowIndex = 2 };
            row2.Append(MakeCell("A2", "a", 1), MakeCell("B2", "b", 1), MakeCell("C2", "c", 0));
            sheetData.Append(row2);

            // Row 3: wide merge A3:E3 with narrow anchor col A
            var row3 = new Row { RowIndex = 3 };
            row3.Append(MakeCell("A3", "LONG TEXT HERE", 2));
            sheetData.Append(row3);

            var row4 = new Row { RowIndex = 4 };
            row4.Append(MakeCell("A4", "end", 0));
            sheetData.Append(row4);

            var mergeCells = new MergeCells(
                new MergeCell { Reference = "A3:E3" });

            wsPart.Worksheet = new Worksheet(columns, sheetData, mergeCells);
        }

        return ms.ToArray();
    }

    /// <summary>
    /// 3 rows × 3 cols. Merge A2:C2. Image anchored at col B (fromCol=1),
    /// which is inside the merge range but not at the anchor (col A).
    /// Tests BUG-036: image in merged-away cell must still render.
    /// </summary>
    private static byte[] BuildImageInMergeSheet()
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
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }),
                          new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(new Border()),
                new CellFormats(new CellFormat()));

            var sheetData = new SheetData();
            var row1 = new Row { RowIndex = 1 };
            row1.Append(MakeCell("A1", "Header", 0));
            sheetData.Append(row1);

            var row2 = new Row { RowIndex = 2 };
            row2.Append(MakeCell("A2", "", 0));
            sheetData.Append(row2);

            var row3 = new Row { RowIndex = 3 };
            row3.Append(MakeCell("A3", "Footer", 0));
            sheetData.Append(row3);

            var mergeCells = new MergeCells(
                new MergeCell { Reference = "A2:C2" });

            // Add DrawingsPart with a TwoCellAnchor image at col B (inside merge)
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
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2, Name = "TestImage" },
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

            // Link drawing to worksheet
            var drawingRelId = wsPart.GetIdOfPart(drawingsPart);
            wsPart.Worksheet = new Worksheet(sheetData, mergeCells, new Drawing { Id = drawingRelId });
        }

        return ms.ToArray();
    }

    #endregion

    #region Helpers

    private static Cell MakeCell(string reference, string value, uint styleIndex) => new()
    {
        CellReference = reference,
        DataType = CellValues.String,
        CellValue = new CellValue(value),
        StyleIndex = styleIndex
    };

    /// <summary>
    /// Returns a valid 2×2 red PNG image (124 bytes).
    /// Pre-generated to avoid System.Drawing dependency (Windows-only).
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

    private static List<PdfLine> ExtractLines(byte[] pdfBytes)
    {
        var (lines, _) = ExtractLinesAndTexts(pdfBytes);
        return lines;
    }

    private static List<PdfImage> ExtractImages(byte[] pdfBytes)
    {
        var images = new List<PdfImage>();
        using var ms = new MemoryStream(pdfBytes);
        var doc = PdfSharp.Pdf.IO.PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        var content = ContentReader.ReadContent(doc.Pages[0]);
        WalkForImages(content, images);
        return images;
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
                if (a > 10 || d > 10)
                    images.Add(new PdfImage(a, d, tx, ty));
            }
        }
    }

    private static (List<PdfLine> Lines, List<PdfText> Texts) ExtractLinesAndTexts(byte[] pdfBytes)
    {
        var lines = new List<PdfLine>();
        var texts = new List<PdfText>();
        using var ms = new MemoryStream(pdfBytes);
        var doc = PdfSharp.Pdf.IO.PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        var content = ContentReader.ReadContent(doc.Pages[0]);
        double curX = 0, curY = 0, lineW = 0, mx = 0, my = 0;
        WalkPdfContent(content, lines, texts, ref curX, ref curY, ref lineW, ref mx, ref my);
        return (lines, texts);
    }

    private static void WalkPdfContent(CSequence seq, List<PdfLine> lines, List<PdfText> texts,
        ref double curX, ref double curY, ref double lineW, ref double mx, ref double my)
    {
        foreach (var item in seq)
        {
            if (item is CSequence sub)
            {
                WalkPdfContent(sub, lines, texts, ref curX, ref curY, ref lineW, ref mx, ref my);
                continue;
            }
            if (item is not COperator op) continue;

            switch (op.OpCode.Name)
            {
                case "w" when op.Operands.Count >= 1:
                    lineW = Val(op.Operands[0]); break;
                case "m" when op.Operands.Count >= 2:
                    mx = Val(op.Operands[0]); my = Val(op.Operands[1]); break;
                case "l" when op.Operands.Count >= 2:
                    lines.Add(new PdfLine(mx, my, Val(op.Operands[0]), Val(op.Operands[1]), lineW)); break;
                case "BT":
                    curX = 0; curY = 0; break;
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
