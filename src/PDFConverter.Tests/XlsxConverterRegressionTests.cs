using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace PDFConverter.Tests;

public class XlsxConverterRegressionTests
{
    [Fact]
    public void XlsxToPdfBytes_LandscapeSheet_ProducesALandscapePage()
    {
        var (width, height) = PdfProbe.PageSize(Converters.XlsxToPdfBytes(BuildSheet(landscape: true)));
        Assert.True(width > height,
            $"expected a landscape page, got {width}x{height}");
    }

    [Fact]
    public void XlsxToPdfBytes_PortraitSheet_ProducesAPortraitPage()
    {
        var (width, height) = PdfProbe.PageSize(Converters.XlsxToPdfBytes(BuildSheet(landscape: false)));
        Assert.True(height > width);
    }

    [Fact]
    public void XlsxToPdfBytes_HiddenSheet_IsNotRendered()
    {
        Assert.Equal(1, PdfProbe.PageCount(Converters.XlsxToPdfBytes(BuildTwoSheetWorkbook(secondHidden: true))));
    }

    [Fact]
    public void XlsxToPdfBytes_VisibleSheets_AreBothRendered()
    {
        Assert.Equal(2, PdfProbe.PageCount(Converters.XlsxToPdfBytes(BuildTwoSheetWorkbook(secondHidden: false))));
    }

    [Theory]
    [InlineData("4.099999999998545", "4.1")]
    [InlineData("21436.400000000001", "21436.4")]
    [InlineData("1456", "1456")]
    [InlineData("0.5", "0.5")]
    public void XlsxToPdfBytes_UnformattedNumbers_UseExcelsGeneralPrecision(string stored, string expected)
    {
        var texts = ExtractTexts(Converters.XlsxToPdfBytes(BuildNumericSheet(stored)));

        Assert.Contains(expected, texts);
    }

    static byte[] BuildNumericSheet(string storedValue)
    {
        using var ms = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
            worksheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell { CellReference = "A1", CellValue = new CellValue(storedValue) })
                { RowIndex = 1 }));
        }
        return ms.ToArray();
    }

    static List<string> ExtractTexts(byte[] pdfBytes) => PdfProbe.Words(pdfBytes);

    static byte[] BuildSheet(bool landscape)
    {
        using var ms = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
            worksheetPart.Worksheet = new Worksheet(
                new SheetData(new Row(new Cell
                {
                    CellReference = "A1",
                    DataType = CellValues.String,
                    CellValue = new CellValue("value"),
                })
                { RowIndex = 1 }),
                new PageSetup
                {
                    PaperSize = 1,
                    Orientation = landscape ? OrientationValues.Landscape : OrientationValues.Portrait,
                });
        }
        return ms.ToArray();
    }

    static byte[] BuildTwoSheetWorkbook(bool secondHidden)
    {
        using var ms = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            var second = new Sheet { Id = "rId2", SheetId = 2, Name = "Sheet2" };
            if (secondHidden) second.State = SheetStateValues.Hidden;

            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" },
                second));

            foreach (var id in new[] { "rId1", "rId2" })
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>(id);
                worksheetPart.Worksheet = new Worksheet(new SheetData(new Row(new Cell
                {
                    CellReference = "A1",
                    DataType = CellValues.String,
                    CellValue = new CellValue(id),
                })
                { RowIndex = 1 }));
            }
        }
        return ms.ToArray();
    }
}
