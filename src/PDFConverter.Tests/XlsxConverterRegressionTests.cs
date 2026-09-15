using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PdfSharp.Pdf.IO;
using Xunit;

namespace PDFConverter.Tests;

public class XlsxConverterRegressionTests
{
    [Fact]
    public void XlsxToPdfBytes_LandscapeSheet_ProducesALandscapePage()
    {
        using var ms = new MemoryStream(Converters.XlsxToPdfBytes(BuildSheet(landscape: true)));
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);

        var page = pdf.Pages[0];
        Assert.True(page.Width.Point > page.Height.Point,
            $"expected a landscape page, got {page.Width.Point}x{page.Height.Point}");
    }

    [Fact]
    public void XlsxToPdfBytes_PortraitSheet_ProducesAPortraitPage()
    {
        using var ms = new MemoryStream(Converters.XlsxToPdfBytes(BuildSheet(landscape: false)));
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);

        var page = pdf.Pages[0];
        Assert.True(page.Height.Point > page.Width.Point);
    }

    [Fact]
    public void XlsxToPdfBytes_HiddenSheet_IsNotRendered()
    {
        using var ms = new MemoryStream(Converters.XlsxToPdfBytes(BuildTwoSheetWorkbook(secondHidden: true)));
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);

        Assert.Equal(1, pdf.PageCount);
    }

    [Fact]
    public void XlsxToPdfBytes_VisibleSheets_AreBothRendered()
    {
        using var ms = new MemoryStream(Converters.XlsxToPdfBytes(BuildTwoSheetWorkbook(secondHidden: false)));
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);

        Assert.Equal(2, pdf.PageCount);
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

    static List<string> ExtractTexts(byte[] pdfBytes)
    {
        var texts = new List<string>();
        using var ms = new MemoryStream(pdfBytes);
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        foreach (var page in pdf.Pages) Walk(PdfSharp.Pdf.Content.ContentReader.ReadContent(page), texts);
        return texts;
    }

    static void Walk(PdfSharp.Pdf.Content.Objects.CSequence sequence, List<string> texts)
    {
        foreach (var item in sequence)
        {
            if (item is PdfSharp.Pdf.Content.Objects.CSequence nested) { Walk(nested, texts); continue; }
            if (item is not PdfSharp.Pdf.Content.Objects.COperator op) continue;
            if (op.OpCode.Name is not ("Tj" or "TJ") || op.Operands.Count == 0) continue;

            var text = string.Empty;
            if (op.Operands[0] is PdfSharp.Pdf.Content.Objects.CString s) text = s.Value;
            else if (op.Operands[0] is PdfSharp.Pdf.Content.Objects.CArray array)
                foreach (var element in array)
                    if (element is PdfSharp.Pdf.Content.Objects.CString cs) text += cs.Value;

            if (!string.IsNullOrWhiteSpace(text)) texts.Add(text.Trim());
        }
    }

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
