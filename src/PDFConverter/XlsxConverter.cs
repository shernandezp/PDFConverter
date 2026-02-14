using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;

namespace PDFConverter;

/// <summary>
/// Converter for XLSX files to PDF.
/// </summary>
public static class XlsxConverter
{
    /// <summary>
    /// Convert an XLSX file to PDF at the specified path.
    /// </summary>
    public static void XlsxToPdf(string xlsxPath, string pdfPath)
    {
        using var spreadsheet = SpreadsheetDocument.Open(xlsxPath, false);
        XlsxToPdfInternal(spreadsheet, pdfPath);
    }

    /// <summary>
    /// Convert an XLSX stream to PDF at the specified path.
    /// The input stream will not be closed by this method.
    /// </summary>
    public static void XlsxToPdf(Stream xlsxStream, string pdfPath)
    {
        using var ms = new MemoryStream();
        xlsxStream.CopyTo(ms);
        ms.Position = 0;
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        XlsxToPdfInternal(spreadsheet, pdfPath);
    }

    /// <summary>
    /// Convert an XLSX byte[] to PDF at the specified path.
    /// </summary>
    public static void XlsxToPdf(byte[] xlsxBytes, string pdfPath)
    {
        using var ms = new MemoryStream(xlsxBytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        XlsxToPdfInternal(spreadsheet, pdfPath);
    }

    /// <summary>
    /// Convert an XLSX byte[] to PDF and return the result as a byte array.
    /// </summary>
    public static byte[] XlsxToPdfBytes(byte[] xlsxBytes)
    {
        using var ms = new MemoryStream(xlsxBytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        return XlsxToPdfToStream(spreadsheet).ToArray();
    }

    /// <summary>
    /// Convert an XLSX stream to PDF and return the result as a byte array.
    /// The input stream is not closed by this method.
    /// </summary>
    public static byte[] XlsxToPdfBytes(Stream xlsxStream)
    {
        using var ms = new MemoryStream();
        xlsxStream.CopyTo(ms);
        ms.Position = 0;
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        return XlsxToPdfToStream(spreadsheet).ToArray();
    }

    static void XlsxToPdfInternal(SpreadsheetDocument spreadsheet, string pdfPath)
    {
        var renderer = BuildRenderer(spreadsheet, out var tempFiles);
        try
        {
            renderer.Save(pdfPath);
        }
        finally
        {
            foreach (var tf in tempFiles)
                ConverterExtensions.TryDeleteTempFile(tf);
        }
    }

    static MemoryStream XlsxToPdfToStream(SpreadsheetDocument spreadsheet)
    {
        var renderer = BuildRenderer(spreadsheet, out var tempFiles);
        try
        {
            var output = new MemoryStream();
            renderer.Save(output, false);
            output.Position = 0;
            return output;
        }
        finally
        {
            foreach (var tf in tempFiles)
                ConverterExtensions.TryDeleteTempFile(tf);
        }
    }

    static PdfDocumentRenderer BuildRenderer(SpreadsheetDocument spreadsheet, out List<string> tempFiles)
    {
        // Ensure fonts are available (system first, explicit mappings as fallback)
        OpenXmlHelpers.EnsureFontResolverInitialized();

        var wbPart = spreadsheet.WorkbookPart;
        var sheets = wbPart.Workbook.Sheets.Elements<S.Sheet>().ToList();
        if (sheets.Count == 0) throw new InvalidOperationException("No sheets found");

        var doc = new Document();
        tempFiles = new List<string>();

        foreach (var sheet in sheets)
        {
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!);
            var ws = wsPart.Worksheet;
            var sheetData = ws.Elements<S.SheetData>().FirstOrDefault();
            var section = doc.AddSection();

            // Read page margins from worksheet (Excel stores in inches)
            var pgMargins = ws.Elements<S.PageMargins>().FirstOrDefault();
            if (pgMargins != null)
            {
                double left = pgMargins.Left?.Value ?? 0.7;
                double right = pgMargins.Right?.Value ?? 0.7;
                double top = pgMargins.Top?.Value ?? 0.75;
                double bottom = pgMargins.Bottom?.Value ?? 0.75;
                // Use sensible minimums so content doesn't clip at page edge
                section.PageSetup.LeftMargin = Unit.FromInch(Math.Max(left, 0.2));
                section.PageSetup.RightMargin = Unit.FromInch(Math.Max(right, 0.2));
                section.PageSetup.TopMargin = Unit.FromInch(Math.Max(top, 0.2));
                section.PageSetup.BottomMargin = Unit.FromInch(Math.Max(bottom, 0.2));
            }
            else
            {
                section.PageSetup.LeftMargin = Unit.FromCentimeter(1.5);
                section.PageSetup.RightMargin = Unit.FromCentimeter(1.5);
                section.PageSetup.TopMargin = Unit.FromCentimeter(1.5);
                section.PageSetup.BottomMargin = Unit.FromCentimeter(1.5);
            }

            // Read page orientation from worksheet
            var pgSetup = ws.Elements<S.PageSetup>().FirstOrDefault();
            if (pgSetup?.Orientation?.Value == S.OrientationValues.Landscape)
            {
                section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;
            }

            // Set page size explicitly (MigraDoc defaults to 0 unless set)
            // Read paper size from Excel (1=Letter, 9=A4) or default to Letter
            var paperSize = pgSetup?.PaperSize?.Value ?? 1;
            if (paperSize == 9) // A4
            {
                section.PageSetup.PageWidth = Unit.FromPoint(595.276);
                section.PageSetup.PageHeight = Unit.FromPoint(841.89);
            }
            else // Letter (default)
            {
                section.PageSetup.PageWidth = Unit.FromInch(8.5);
                section.PageSetup.PageHeight = Unit.FromInch(11);
            }

            if (sheetData == null) continue;
            var rows = sheetData.Elements<S.Row>().ToList();
            if (rows.Count == 0) continue;

            // Images are now placed in table cells by RenderTable
            ExcelTableRenderer.RenderTable(section, wsPart, tempFiles);
        }

        var renderer = new PdfDocumentRenderer()
        {
            Document = doc
        };
        renderer.RenderDocument();

        return renderer;
    }

}
