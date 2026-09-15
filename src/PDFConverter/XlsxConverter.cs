using DocumentFormat.OpenXml.Packaging;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace PDFConverter;

/// <summary>Converter for XLSX files to PDF.</summary>
public static class XlsxConverter
{
    /// <summary>Convert an XLSX file to PDF at the specified path.</summary>
    public static void XlsxToPdf(string xlsxPath, string pdfPath)
    {
        using var spreadsheet = SpreadsheetDocument.Open(xlsxPath, false);
        XlsxToPdfInternal(spreadsheet, pdfPath);
    }

    /// <summary>Convert an XLSX stream to PDF at the specified path. The stream is left open.</summary>
    public static void XlsxToPdf(Stream xlsxStream, string pdfPath)
    {
        using var buffer = Copy(xlsxStream);
        using var spreadsheet = SpreadsheetDocument.Open(buffer, false);
        XlsxToPdfInternal(spreadsheet, pdfPath);
    }

    /// <summary>Convert XLSX bytes to PDF at the specified path.</summary>
    public static void XlsxToPdf(byte[] xlsxBytes, string pdfPath)
    {
        using var buffer = new MemoryStream(xlsxBytes);
        using var spreadsheet = SpreadsheetDocument.Open(buffer, false);
        XlsxToPdfInternal(spreadsheet, pdfPath);
    }

    /// <summary>Convert XLSX bytes to PDF and return the PDF bytes.</summary>
    public static byte[] XlsxToPdfBytes(byte[] xlsxBytes)
    {
        using var buffer = new MemoryStream(xlsxBytes);
        using var spreadsheet = SpreadsheetDocument.Open(buffer, false);
        return RenderToBytes(spreadsheet);
    }

    /// <summary>Convert an XLSX stream to PDF and return the PDF bytes. The stream is left open.</summary>
    public static byte[] XlsxToPdfBytes(Stream xlsxStream)
    {
        using var buffer = Copy(xlsxStream);
        using var spreadsheet = SpreadsheetDocument.Open(buffer, false);
        return RenderToBytes(spreadsheet);
    }

    static void XlsxToPdfInternal(SpreadsheetDocument spreadsheet, string pdfPath)
    {
        using var images = new TempImageStore();
        BuildRenderer(spreadsheet, images).Save(pdfPath);
    }

    static byte[] RenderToBytes(SpreadsheetDocument spreadsheet)
    {
        using var images = new TempImageStore();
        var renderer = BuildRenderer(spreadsheet, images);
        using var output = new MemoryStream();
        renderer.Save(output, false);
        return output.ToArray();
    }

    static PdfDocumentRenderer BuildRenderer(SpreadsheetDocument spreadsheet, TempImageStore images)
    {
        OpenXmlHelpers.EnsureFontResolverInitialized();

        var workbookPart = spreadsheet.WorkbookPart
            ?? throw new InvalidOperationException("The workbook has no content.");
        var sheets = workbookPart.Workbook?.Sheets?.Elements<S.Sheet>().ToList() ?? [];
        if (sheets.Count == 0) throw new InvalidOperationException("No sheets found");

        var document = new Document();

        foreach (var sheet in sheets)
        {
            if (IsHidden(sheet)) continue;
            if (sheet.Id?.Value == null) continue;
            if (workbookPart.GetPartById(sheet.Id.Value) is not WorksheetPart worksheetPart) continue;

            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet?.Elements<S.SheetData>().FirstOrDefault();
            if (sheetData?.Elements<S.Row>().Any() != true) continue;

            var section = document.AddSection();
            ApplyPageSetup(section, worksheet!);
            ExcelTableRenderer.RenderSheet(section, worksheetPart, images);
        }

        if (document.Sections.Count == 0) document.AddSection();

        var renderer = new PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();
        return renderer;
    }

    static bool IsHidden(S.Sheet sheet) =>
        sheet.State?.Value == S.SheetStateValues.Hidden || sheet.State?.Value == S.SheetStateValues.VeryHidden;

    static void ApplyPageSetup(Section section, S.Worksheet worksheet)
    {
        var pageSetup = section.PageSetup;
        var margins = worksheet.Elements<S.PageMargins>().FirstOrDefault();

        if (margins != null)
        {
            pageSetup.LeftMargin = Margin(margins.Left?.Value);
            pageSetup.RightMargin = Margin(margins.Right?.Value);
            pageSetup.TopMargin = Margin(margins.Top?.Value);
            pageSetup.BottomMargin = Margin(margins.Bottom?.Value);
        }
        else
        {
            var fallback = Unit.FromCentimeter(LayoutDefaults.ExcelPageMarginCentimeters);
            pageSetup.LeftMargin = pageSetup.RightMargin = fallback;
            pageSetup.TopMargin = pageSetup.BottomMargin = fallback;
        }

        var printSetup = worksheet.Elements<S.PageSetup>().FirstOrDefault();
        var (width, height) = PaperSize(printSetup?.PaperSize?.Value);

        // MigraDoc ignores Orientation once the page size is set explicitly, so swap it here.
        if (printSetup?.Orientation?.Value == S.OrientationValues.Landscape)
            (width, height) = (height, width);

        pageSetup.PageWidth = Unit.FromPoint(width);
        pageSetup.PageHeight = Unit.FromPoint(height);
    }

    static Unit Margin(double? inches) =>
        Unit.FromInch(Math.Max(inches ?? LayoutDefaults.ExcelMinPageMarginInches,
            LayoutDefaults.ExcelMinPageMarginInches));

    static (double Width, double Height) PaperSize(uint? paperSize) => paperSize switch
    {
        5 or 6 => (612, 1008),
        8 => (841.89, 1190.55),
        9 or 70 => (595.276, 841.89),
        11 => (419.528, 595.276),
        13 => (515.906, 728.504),
        _ => (612, 792),
    };

    static MemoryStream Copy(Stream source)
    {
        var buffer = new MemoryStream();
        source.CopyTo(buffer);
        buffer.Position = 0;
        return buffer;
    }
}
