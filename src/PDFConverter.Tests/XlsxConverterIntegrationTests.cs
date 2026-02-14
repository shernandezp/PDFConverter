using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;
using Xunit;

namespace PDFConverter.Tests;

/// <summary>
/// Integration tests for XLSX to PDF conversion.
/// </summary>
public class XlsxConverterIntegrationTests : IDisposable
{
    private readonly string _testDocsDir;
    private readonly List<string> _outputFiles = new();

    public XlsxConverterIntegrationTests()
    {
        _testDocsDir = Path.GetFullPath(
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "PDFConverter", "TestDocuments"));
    }

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

    private void SkipIfMissing(string fileName)
    {
        var path = Path.Combine(_testDocsDir, fileName);
        if (!File.Exists(path))
            Assert.Fail($"Test document '{fileName}' not found at {path}. Copy test documents to the TestDocuments folder.");
    }

    private record PdfImage(double Width, double Height, double X, double Y);
    private record PdfText(double X, double Y, string Content);

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

    private void WalkContent(CSequence seq, List<PdfImage> images, List<PdfText> texts, ref double curX, ref double curY)
    {
        for (int i = 0; i < seq.Count; i++)
        {
            if (seq[i] is CSequence sub)
            {
                WalkContent(sub, images, texts, ref curX, ref curY);
                continue;
            }
            if (seq[i] is COperator op)
            {
                if (op.OpCode.Name == "cm" && op.Operands.Count == 6)
                {
                    double a = GetOperandValue(op.Operands[0]);
                    double d = GetOperandValue(op.Operands[3]);
                    double tx = GetOperandValue(op.Operands[4]);
                    double ty = GetOperandValue(op.Operands[5]);
                    if (a > 10 || d > 10)
                        images.Add(new PdfImage(a, d, tx, ty));
                }
                else if (op.OpCode.Name == "BT")
                {
                    curX = 0; curY = 0;
                }
                else if (op.OpCode.Name == "Tm" && op.Operands.Count == 6)
                {
                    curX = GetOperandValue(op.Operands[4]);
                    curY = GetOperandValue(op.Operands[5]);
                }
                else if (op.OpCode.Name == "Td" && op.Operands.Count == 2)
                {
                    double tx = GetOperandValue(op.Operands[0]);
                    double ty = GetOperandValue(op.Operands[1]);
                    curX += tx;
                    curY += ty;
                }
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
    }

    private static double GetOperandValue(CObject operand) => operand switch
    {
        CReal r => r.Value,
        CInteger ci => ci.Value,
        _ => 0
    };

    private byte[] ConvertEmailTemplate()
    {
        const string fileName = "EmailTemplateSyE.xlsx";
        SkipIfMissing(fileName);
        var input = Path.Combine(_testDocsDir, fileName);
        return Converters.XlsxToPdfBytes(File.ReadAllBytes(input));
    }

    [Fact]
    public void XlsxToPdf_EmailTemplate_ProducesValidPdf()
    {
        const string fileName = "EmailTemplateSyE.xlsx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("EmailTemplate");

        XlsxConverter.XlsxToPdf(input, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
    }

    [Fact]
    public void XlsxToPdf_FromStream_ProducesValidPdf()
    {
        const string fileName = "EmailTemplateSyE.xlsx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("EmailTemplate_Stream");

        using var stream = File.OpenRead(input);
        XlsxConverter.XlsxToPdf(stream, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
    }

    [Fact]
    public void XlsxToPdf_FromByteArray_ProducesValidPdf()
    {
        const string fileName = "EmailTemplateSyE.xlsx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("EmailTemplate_Bytes");

        var bytes = File.ReadAllBytes(input);
        XlsxConverter.XlsxToPdf(bytes, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
    }

    [Fact]
    public void XlsxToPdf_NonExistentFile_ThrowsException()
    {
        var output = GetOutputPath("NonExistent");
        Assert.ThrowsAny<Exception>(() =>
            XlsxConverter.XlsxToPdf(@"C:\nonexistent_xlsx_12345.xlsx", output));
    }

    [Fact]
    public void XlsxToPdfBytes_FromByteArray_ReturnsValidPdf()
    {
        const string fileName = "EmailTemplateSyE.xlsx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var pdfBytes = Converters.XlsxToPdfBytes(File.ReadAllBytes(input));

        Assert.NotNull(pdfBytes);
        Assert.True(pdfBytes.Length > 0);
        Assert.Equal((byte)'%', pdfBytes[0]);
        Assert.Equal((byte)'P', pdfBytes[1]);
    }

    [Fact]
    public void XlsxToPdfBytes_FromStream_ReturnsValidPdf()
    {
        const string fileName = "EmailTemplateSyE.xlsx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        using var stream = File.OpenRead(input);
        var pdfBytes = Converters.XlsxToPdfBytes(stream);

        Assert.NotNull(pdfBytes);
        Assert.True(pdfBytes.Length > 0);
        Assert.Equal((byte)'%', pdfBytes[0]);
    }

    [Fact]
    public void XlsxToPdf_EmailTemplate_ProducesSinglePage()
    {
        var pdfBytes = ConvertEmailTemplate();
        using var ms = new MemoryStream(pdfBytes);
        var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Import);

        Assert.Equal(1, doc.PageCount);
    }

    [Fact]
    public void XlsxToPdf_EmailTemplate_ContainsFourImages()
    {
        var pdfBytes = ConvertEmailTemplate();
        var (images, _) = AnalyzePdf(pdfBytes);

        Assert.Equal(4, images.Count);
    }

    [Fact]
    public void XlsxToPdf_EmailTemplate_LogoImageHasCorrectDimensions()
    {
        var pdfBytes = ConvertEmailTemplate();
        var (images, _) = AnalyzePdf(pdfBytes);

        // Logo is the image with the highest top position (ty + height)
        var logo = images.OrderByDescending(img => img.Y + img.Height).First();

        Assert.InRange(logo.Width, 104.8 - 2, 104.8 + 2);
        Assert.InRange(logo.Height, 42.7 - 2, 42.7 + 2);
    }

    [Fact]
    public void XlsxToPdf_EmailTemplate_TitleAndSubtitleAreClose()
    {
        var pdfBytes = ConvertEmailTemplate();
        var (_, texts) = AnalyzePdf(pdfBytes);

        var certifText = texts.FirstOrDefault(t => t.Content.Contains("CERTIF"));
        var r543Text = texts.FirstOrDefault(t => t.Content.Contains("R543"));

        Assert.NotNull(certifText);
        Assert.NotNull(r543Text);

        double gap = Math.Abs(certifText.Y - r543Text.Y);
        Assert.True(gap < 20, $"Gap between CERTIFICADO and R543 is {gap:F1}pt, expected < 20pt");
    }

    [Fact]
    public void XlsxToPdf_EmailTemplate_SignatureImagesAreSideBySide()
    {
        var pdfBytes = ConvertEmailTemplate();
        var (images, _) = AnalyzePdf(pdfBytes);

        // Bottom two images sorted by Y ascending (lowest Y = bottom of page)
        var bottomTwo = images.OrderBy(img => img.Y).Take(2).OrderBy(img => img.X).ToList();
        Assert.Equal(2, bottomTwo.Count);

        var left = bottomTwo[0];
        var right = bottomTwo[1];

        // Same top position (within 20pt tolerance for row alignment)
        Assert.InRange(Math.Abs(left.Y - right.Y), 0, 20);

        // No overlap: right image X > left image X + width
        Assert.True(right.X > left.X + left.Width,
            $"Signature images overlap: right.X={right.X:F1}, left.X+W={left.X + left.Width:F1}");
    }

    [Fact]
    public void XlsxToPdf_EmailTemplate_HasConnectorUnderscoreLines()
    {
        var pdfBytes = ConvertEmailTemplate();
        var (_, texts) = AnalyzePdf(pdfBytes);

        var underscoreTexts = texts.Where(t => t.Content.Length > 1 && t.Content.All(c => c == '_')).ToList();

        Assert.True(underscoreTexts.Count >= 2,
            $"Expected at least 2 underscore lines, found {underscoreTexts.Count}");
    }

    [Fact]
    public void XlsxToPdf_EmailTemplate_UnderscoreLinesAreSeparated()
    {
        var pdfBytes = ConvertEmailTemplate();
        var (_, texts) = AnalyzePdf(pdfBytes);

        var underscoreTexts = texts.Where(t => t.Content.Length > 1 && t.Content.All(c => c == '_')).ToList();
        Assert.True(underscoreTexts.Count >= 2, "Need at least 2 underscore lines");

        var first = underscoreTexts[0];
        var second = underscoreTexts[1];

        Assert.True(Math.Abs(first.X - second.X) > 1,
            $"Underscore lines should be at different X positions: {first.X:F1} vs {second.X:F1}");
    }

    [Fact]
    public void XlsxToPdf_EmailTemplate_BannerImageDoesNotOverlapLogo()
    {
        var pdfBytes = ConvertEmailTemplate();
        var (images, _) = AnalyzePdf(pdfBytes);

        // Logo: highest top (Y + Height)
        var logo = images.OrderByDescending(img => img.Y + img.Height).First();

        // Banner: second highest top, wide image (~300pt)
        var banner = images
            .Where(img => img != logo)
            .OrderByDescending(img => img.Y + img.Height)
            .First();

        // Banner's top (Y + Height) should be <= logo's bottom (Y)
        Assert.True(banner.Y + banner.Height <= logo.Y,
            $"Banner top ({banner.Y + banner.Height:F1}) overlaps logo bottom ({logo.Y:F1})");
    }

    [Fact]
    public void XlsxToPdf_EmailTemplate_FooterIsCentered()
    {
        var pdfBytes = ConvertEmailTemplate();
        var (_, texts) = AnalyzePdf(pdfBytes);

        // Company website in the footer area
        var footer = texts.FirstOrDefault(t =>
            t.Content.Contains("www", StringComparison.OrdinalIgnoreCase));

        Assert.NotNull(footer);

        // US Letter width = 612pt, center = 306pt
        const double pageCenter = 306.0;
        Assert.InRange(footer.X, pageCenter - 150, pageCenter + 150);
    }

    // --- EmailTemplate_MAQ_DETAIL.xlsx tests ---
    // Removed — file-dependent tests replaced by in-memory tests in ExcelMergedCellTests.cs
}
