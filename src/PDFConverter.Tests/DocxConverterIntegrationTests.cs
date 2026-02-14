using Xunit;

namespace PDFConverter.Tests;

/// <summary>
/// Integration tests that convert actual test DOCX/XLSX documents to PDF
/// and validate the output (no exceptions, file produced, correct page count).
/// These tests use the real test documents from the TestDocuments folder.
/// </summary>
public class DocxConverterIntegrationTests : IDisposable
{
    private readonly string _testDocsDir;
    private readonly List<string> _outputFiles = new();

    public DocxConverterIntegrationTests()
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

    private static int GetPdfPageCount(string pdfPath)
    {
        // PdfSharp can read page count
        using var pdf = PdfSharp.Pdf.IO.PdfReader.Open(pdfPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.InformationOnly);
        return pdf.PageCount;
    }

    private void SkipIfMissing(string fileName)
    {
        var path = Path.Combine(_testDocsDir, fileName);
        if (!File.Exists(path))
            Assert.Fail($"Test document '{fileName}' not found at {path}. Copy test documents to the TestDocuments folder.");
    }

    #region DOCX Conversion — File Path Overload

    [Fact]
    public void DocxToPdf_CertTemp_ProducesSinglePagePdf()
    {
        const string fileName = "Cert_Temp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("CertTemp");

        DocxConverter.DocxToPdf(input, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
        Assert.Equal(1, GetPdfPageCount(output));
    }

    [Fact]
    public void DocxToPdf_CheckListTemp_ProducesTwoPagePdf()
    {
        const string fileName = "CheckList_Temp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("CheckListTemp");

        DocxConverter.DocxToPdf(input, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
        Assert.Equal(2, GetPdfPageCount(output));
    }

    [Fact]
    public void DocxToPdf_MaqCertTemp_ProducesSinglePagePdf()
    {
        const string fileName = "Maq_Cert_Temp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("MaqCertTemp");

        DocxConverter.DocxToPdf(input, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
        Assert.Equal(1, GetPdfPageCount(output));
    }

    [Fact]
    public void DocxToPdf_MessageTemplateTemp_ProducesSinglePagePdf()
    {
        const string fileName = "MessageTemplate_Temp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("MessageTemplateTemp");

        DocxConverter.DocxToPdf(input, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
        Assert.Equal(1, GetPdfPageCount(output));
    }

    [Fact]
    public void DocxToPdf_TripControlTemp_ProducesSinglePagePdf()
    {
        const string fileName = "tripcontroltemp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("TripControlTemp");

        DocxConverter.DocxToPdf(input, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
        Assert.Equal(1, GetPdfPageCount(output));
    }

    [Fact]
    public void DocxToPdf_TripControlMovexx_ProducesSinglePagePdf()
    {
        const string fileName = "TripControl_Movexx.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("TripControlMovexx");

        DocxConverter.DocxToPdf(input, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
        Assert.Equal(1, GetPdfPageCount(output));
    }

    #endregion

    #region DOCX Conversion — Stream Overload

    [Fact]
    public void DocxToPdf_FromStream_ProducesValidPdf()
    {
        const string fileName = "Cert_Temp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("CertTemp_Stream");

        using var stream = File.OpenRead(input);
        DocxConverter.DocxToPdf(stream, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
        Assert.Equal(1, GetPdfPageCount(output));
    }

    #endregion

    #region DOCX Conversion — Byte Array Overload

    [Fact]
    public void DocxToPdf_FromByteArray_ProducesValidPdf()
    {
        const string fileName = "Cert_Temp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var output = GetOutputPath("CertTemp_Bytes");

        var bytes = File.ReadAllBytes(input);
        DocxConverter.DocxToPdf(bytes, output);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
        Assert.Equal(1, GetPdfPageCount(output));
    }

    #endregion

    #region DOCX Conversion — All Overloads Produce Same Page Count

    [Fact]
    public void DocxToPdf_AllOverloads_ProduceSamePageCount()
    {
        const string fileName = "CheckList_Temp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);

        var outputFile = GetOutputPath("CheckList_File");
        DocxConverter.DocxToPdf(input, outputFile);

        var outputStream = GetOutputPath("CheckList_Stream");
        using (var stream = File.OpenRead(input))
            DocxConverter.DocxToPdf(stream, outputStream);

        var outputBytes = GetOutputPath("CheckList_Bytes");
        DocxConverter.DocxToPdf(File.ReadAllBytes(input), outputBytes);

        var countFile = GetPdfPageCount(outputFile);
        var countStream = GetPdfPageCount(outputStream);
        var countBytes = GetPdfPageCount(outputBytes);

        Assert.Equal(countFile, countStream);
        Assert.Equal(countFile, countBytes);
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

    #region Byte Array Return Overloads

    [Fact]
    public void DocxToPdfBytes_FromByteArray_ReturnsValidPdf()
    {
        const string fileName = "Cert_Temp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var pdfBytes = Converters.DocxToPdfBytes(File.ReadAllBytes(input));

        Assert.NotNull(pdfBytes);
        Assert.True(pdfBytes.Length > 0);
        // PDF files start with %PDF
        Assert.Equal((byte)'%', pdfBytes[0]);
        Assert.Equal((byte)'P', pdfBytes[1]);
        Assert.Equal((byte)'D', pdfBytes[2]);
        Assert.Equal((byte)'F', pdfBytes[3]);
    }

    [Fact]
    public void DocxToPdfBytes_FromStream_ReturnsValidPdf()
    {
        const string fileName = "Cert_Temp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        using var stream = File.OpenRead(input);
        var pdfBytes = Converters.DocxToPdfBytes(stream);

        Assert.NotNull(pdfBytes);
        Assert.True(pdfBytes.Length > 0);
        Assert.Equal((byte)'%', pdfBytes[0]);
    }

    [Fact]
    public void DocxToPdfBytes_MatchesFileOutput_PageCount()
    {
        const string fileName = "Cert_Temp.docx";
        SkipIfMissing(fileName);

        var input = Path.Combine(_testDocsDir, fileName);
        var pdfBytes = Converters.DocxToPdfBytes(File.ReadAllBytes(input));

        // Write to temp file to verify page count matches file-based output
        var tempPath = GetOutputPath("CertTemp_BytesReturn");
        File.WriteAllBytes(tempPath, pdfBytes);
        Assert.Equal(1, GetPdfPageCount(tempPath));
    }

    #endregion

}
