using DocumentFormat.OpenXml.Packaging;

namespace PDFConverter;

/// <summary>Converter for DOCX files to PDF.</summary>
public static class DocxConverter
{
    /// <summary>Convert a DOCX file to PDF at the specified path.</summary>
    public static void DocxToPdf(string docxPath, string pdfPath)
    {
        using var word = WordprocessingDocument.Open(docxPath, false);
        DocxToPdfInternal(word, pdfPath);
    }

    /// <summary>Convert a DOCX stream to PDF at the specified path. The stream is left open.</summary>
    public static void DocxToPdf(Stream docxStream, string pdfPath)
    {
        using var buffer = Copy(docxStream);
        using var word = WordprocessingDocument.Open(buffer, false);
        DocxToPdfInternal(word, pdfPath);
    }

    /// <summary>Convert DOCX bytes to PDF at the specified path.</summary>
    public static void DocxToPdf(byte[] docxBytes, string pdfPath)
    {
        using var buffer = new MemoryStream(docxBytes);
        using var word = WordprocessingDocument.Open(buffer, false);
        DocxToPdfInternal(word, pdfPath);
    }

    /// <summary>Convert DOCX bytes to PDF and return the PDF bytes.</summary>
    public static byte[] DocxToPdfBytes(byte[] docxBytes)
    {
        using var buffer = new MemoryStream(docxBytes);
        using var word = WordprocessingDocument.Open(buffer, false);
        return RenderToBytes(word);
    }

    /// <summary>Convert a DOCX stream to PDF and return the PDF bytes. The stream is left open.</summary>
    public static byte[] DocxToPdfBytes(Stream docxStream)
    {
        using var buffer = Copy(docxStream);
        using var word = WordprocessingDocument.Open(buffer, false);
        return RenderToBytes(word);
    }

    internal static void DocxToPdfInternal(WordprocessingDocument word, string pdfPath)
    {
        using var images = new TempImageStore();
        WordDocumentBuilder.Build(word, images).Save(pdfPath);
    }

    internal static byte[] RenderToBytes(WordprocessingDocument word)
    {
        using var images = new TempImageStore();
        var renderer = WordDocumentBuilder.Build(word, images);
        using var output = new MemoryStream();
        renderer.Save(output, false);
        return output.ToArray();
    }

    static MemoryStream Copy(Stream source)
    {
        var buffer = new MemoryStream();
        source.CopyTo(buffer);
        buffer.Position = 0;
        return buffer;
    }
}
