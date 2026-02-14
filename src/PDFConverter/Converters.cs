namespace PDFConverter;

/// <summary>
/// Simple facade for converting Office documents to PDF.
/// </summary>
public static class Converters
{
    /// <summary>
    /// Convert a DOCX file to PDF, writing to the specified output path.
    /// </summary>
    /// <param name="docxPath">Path to the source DOCX file.</param>
    /// <param name="pdfPath">Path where the PDF will be written.</param>
    public static void DocxToPdf(string docxPath, string pdfPath)
    {
        DocxConverter.DocxToPdf(docxPath, pdfPath);
    }

    /// <summary>
    /// Convert a DOCX stream to PDF, writing to the specified output path.
    /// The input stream is not closed by this method.
    /// </summary>
    public static void DocxToPdf(Stream docxStream, string pdfPath)
    {
        DocxConverter.DocxToPdf(docxStream, pdfPath);
    }

    /// <summary>
    /// Convert a DOCX byte array to PDF, writing to the specified output path.
    /// </summary>
    public static void DocxToPdf(byte[] docxBytes, string pdfPath)
    {
        DocxConverter.DocxToPdf(docxBytes, pdfPath);
    }

    /// <summary>
    /// Convert a DOCX byte array to PDF and return the PDF as a byte array.
    /// </summary>
    public static byte[] DocxToPdfBytes(byte[] docxBytes)
    {
        return DocxConverter.DocxToPdfBytes(docxBytes);
    }

    /// <summary>
    /// Convert a DOCX stream to PDF and return the PDF as a byte array.
    /// The input stream is not closed by this method.
    /// </summary>
    public static byte[] DocxToPdfBytes(Stream docxStream)
    {
        return DocxConverter.DocxToPdfBytes(docxStream);
    }

    /// <summary>
    /// Convert an XLSX file to PDF, writing to the specified output path.
    /// </summary>
    /// <param name="xlsxPath">Path to the source XLSX file.</param>
    /// <param name="pdfPath">Path where the PDF will be written.</param>
    public static void XlsxToPdf(string xlsxPath, string pdfPath)
    {
        XlsxConverter.XlsxToPdf(xlsxPath, pdfPath);
    }

    /// <summary>
    /// Convert an XLSX stream to PDF, writing to the specified output path.
    /// The input stream is not closed by this method.
    /// </summary>
    public static void XlsxToPdf(Stream xlsxStream, string pdfPath)
    {
        XlsxConverter.XlsxToPdf(xlsxStream, pdfPath);
    }

    /// <summary>
    /// Convert an XLSX byte array to PDF, writing to the specified output path.
    /// </summary>
    public static void XlsxToPdf(byte[] xlsxBytes, string pdfPath)
    {
        XlsxConverter.XlsxToPdf(xlsxBytes, pdfPath);
    }

    /// <summary>
    /// Convert an XLSX byte array to PDF and return the PDF as a byte array.
    /// </summary>
    public static byte[] XlsxToPdfBytes(byte[] xlsxBytes)
    {
        return XlsxConverter.XlsxToPdfBytes(xlsxBytes);
    }

    /// <summary>
    /// Convert an XLSX stream to PDF and return the PDF as a byte array.
    /// The input stream is not closed by this method.
    /// </summary>
    public static byte[] XlsxToPdfBytes(Stream xlsxStream)
    {
        return XlsxConverter.XlsxToPdfBytes(xlsxStream);
    }
}
