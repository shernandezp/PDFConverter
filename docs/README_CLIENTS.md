# PDFConverter - Client Documentation

Complete guide for integrating PDFConverter into your application.

## Installation

```bash
dotnet add package DocxXlsx.PDFConverter
```

Or via the NuGet Package Manager in Visual Studio, search for **DocxXlsx.PDFConverter**.

## API Reference

### DOCX to PDF

| Method | Description |
|--------|-------------|
| `Converters.DocxToPdf(string docxPath, string pdfPath)` | Convert from file path to file |
| `Converters.DocxToPdf(Stream docxStream, string pdfPath)` | Convert from stream to file |
| `Converters.DocxToPdf(byte[] docxBytes, string pdfPath)` | Convert from byte array to file |
| `Converters.DocxToPdfBytes(byte[] docxBytes)` | Convert from byte array, returns PDF bytes |
| `Converters.DocxToPdfBytes(Stream docxStream)` | Convert from stream, returns PDF bytes |

### XLSX to PDF

| Method | Description |
|--------|-------------|
| `Converters.XlsxToPdf(string xlsxPath, string pdfPath)` | Convert from file path to file |
| `Converters.XlsxToPdf(Stream xlsxStream, string pdfPath)` | Convert from stream to file |
| `Converters.XlsxToPdf(byte[] xlsxBytes, string pdfPath)` | Convert from byte array to file |
| `Converters.XlsxToPdfBytes(byte[] xlsxBytes)` | Convert from byte array, returns PDF bytes |
| `Converters.XlsxToPdfBytes(Stream xlsxStream)` | Convert from stream, returns PDF bytes |

## Usage Examples

### Basic file conversion

```csharp
using PDFConverter;

Converters.DocxToPdf("document.docx", "document.pdf");
Converters.XlsxToPdf("spreadsheet.xlsx", "spreadsheet.pdf");
```

### Stream-based conversion

Useful when receiving files via HTTP uploads or reading from cloud storage:

```csharp
using var fs = File.OpenRead("report.docx");
Converters.DocxToPdf(fs, "report.pdf");
```

### Byte array conversion (file to file)

```csharp
var bytes = File.ReadAllBytes("workbook.xlsx");
Converters.XlsxToPdf(bytes, "workbook.pdf");
```

### In-memory conversion (no file I/O)

Ideal for web APIs, cloud functions, or when you need to return the PDF directly:

```csharp
byte[] docxBytes = File.ReadAllBytes("report.docx");
byte[] pdfBytes = Converters.DocxToPdfBytes(docxBytes);

// Example: return as HTTP response in ASP.NET
return File(pdfBytes, "application/pdf", "report.pdf");
```

### Stream to bytes

```csharp
using var stream = File.OpenRead("report.docx");
byte[] pdfBytes = Converters.DocxToPdfBytes(stream);
```

## Font Registration

If your documents use specific fonts, register them at application startup to ensure correct rendering:

```csharp
// Register all fonts from a directory
OpenXmlHelpers.RegisterFontsFromDirectory("/path/to/fonts");

// Register specific font family mappings
OpenXmlHelpers.RegisterFontMappings(new Dictionary<string, string>
{
    { "Calibri", "/path/to/Calibri.ttf" },
    { "Arial MT", "/path/to/ArialMT.ttf" }
});
```

If a font is not found, PdfSharp will substitute a default font. The output will still be readable, but character widths may differ slightly from the original document.

## Logging

By default the library does not produce any console output. Two optional hooks report what the converter did with images and fonts:

```csharp
OpenXmlHelpers.ImageLoadLogger = message => Console.WriteLine(message);
OpenXmlHelpers.FontLoadLogger = message => Console.WriteLine(message);
```

## Important Notes

- **Stream handling**: Stream overloads copy the input into memory to allow seeking. The caller's stream is not closed.
- **Thread safety**: Each conversion call is independent. Multiple conversions can run concurrently.
- **Temp files**: The library creates temporary files for images during conversion and cleans them up automatically.
- **Windows-only image paths**: `srcRect` cropping and re-encoding of BMP/TIFF/EMF/WMF use `System.Drawing` (Windows-only). Elsewhere the original image is used as a graceful fallback.
- **XLSX floating images**: Images in spacer columns that overflow rows are rendered as absolutely positioned floating images to match Excel's layout.
- **Culture**: Conversion is culture-independent for parsing; number and date *output* follows the current culture, as Excel does.
- **Scope**: The library is designed for documents with common formatting and objects. Very complex Office features (SmartArt, charts, advanced VML) may not render perfectly.
