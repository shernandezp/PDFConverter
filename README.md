# PDFConverter

[![CI Build](https://github.com/shernandezp/PDFConverter/actions/workflows/ci-build.yml/badge.svg)](https://github.com/shernandezp/PDFConverter/actions/workflows/ci-build.yml)
[![NuGet](https://img.shields.io/nuget/v/PDFConverter.svg)](https://www.nuget.org/packages/PDFConverter/)

A lightweight, free and open-source .NET library that converts **DOCX** and **XLSX** files to **PDF** using [OpenXML SDK](https://github.com/dotnet/Open-XML-SDK) and [PdfSharp/MigraDoc](https://github.com/empira/PDFsharp). No Microsoft Office installation required.

## Why This Library?

I don't currently have the budget to pay for a licensed library to convert DOCX and XLSX files to PDF, even though there are many very good options on the market. So, I decided to build my own converter. It may or may not fit your needs. It's not the fastest option if you're looking for high performance, but it's good enough for me, especially considering that it's free. For simple files with simple objects, it should be enough.

This project was built with significant help from [GitHub Copilot](https://github.com/features/copilot) — from diagnosing tricky OpenXML quirks to fixing MigraDoc rendering issues, Copilot was an invaluable coding partner throughout the entire development process.

## Features

- **DOCX → PDF** — paragraphs, tables, images, headers/footers, styles, hyperlinks, VML textboxes, emoji
- **XLSX → PDF** — cell formatting, merged cells, column widths, images (including floating images that overflow rows), connector shape rendering (signature lines)
- Multiple input formats: file path, `Stream`, or `byte[]`
- Output to file or `byte[]` for in-memory workflows
- Embedded emoji font (Noto Emoji) for cross-platform rendering
- No Office interop — runs on servers and containers

## Installation

```bash
dotnet add package PDFConverter
```

## Quick Start

```csharp
using PDFConverter;

// DOCX to PDF — from file
Converters.DocxToPdf("report.docx", "report.pdf");

// DOCX to PDF — from stream
using var stream = File.OpenRead("report.docx");
Converters.DocxToPdf(stream, "report.pdf");

// DOCX to PDF — get bytes (no file I/O)
byte[] docxBytes = File.ReadAllBytes("report.docx");
byte[] pdfBytes = Converters.DocxToPdfBytes(docxBytes);

// XLSX to PDF
Converters.XlsxToPdf("data.xlsx", "data.pdf");

// XLSX to PDF — get bytes
byte[] xlsxBytes = File.ReadAllBytes("data.xlsx");
byte[] pdfBytes2 = Converters.XlsxToPdfBytes(xlsxBytes);
```

For more usage examples, see the [Client Documentation](docs/README_CLIENTS.md).

## Font Registration

To ensure fonts used in documents are available for PDF rendering:

```csharp
// Register all fonts from a directory
OpenXmlHelpers.RegisterFontsFromDirectory("/path/to/fonts");

// Register specific font mappings
OpenXmlHelpers.RegisterFontMappings(new Dictionary<string, string>
{
    { "Calibri", "/path/to/Calibri.ttf" },
    { "Arial MT", "/path/to/ArialMT.ttf" }
});
```

## Supported Features

| Feature | DOCX | XLSX |
|---------|------|------|
| Text formatting (bold, italic, underline, color, size) | ✅ | ✅ |
| Paragraph alignment and indentation | ✅ | — |
| Tables with borders and shading | ✅ | ✅ |
| Conditional table formatting (header row, banded rows) | ✅ | — |
| Inline and floating images | ✅ | ✅ |
| Connector shapes (signature lines) | — | ✅ |
| Image cropping (srcRect) | ✅ | — |
| Headers and footers | ✅ | — |
| Hyperlinks | ✅ | — |
| VML textboxes | ✅ | — |
| Emoji | ✅ | — |
| Tab stops | ✅ | — |
| Merged cells | ✅ | ✅ |
| Landscape orientation | ✅ | — |

## Known Limitations

- **Font substitution**: If a document uses fonts not installed on the system, PdfSharp falls back to a default font which may cause minor layout differences.
- **EMF/WMF images**: Vector image formats are not supported by PdfSharp and will be skipped.
- **Color emoji**: Only monochrome emoji glyphs (Noto Emoji) are rendered; color emoji (COLR/CPAL) is not supported.
- **Image cropping**: `srcRect` cropping uses `System.Drawing.Bitmap` which is Windows-only. On other platforms, the uncropped image is used as a graceful fallback.
- **Character width scaling**: The `w:w` attribute is not supported by MigraDoc.
- **Complex layouts**: Very complex Office features (advanced VML, SmartArt, charts) may not be reproduced exactly.

## Building & Contributing

See the [Developer Documentation](docs/README_DEVELOPERS.md) for project structure, build instructions, and contribution guidelines.

```bash
dotnet build
dotnet test
```

## License

[MIT](LICENSE)

The embedded [Noto Emoji](https://github.com/googlefonts/noto-emoji) font is licensed under the [SIL Open Font License 1.1](https://opensource.org/licenses/OFL-1.1).
