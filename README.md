# PDFConverter

[![CI Build](https://github.com/shernandezp/PDFConverter/actions/workflows/ci-build.yml/badge.svg)](https://github.com/shernandezp/PDFConverter/actions/workflows/ci-build.yml)
[![NuGet](https://img.shields.io/nuget/v/DocxXlsx.PDFConverter.svg)](https://www.nuget.org/packages/DocxXlsx.PDFConverter/)

A lightweight, free and open-source .NET library that converts **DOCX** and **XLSX** files to **PDF** using [OpenXML SDK](https://github.com/dotnet/Open-XML-SDK) and [SkiaSharp](https://github.com/mono/SkiaSharp). No Microsoft Office installation required.

## Features

- **DOCX → PDF** — paragraphs, tables, images, headers/footers, styles, hyperlinks, VML textboxes, emoji
- **XLSX → PDF** — cell formatting, merged cells, column widths, images (including floating images that overflow rows), connector shape rendering (signature lines)
- Multiple input formats: file path, `Stream`, or `byte[]`
- Output to file or `byte[]` for in-memory workflows
- Embedded emoji font (Noto Emoji) for cross-platform rendering
- No Office interop — runs on servers and containers

## Installation

```bash
dotnet add package DocxXlsx.PDFConverter
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
| Connector shapes (signature lines) | ✅ | ✅ |
| Image cropping (srcRect) | ✅ | — |
| Headers and footers | ✅ | — |
| Hyperlinks | ✅ | — |
| VML textboxes | ✅ | — |
| Emoji | ✅ | — |
| Tab stops | ✅ | — |
| Merged cells | ✅ | ✅ |
| Lists and numbering (bullets, multi-level) | ✅ | — |
| Page number fields (PAGE, NUMPAGES) | ✅ | — |
| Landscape orientation | ✅ | ✅ |
| Section breaks with their own page setup, headers and footers | ✅ | — |
| Widow/orphan control, keep with next, keep lines together | ✅ | — |
| Character spacing, width scaling, strike-through, highlight, hidden text | ✅ | — |
| Paragraph borders | ✅ | — |
| Dashed and dotted borders | ✅ | ✅ |
| PDF metadata from document properties | ✅ | ✅ |
| Text flowing beside pictures (square, tight, through) | ✅ | — |
| Complex scripts, ligatures and kerning (HarfBuzz) | ✅ | ✅ |
| Right-to-left paragraphs and mixed-direction lines | ✅ | — |
| TIFF, EMF and WMF pictures on every platform | ✅ | ✅ |

## Known Limitations

- **Font substitution**: If a document uses fonts not installed on the system, the closest available family is used, which may cause minor layout differences.
- **Character spacing**: condensed spacing is not applied to space characters, because converter-generated documents use it to encode line breaks; letters may therefore wrap slightly differently from Word.
- **Color emoji**: Only monochrome emoji glyphs (Noto Emoji) are rendered; color emoji (COLR/CPAL) is not supported.
- **Tight and through wrapping** use the picture's bounding box, like square wrapping; text never follows the picture's outline.
- **Metafiles (EMF, WMF)**: only the raster picture embedded in the metafile is rendered; vector-only metafiles need Windows, where GDI+ converts them.
- **Bidirectional text**: right-to-left runs are reordered per line with a simplified algorithm; nested embeddings and bracket mirroring are not handled.
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
