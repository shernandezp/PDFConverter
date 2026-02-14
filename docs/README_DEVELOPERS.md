# PDFConverter - Developer Documentation

This document explains the project structure, architecture, and how to contribute.

## Project Structure

```
PDFConverter/
+-- src/
|   +-- PDFConverter/              # Main library
|   |   +-- Converters.cs          # Public API facade
|   |   +-- DocxConverter.cs       # DOCX to PDF pipeline (~1200 lines)
|   |   +-- XlsxConverter.cs       # XLSX to PDF pipeline
|   |   +-- WordHelpers.cs         # OpenXML parsing: styles, formatting, borders
|   |   +-- WordTableRenderer.cs   # Table rendering with conditional formatting
|   |   +-- ConverterExtensions.cs # Image extraction, emoji, Roman numerals
|   |   +-- OpenXmlHelpers.cs      # Font resolver, image byte retrieval
|   |   +-- ExcelHelpers.cs        # Excel cell value/format helpers
|   |   +-- ExcelTableRenderer.cs  # Excel table rendering
|   |   +-- ParagraphFormat.cs     # Paragraph formatting record
|   |   +-- RunFormat.cs           # Run formatting record with ApplyTo
|   |   +-- BorderInfo.cs          # Border info record
|   |   +-- FontUtils.cs           # Font utilities
|   |   +-- Fonts/                 # Embedded NotoEmoji-Regular.ttf
|   |   +-- TestDocuments/         # Sample DOCX/XLSX files for testing
|   |   +-- pdfconverter-findings.json  # Detailed findings and gotchas
|   +-- PDFConverter.Tests/        # xUnit test project (139 tests)
|   +-- TestConsole/               # Console app for manual testing
+-- docs/
|   +-- README_CLIENTS.md          # Client/consumer documentation
|   +-- README_DEVELOPERS.md       # This file
+-- .github/workflows/
|   +-- ci-build.yml               # CI: build + test on push/PR
|   +-- nuget-publish.yml          # Publish NuGet on version tag
+-- README.md                      # Main README
+-- CHANGELOG.md                   # Release notes
+-- LICENSE                        # MIT License
```

## Architecture

The conversion pipeline follows this flow:

```
Input (file/stream/bytes)
  -> OpenXML SDK parses the document
  -> WordHelpers / ExcelHelpers extract formatting
  -> Style resolution chain (inline -> style -> basedOn -> docDefaults)
  -> MigraDoc Document model is built
  -> PdfDocumentRenderer renders to PDF
  -> Output (file or byte[])
```

### Key Design Decisions

- **Facade pattern**: `Converters.cs` is the only class most consumers need. Internal converters are also public for advanced usage.
- **BuildRenderer pattern**: DocxConverter and XlsxConverter use a `BuildRenderer()` method that builds the MigraDoc document and returns a `PdfDocumentRenderer`. This allows both file and stream output without duplicating rendering logic.
- **Style resolution**: OpenXML styles form an inheritance chain (inline properties -> named style -> basedOn parent -> docDefaults). The code walks this chain for both paragraph and run properties.
- **No logging by default**: `OpenXmlHelpers.ImageLoadLogger` is null by default. Consumers can set it to receive diagnostic messages.
- **Floating image pattern**: Spacer-column images that share rows with text are rendered as absolutely positioned section-level images (`section.AddImage` with `WrapStyle.None`) to simulate Excel's image overflow behavior. The `rowYOffsets` array tracks cumulative row heights for positioning floating images.

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| DocumentFormat.OpenXml | 3.4.1 | Parse DOCX/XLSX OpenXML packages |
| PdfSharp-MigraDoc | 6.2.4 | Build and render PDF documents |
| System.Drawing.Common | 10.0.3 | Image cropping (Windows-only, graceful fallback) |

## Building

Requirements:
- .NET 10 SDK

```bash
# Build
dotnet build

# Run tests
dotnet test

# Pack NuGet package
dotnet pack -c Release
```

## Testing

The test project (`PDFConverter.Tests`) contains 139 tests:

- **Unit tests**: Test individual methods with in-memory OpenXML documents (WordHelpers, ConverterExtensions, RunFormat, ParagraphFormat, BorderInfo)
- **Integration tests**: Convert real test documents to PDF and validate page counts, file sizes, and PDF header bytes

```bash
dotnet test --verbosity normal
```

### Test Documents

Test documents are stored in `src/PDFConverter/TestDocuments/`. Integration tests locate them via relative path from the test assembly output directory.

## Known OpenXML Gotchas

These are critical pitfalls discovered during development. Read `pdfconverter-findings.json` for the complete list.

- **`StyleParagraphProperties` is NOT `ParagraphProperties`** — `CloneNode(true) as ParagraphProperties` returns null. Must manually build properties.
- **`EnumValue<T>.ToString()` returns garbage in OpenXml 3.x** — e.g., `LineSpacingRuleValues.Exact.ToString()` returns `"LineSpacingRuleValues { }"`. Always compare with `==`.
- **MigraDoc `Row.Cells.Count` returns 0** before rendering. Use the known column count from the source document.
- **MigraDoc border extension across merged cells** — MigraDoc computes `max(cell_above.bottom, cell_below.top)` across ALL columns of a merged cell, creating spurious full-width lines when borders above are inconsistent. No API-level workaround; must skip MergeRight and simulate centering.
- **`w:hyperlink` inside DrawingML** is parsed as `OpenXmlUnknownElement`, not `W.Hyperlink`. Match by `LocalName`/`NamespaceUri`.

## Contributing

1. Fork the repository and create a feature branch
2. Read `pdfconverter-findings.json` for context on past issues and architectural decisions
3. Add unit tests for any new behavior
4. Ensure all existing tests still pass
5. Submit a PR with a clear description of the changes

## Publishing

The CI/CD pipeline handles publishing automatically:

1. **CI Build** (`ci-build.yml`): Runs on every push to `main` and on PRs. Builds and runs tests.
2. **NuGet Publish** (`nuget-publish.yml`): Triggered by version tags (e.g., `v1.0.0`). Packs and publishes to nuget.org.

To release a new version:

```bash
# Update version in PDFConverter.csproj
# Update CHANGELOG.md
git tag v1.0.1
git push --tags
```

The `NUGET_API_KEY` secret must be configured in the repository settings.
