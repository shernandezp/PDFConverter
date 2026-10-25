# PDFConverter - Developer Documentation

This document explains the project structure, architecture, and how to contribute.

## Project Structure

```
src/
+-- PDFConverter/          # Main library
+-- PDFConverter.Tests/    # xUnit test project
+-- PdfInspector/          # Diagnostic tool: dumps page structure, images and text of a PDF
+-- TestConsole/           # Console app for manual conversions
```

Inside `src/PDFConverter/`, files group as follows:

| Group | Files | Role |
|-------|-------|------|
| Public API | `Converters`, `DocxConverter`, `XlsxConverter`, `OpenXmlHelpers` | Entry points, font registration, diagnostic hooks |
| Word rendering | `WordDocumentBuilder`, `WordContentRenderer`, `WordTableRenderer`, `VmlTextBoxRenderer`, `WordRenderContext` | DOCX -> layout model: page setup, runs, tables, VML shapes, per-conversion state |
| Word parsing | `WordHelpers`, `WordStyleCache`, `WordNumbering`, `WordImage` | Styles and formatting, cached style/theme/numbering lookups, list labels, picture and line extraction |
| Excel | `ExcelTableRenderer`, `ExcelStyles`, `ExcelNumberFormat`, `ExcelHelpers` | Sheet grid and merges, cached cell formats, format codes, images and column widths |
| Shared | `Units`, `LayoutDefaults`, `ColorUtils`, `ParagraphFormat`, `RunFormat`, `BorderInfo` | Conversions and invariant parsing, fallback measurements, colours, formatting records |
| Layout (`Layout/`) | `Model`, `TextLayout`, `LayoutEngine`, `DisplayList`, `SkiaPdfWriter`, `FontRegistry`, `FontSubsetter`, `ImageSource`, `TiffDecoder`, `MetafileImages` | Backend-neutral document model, shaping, line breaking and bidi reordering, pagination, tables and wrap-around, draw ops, Skia PDF painting, font resolution and fallback, glyph stripping, image decoding |
| Output | `ConverterExtensions`, `FontUtils`, `Fonts/` | Image format detection and GDI+ re-encoding, font name parsing, embedded Noto Emoji |

## Architecture

```
Input (file/stream/bytes)
  -> OpenXML SDK parses the document
  -> WordHelpers / ExcelHelpers extract formatting
  -> Style resolution chain (docDefaults -> paragraph style -> character style -> inline)
  -> Layout.Document model is built (sections, paragraphs, tables, floats, headers/footers)
  -> LayoutEngine breaks text into lines, paginates and produces per-page draw ops
  -> SkiaPdfWriter strips unused glyphs from the fonts and paints the pages into a Skia PDF
  -> Output (file or byte[])
```

### Key Design Decisions

- **Facade pattern**: `Converters` is the only class most consumers need. `DocxConverter` and `XlsxConverter` are also public for advanced usage.
- **Builder pattern**: `WordDocumentBuilder.Build()` and `XlsxConverter.BuildDocument()` return a `Layout.Document`; `LayoutEngine` and `SkiaPdfWriter` turn it into a PDF, so file and stream output share one rendering path.
- **Style resolution**: formatting cascades from docDefaults through the paragraph style chain and character style to inline properties. Toggle properties (bold, italic, caps) are tri-state, so an explicit `w:b val="0"` turns off a style that enables bold. `WordStyleCache` and `ExcelStyles` cache these lookups per document.
- **Units**: every OpenXML measurement goes through `Units`, which parses with the invariant culture — OpenXML always writes `.` as the decimal separator regardless of the machine locale.
- **No logging by default**: `OpenXmlHelpers.ImageLoadLogger` and `FontLoadLogger` are null unless the consumer sets them.
- **Own layout engine**: Skia only draws, so `Layout/` owns line breaking, justification, tab stops, per-glyph font fallback, table measuring, repeated header rows and pagination. Page backgrounds, hyperlinks on images and floating shapes are ordinary draw ops rather than post-render passes.
- **Font subsetting**: Skia's PDF backend embeds whole font files, so `FontSubsetter` empties the unused glyphs (keeping glyph ids and the cmap intact) before the typeface is handed to Skia.
- **Shaped glyph runs**: every run is shaped with HarfBuzz once at layout time; the resulting glyph ids and positions are what gets measured, painted (as a positioned text run with clusters, so the PDF keeps a correct ToUnicode map) and subset.
- **Wrap-around**: a wrapped floating picture registers an exclusion rectangle on its page; paragraphs ask for the free segment line by line and move below the picture when less than half an inch remains.
- **Excel pictures and overflow**: pictures are anchored to their row and may extend across cells and rows; unwrapped text runs into neighbouring empty cells and is clipped there, as Excel does.

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| DocumentFormat.OpenXml | 3.5.1 | Parse DOCX/XLSX OpenXML packages |
| SkiaSharp | 4.153.1 | Text measuring, image decoding and PDF output |
| SkiaSharp.HarfBuzz | 4.153.1 | Text shaping (kerning, ligatures, complex scripts) |
| SkiaSharp.NativeAssets.Linux.NoDependencies | 4.153.1 | Linux native Skia build without system library dependencies |
| System.Drawing.Common | 10.0.13 | Last-resort rasterising of vector-only metafiles (Windows-only, skipped elsewhere) |

The test project and `PdfInspector` read PDFs back with [PdfPig](https://github.com/UglyToad/PdfPig) (Apache-2.0).

## Building & Testing

Requirements: .NET 10 SDK.

```bash
dotnet build
dotnet test
dotnet pack -c Release
```

`global.json` opts `dotnet test` into Microsoft.Testing.Platform, which xUnit v3 requires on the .NET 10 SDK; the test project can also be run directly with `dotnet run --project src/PDFConverter.Tests`.

The test project contains 248 tests: unit tests over individual methods, and integration tests that convert in-memory OpenXML documents and assert on the rendered PDF (page counts, text, images, borders). All test documents are constructed programmatically — no external or confidential files are needed.

## Known OpenXML and Skia Gotchas

- **Duplicate types that look interchangeable**: `StyleRunProperties`/`StyleParagraphProperties`, `RunPropertiesBaseStyle`/`ParagraphPropertiesBaseStyle` (what `w:docDefaults` parses as) and `ParagraphMarkRunProperties` all carry the same children as `RunProperties`/`ParagraphProperties` but are distinct classes. `GetFirstChild<RunProperties>()` silently returns null and a cast-based clone always fails — copy the children instead (`WordStyleCache.CopyProperties`).
- **`EnumValue<T>.Value.ToString()` returns garbage** in OpenXml 3.x (e.g. `"LineSpacingRuleValues { }"`). Compare with `==`, or read `.InnerText` when the raw token is needed.
- **`w:hyperlink` inside DrawingML** parses as `OpenXmlUnknownElement`, not `W.Hyperlink`. Match by `LocalName`/`NamespaceUri`.
- **Skia PDF does not subset fonts**: every typeface drawn is embedded in full unless `FontSubsetter` strips it first; CFF-based OpenType fonts are embedded whole.
- **Skia honours PNG transparency exactly**: a palette PNG whose border colour is flagged transparent renders without that border, where GDI+ used to paint it.
- **`SKDocumentPdfMetadata` must start from `Default`**: an object initializer leaves `EncodingQuality` at 0, which re-encodes every picture as JPEG at quality 0.
- **`SKTypeface.ContainsGlyph` is obsolete**: query glyph coverage through an `SKFont`.

## Contributing

1. Fork the repository and create a feature branch
2. Add unit tests for any new behavior
3. Ensure all existing tests still pass
4. Submit a PR with a clear description of the changes

## Publishing

1. **CI Build** (`ci-build.yml`): runs on every push to `main` and on PRs. Builds and runs tests.
2. **NuGet Publish** (`nuget-publish.yml`): triggered by version tags (e.g. `v1.0.0`). Packs and publishes to nuget.org.

To release a new version:

```bash
# Update version in PDFConverter.csproj
# Update CHANGELOG.md
git tag v1.0.1
git push --tags
```

The `NUGET_API_KEY` secret must be configured in the repository settings.
