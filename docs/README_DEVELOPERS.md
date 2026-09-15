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
| Word rendering | `WordDocumentBuilder`, `WordContentRenderer`, `WordTableRenderer`, `VmlTextBoxRenderer`, `WordRenderContext` | DOCX -> MigraDoc: page setup, runs, tables, VML shapes, per-conversion state |
| Word parsing | `WordHelpers`, `WordStyleCache`, `WordNumbering`, `WordImage` | Styles and formatting, cached style/theme/numbering lookups, list labels, picture and line extraction |
| Excel | `ExcelTableRenderer`, `ExcelStyles`, `ExcelNumberFormat`, `ExcelHelpers` | Sheet grid and merges, cached cell formats, format codes, images and column widths |
| Shared | `Units`, `LayoutDefaults`, `ColorUtils`, `TextMeasure`, `ParagraphFormat`, `RunFormat`, `BorderInfo` | Conversions and invariant parsing, fallback measurements, colours, text width estimation, formatting records |
| Output | `TempImageStore`, `ConverterExtensions`, `PdfImageLinks`, `DirectoryFontResolver`, `FontUtils`, `Fonts/` | Image deduplication and encoding, emoji segmentation, post-render link annotations, font resolution, embedded Noto Emoji |

## Architecture

```
Input (file/stream/bytes)
  -> OpenXML SDK parses the document
  -> WordHelpers / ExcelHelpers extract formatting
  -> Style resolution chain (docDefaults -> paragraph style -> character style -> inline)
  -> MigraDoc Document model is built
  -> PdfDocumentRenderer renders to PDF
  -> Post-render passes (page backgrounds, image link annotations)
  -> Output (file or byte[])
```

### Key Design Decisions

- **Facade pattern**: `Converters` is the only class most consumers need. `DocxConverter` and `XlsxConverter` are also public for advanced usage.
- **Builder pattern**: `WordDocumentBuilder.Build()` and `XlsxConverter.BuildRenderer()` return a `PdfDocumentRenderer`, so file and stream output share one rendering path.
- **Style resolution**: formatting cascades from docDefaults through the paragraph style chain and character style to inline properties. Toggle properties (bold, italic, caps) are tri-state, so an explicit `w:b val="0"` turns off a style that enables bold. `WordStyleCache` and `ExcelStyles` cache these lookups per document.
- **Units**: every OpenXML measurement goes through `Units`, which parses with the invariant culture — OpenXML always writes `.` as the decimal separator regardless of the machine locale.
- **No logging by default**: `OpenXmlHelpers.ImageLoadLogger` and `FontLoadLogger` are null unless the consumer sets them.
- **Floating images**: images that overflow their row are rendered as absolutely positioned section-level images (`WrapStyle.None`) at coordinates computed from the sheet grid, simulating Excel's overflow behaviour.

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| DocumentFormat.OpenXml | 3.5.1 | Parse DOCX/XLSX OpenXML packages |
| PdfSharp-MigraDoc | 6.2.4 | Build and render PDF documents |
| System.Drawing.Common | 10.0.12 | Image cropping and re-encoding (Windows-only, graceful fallback) |

## Building & Testing

Requirements: .NET 10 SDK.

```bash
dotnet build
dotnet test
dotnet pack -c Release
```

`global.json` opts `dotnet test` into Microsoft.Testing.Platform, which xUnit v3 requires on the .NET 10 SDK; the test project can also be run directly with `dotnet run --project src/PDFConverter.Tests`.

The test project contains 209 tests: unit tests over individual methods, and integration tests that convert in-memory OpenXML documents and assert on the rendered PDF (page counts, text, images, borders). All test documents are constructed programmatically — no external or confidential files are needed.

## Known OpenXML and MigraDoc Gotchas

- **Duplicate types that look interchangeable**: `StyleRunProperties`/`StyleParagraphProperties`, `RunPropertiesBaseStyle`/`ParagraphPropertiesBaseStyle` (what `w:docDefaults` parses as) and `ParagraphMarkRunProperties` all carry the same children as `RunProperties`/`ParagraphProperties` but are distinct classes. `GetFirstChild<RunProperties>()` silently returns null and a cast-based clone always fails — copy the children instead (`WordStyleCache.CopyProperties`).
- **`EnumValue<T>.Value.ToString()` returns garbage** in OpenXml 3.x (e.g. `"LineSpacingRuleValues { }"`). Compare with `==`, or read `.InnerText` when the raw token is needed.
- **`w:hyperlink` inside DrawingML** parses as `OpenXmlUnknownElement`, not `W.Hyperlink`. Match by `LocalName`/`NamespaceUri`.
- **MigraDoc `Row.Cells.Count` returns 0** before rendering. Use the column count from the source document.
- **MigraDoc ignores `Top` on paragraph- and line-relative shapes**; only `Margin` and `Page` anchors honour it. It also drops `SpaceBefore` on the first paragraph of a section — add it to the top margin instead.
- **MigraDoc ignores `Orientation`** when the page size is set explicitly; swap the width and height.
- **MigraDoc border extension across merged cells**: it computes `max(cell_above.bottom, cell_below.top)` across ALL columns of a merged cell, creating spurious full-width lines when the borders above are inconsistent. No API-level workaround; skip `MergeRight` and simulate centering.
- **MigraDoc drops leading plain spaces** but preserves non-breaking spaces, and has no letter-spacing primitive.

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
