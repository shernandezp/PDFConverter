# Changelog

All notable changes to this project will be documented in this file.

## [0.0.2] - 2026-02-15

### Fixed
- DOCX: Footer distance (`pgMar.Footer`) was not read from section properties, causing tables near the bottom of the page to overlap the footer in the PDF output (e.g., CheckList_Temp.docx)

## [0.0.1] - 2026-02-14

### Added
- DOCX to PDF conversion with full formatting support
- XLSX to PDF conversion with cell styles, merged cells, and images
- `byte[]` return overloads (`DocxToPdfBytes`, `XlsxToPdfBytes`) for in-memory workflows
- Stream-based input overloads for all converters
- Embedded Noto Emoji font for cross-platform emoji rendering
- Table style resolution with conditional formatting (firstRow, lastRow, firstColumn, lastColumn)
- VML textbox extraction and rendering
- Floating anchor image positioning
- Header/footer rendering with image support
- Tab stop parsing with default fallback
- Image format detection from magic bytes (not file extension)
- srcRect image cropping support
- Right indent and hanging indent support
- Landscape orientation support
- Hyperlink rendering in tables (WordprocessingML and DrawingML)
- MigraDoc border extension detection algorithm for XLSX merged cells
- Centering simulation without MergeRight via LeftIndent for border-extension merges
- 138 unit and integration tests (all in-memory, no file dependencies)

### Fixed
- Style paragraph/run property resolution (OpenXML 3.x type cast issue)
- Line spacing rule comparison (OpenXML 3.x enum ToString() issue)
- MigraDoc 6.x Row.Cells.Count returning 0
- RunFormat.ApplyTo replacing entire Font object
- Auto line spacing rendered as Exactly instead of Multiple
- ProcessRun tab/text ordering for columnar layouts
- behindDoc anchor images stretching to full page
- DrawingML hyperlinks parsed as OpenXmlUnknownElement
- Header images not grouped per paragraph
- Redundant spacer paragraph after tables causing extra pages
- XLSX: 25 bug fixes for XLSX rendering (BUG-012 through BUG-036)
- XLSX: Merged-away cells now apply borders to maintain continuous outer frame
- XLSX: Spurious full-width border lines across merged cells detected and suppressed
- XLSX: Images anchored in merged-away cells rendered through merge anchor matching
- XLSX: Floating image dimensions from EMU extents
- XLSX: Connector underscore signature lines clearly separated
- XLSX: Text labels centered under connector underscore lines
