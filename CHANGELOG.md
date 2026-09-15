# Changelog

All notable changes to this project will be documented in this file.

## [0.1.0] - 2026-09-14

### Added
- DOCX: `PAGE` and `NUMPAGES` fields, manual page breaks and `w:pageBreakBefore`
- DOCX: bullet glyphs, multi-level list labels (`%1.%2`) and `w:lvlOverride`
- DOCX: superscript/subscript, all caps, underline styles, paragraph shading, bookmark links, `w:sym` and non-breaking hyphens
- DOCX: header and footer text, with first-page and even-page variants kept separate
- DOCX: DrawingML group shapes and straight-connector lines (the rules above signature blocks)
- DOCX: table width (`w:tblW`), cell margins (`w:tblCellMar`), table indent and alignment, repeating header rows
- XLSX: theme and indexed colours, `defaultColWidth`, error and date cell types
- Images: BMP, TIFF, EMF and WMF are re-encoded so PdfSharp can embed them

### Fixed
- All: numbers were parsed with the current culture, so a locale using `,` as the decimal separator misread Excel values and OpenXML measurements
- DOCX: `w:docDefaults` was never read — it parses as `*BaseStyle` types — so the document's default font, size and spacing were ignored, as was the paragraph mark's own size
- DOCX: an explicit `w:b val="0"` was overridden by a style that enables bold, and likewise for italic and underline
- DOCX: relationship ids resolved against the main part first, so a header or footer picture could pick up a different body image; an unresolvable id returned the first picture in the package
- DOCX: table cell content lost document order — hyperlinks and content controls were appended after the surrounding text
- DOCX: pictures were scaled from the image file's aspect ratio rather than the drawing's extent, and page art was stretched to the page rather than its declared size
- DOCX: percentage column widths resolved against a placeholder width, and a word wider than its cell spilled into the next column
- DOCX: background pictures were drawn twice, and a standalone image gained a phantom empty paragraph after it
- DOCX: runs marked `xml:space="preserve"` lost their spacing; the first paragraph lost its `w:spacing w:before`
- DOCX: square and tight wrapping reserved a full-width band, displacing far more text than Word does, and a picture anchored to the first block lost its vertical offset
- XLSX: landscape sheets rendered as portrait, hidden sheets were rendered, and empty sheets produced blank pages
- XLSX: number format codes were converted by chained string replace, turning minutes into months and emitting raw format codes for anything it could not handle
- XLSX: unformatted numbers printed their full binary value (`4.099999999998545` rather than `4.1`)
- XLSX: pictures and connector lines ignored their anchor's `colOff`, and cells carried MigraDoc's default padding instead of the width Excel declares
- XLSX: text parked in a short spacer row was visible, where Excel clips a row given an explicit height
- Fonts: every system font file was read into memory at startup; faces are now indexed by path and loaded on demand
- Fonts: bold and italic are simulated for a family shipping only a regular face, and an unknown family falls back to a known text font rather than an arbitrary one
- Images: a picture reused across a document is embedded once instead of once per occurrence
- Tooling: `dotnet test` failed outright on the .NET 10 SDK, as xUnit v3 requires Microsoft.Testing.Platform; `global.json` now opts into it, so CI actually runs the suite

### Changed
- `OpenXmlHelpers` no longer forwards OpenXML parsing helpers; it keeps font registration and the diagnostic loggers. `Converters`, `DocxConverter`, `XlsxConverter` and `FontUtils` are unchanged
- Tables are no longer scaled up to fill the page when narrower than half the content width; the declared table width is used instead

## [0.0.4] - 2026-04-04

### Added
- DOCX: Structured Document Tags (SDT) support — `w:sdt` content controls now render their inner text, images, and hyperlinks in both body paragraphs and table cells
- DOCX: Theme font resolution — runs inheriting fonts from document theme (e.g., minor font) now resolve correctly
- DOCX: Diagnostic tool (PdfInspector) for inspecting generated PDF page structure, images, and text content
- DOCX: IMGLOG environment variable support in TestConsole for image processing diagnostics
- DOCX: PNG format conversion for indexed-palette and low-bit-depth images incompatible with PdfSharp

### Fixed
- DOCX: Bold formatting override in conditional table styles — explicit `w:b val="0"` now respected when conditional style applies bold
- DOCX: Anchor images with `behindDoc` flag rendered inline instead of as floating shapes at their specified positions
- DOCX: Anchor image horizontal/vertical positioning now maps Word `relativeFrom` attributes (page, margin, column, paragraph) to correct MigraDoc relative positioning
- DOCX: VML group shape images no longer extracted as individual broken inline images
- DOCX: VML fallback images no longer duplicated when DrawingML images exist in the same paragraph
- DOCX: Document default font size (docDefaults) now used as fallback when no style or run specifies a size

## [0.0.3] - 2026-02-15

### Fixed
- DOCX: Images with DrawingML hyperlinks (`a:hlinkClick` in `wp:docPr`) now produce clickable PDF link annotations via post-render overlay

## [0.0.2] - 2026-02-15

### Fixed
- DOCX: Footer distance (`pgMar.Footer`) was not read from section properties, causing tables near the bottom of the page to overlap the footer in the PDF output

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
