using DocumentFormat.OpenXml.Packaging;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal static class WordTableRenderer
{
    internal static void RenderTable(WordprocessingDocument doc, Section section, W.Table table, List<string> tempFiles)
    {
        if (section == null || table == null) return;
        var rows = table.Elements<W.TableRow>().ToList();
        if (rows.Count == 0) return;

        var migTable = section.AddTable();

        // First, determine the actual number of columns from the table data
        int actualCols = 0;
        foreach (var row in rows)
        {
            int rowCols = 0;
            foreach (var cell in row.Elements<W.TableCell>())
            {
                var tcPr = cell.GetFirstChild<W.TableCellProperties>();
                var gridSpan = tcPr?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1;
                rowCols += (int)gridSpan;
            }
            if (rowCols > actualCols) actualCols = rowCols;
        }

        if (actualCols == 0) actualCols = 1;

        // Get column widths from table grid
        var colWidths = OpenXmlHelpers.GetTableGridColumnWidths(table);

        // If we don't have enough column widths, calculate defaults
        double pageContentWidth = section.PageSetup.PageWidth.Point - 
            section.PageSetup.LeftMargin.Point - section.PageSetup.RightMargin.Point;
        if (colWidths.Count < actualCols)
        {
            double defaultColWidth = pageContentWidth / actualCols;
            while (colWidths.Count < actualCols)
                colWidths.Add(defaultColWidth);
        }

        int cols = actualCols;

        double totalTableWidth = colWidths.Take(cols).Sum();

        // Scale columns if total width exceeds page content width
        double scaleFactor = 1.0;
        if (totalTableWidth > pageContentWidth && totalTableWidth > 0)
            scaleFactor = pageContentWidth / totalTableWidth;
        else if (totalTableWidth < pageContentWidth * 0.5 && totalTableWidth > 0)
            scaleFactor = Math.Min(pageContentWidth * 0.9 / totalTableWidth, 1.5);

        for (int i = 0; i < cols; i++)
        {
            double w = i < colWidths.Count ? colWidths[i] : (pageContentWidth / cols);
            migTable.AddColumn(Unit.FromPoint(w * scaleFactor));
        }

        // Read table-level properties and borders
        var tblPr = table.GetFirstChild<W.TableProperties>();
        var tblBorders = OpenXmlHelpers.ResolveTableBorders(doc.MainDocumentPart, tblPr);

        // Apply table-level borders to MigraDoc table as defaults
        if (tblBorders != null)
        {
            // Read the actual border values from the document
            ApplyTableLevelBorders(migTable, tblBorders);
        }
        else
        {
            // No borders defined at all — set a thin default for visibility
            migTable.Borders.Width = Unit.FromPoint(0.5);
            migTable.Borders.Color = Colors.Black;
        }

        // Build cell list per row (flat list of cells, NOT a grid — cidx is the cell index, not column index)
        var cellsPerRow = new List<List<W.TableCell>>(rows.Count);
        foreach (var r in rows)
            cellsPerRow.Add(r.Elements<W.TableCell>().ToList());

        // First pass: build a column-aware map and detect vertical merge spans
        // vMerge tracking uses the COLUMN index (accounting for gridSpan), not cell index
        var vMergeSpans = new Dictionary<(int row, int col), int>();

        for (int ri = 0; ri < rows.Count; ri++)
        {
            int colIdx = 0;
            foreach (var c in cellsPerRow[ri])
            {
                var tcPr = c.GetFirstChild<W.TableCellProperties>();
                var gridSpan = (int)(tcPr?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1);
                var vMerge = tcPr?.GetFirstChild<W.VerticalMerge>();

                if (vMerge != null)
                {
                    // In OOXML: <w:vMerge val="restart"/> = start of merge
                    //           <w:vMerge/> (no val) = continuation
                    var vMergeVal = vMerge.Val;
                    bool isRestart = vMergeVal != null && vMergeVal.Value == W.MergedCellValues.Restart;

                    if (isRestart)
                    {
                        // Count how many rows below continue this merge
                        int spanDown = 0;
                        for (int rr = ri + 1; rr < rows.Count; rr++)
                        {
                            var belowCell = FindCellAtColumn(cellsPerRow[rr], colIdx);
                            if (belowCell == null) break;
                            var belowTcPr = belowCell.GetFirstChild<W.TableCellProperties>();
                            var belowVMerge = belowTcPr?.GetFirstChild<W.VerticalMerge>();
                            if (belowVMerge == null) break;
                            // Continuation cell: element exists but val is absent or val="continue"
                            var belowVal = belowVMerge.Val;
                            if (belowVal == null || belowVal.Value == W.MergedCellValues.Continue)
                                spanDown++;
                            else
                                break;
                        }
                        if (spanDown > 0)
                            vMergeSpans[(ri, colIdx)] = spanDown;
                    }
                }

                colIdx += gridSpan;
            }
        }

        // Second pass: render rows
        // Resolve conditional formatting for header row from table style
        var firstRowStyle = WordHelpers.GetTableStyleConditionalFormatting(
            doc.MainDocumentPart, tblPr, W.TableStyleOverrideValues.FirstRow);
        var firstColStyle = WordHelpers.GetTableStyleConditionalFormatting(
            doc.MainDocumentPart, tblPr, W.TableStyleOverrideValues.FirstColumn);

        for (int ri = 0; ri < rows.Count; ri++)
        {
            var mRow = migTable.AddRow();

            // Get row height if specified
            var rowPr = rows[ri].GetFirstChild<W.TableRowProperties>();
            var rowHeight = rowPr?.GetFirstChild<W.TableRowHeight>();
            if (rowHeight?.Val != null && double.TryParse(rowHeight.Val, out var rh))
                mRow.Height = Unit.FromPoint(rh / 20.0);

            bool isFirstRow = WordHelpers.IsConditionalRow(
                rows[ri], tblPr, ri, rows.Count, W.TableStyleOverrideValues.FirstRow);

            int colIdx = 0;
            foreach (var c in cellsPerRow[ri])
            {
                var tcPr = c.GetFirstChild<W.TableCellProperties>();
                var gridSpan = (int)(tcPr?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1);

                if (colIdx >= cols) break;
                var mCell = mRow.Cells[colIdx];

                // Check if this cell is a vertical merge continuation — skip rendering content
                var vMerge = tcPr?.GetFirstChild<W.VerticalMerge>();
                bool isContinuation = vMerge != null && 
                    (vMerge.Val == null || vMerge.Val.Value == W.MergedCellValues.Continue);

                if (isContinuation)
                {
                    colIdx += gridSpan;
                    continue;
                }

                // Determine which conditional style applies to this cell
                W.TableStyleProperties? activeCondStyle = null;
                if (isFirstRow && firstRowStyle != null)
                    activeCondStyle = firstRowStyle;
                else if (colIdx == 0 && firstColStyle != null)
                    activeCondStyle = firstColStyle;

                // Apply cell shading — prefer inline, then conditional style, then nothing
                var shading = tcPr?.GetFirstChild<W.Shading>()?.Fill?.Value;
                if (!string.IsNullOrEmpty(shading) && !string.Equals(shading, "auto", StringComparison.OrdinalIgnoreCase))
                {
                    try { mCell.Shading.Color = MigraDoc.DocumentObjectModel.Color.Parse("#" + shading); } catch { }
                }
                else if (activeCondStyle != null)
                {
                    var condShading = activeCondStyle.TableStyleConditionalFormattingTableCellProperties?
                        .GetFirstChild<W.Shading>()?.Fill?.Value;
                    if (!string.IsNullOrEmpty(condShading) && !string.Equals(condShading, "auto", StringComparison.OrdinalIgnoreCase))
                    {
                        try { mCell.Shading.Color = MigraDoc.DocumentObjectModel.Color.Parse("#" + condShading); } catch { }
                    }
                }

                // Apply cell borders (with table-level fallback)
                var bordersInfo = OpenXmlHelpers.GetWordCellBorders(tcPr, tblPr);
                ApplyCellBorders(mCell, bordersInfo);

                // Apply cell padding
                if (bordersInfo.PaddingTop > 0 || bordersInfo.PaddingBottom > 0)
                {
                    try
                    {
                        mCell.Format.SpaceBefore = Unit.FromPoint(bordersInfo.PaddingTop);
                        mCell.Format.SpaceAfter = Unit.FromPoint(bordersInfo.PaddingBottom);
                    }
                    catch { }
                }

                // Handle horizontal merge (grid span)
                if (gridSpan > 1)
                {
                    int mergeRight = Math.Min(gridSpan - 1, cols - colIdx - 1);
                    if (mergeRight > 0)
                        mCell.MergeRight = mergeRight;
                }

                // Handle vertical merge (restart cell sets MergeDown)
                if (vMergeSpans.TryGetValue((ri, colIdx), out var vSpan) && vSpan > 0)
                    mCell.MergeDown = vSpan;

                // Vertical alignment
                var vAlign = tcPr?.GetFirstChild<W.TableCellVerticalAlignment>()?.Val?.Value;
                if (vAlign != null)
                {
                    if (vAlign == W.TableVerticalAlignmentValues.Center)
                        mCell.VerticalAlignment = VerticalAlignment.Center;
                    else if (vAlign == W.TableVerticalAlignmentValues.Bottom)
                        mCell.VerticalAlignment = VerticalAlignment.Bottom;
                    else
                        mCell.VerticalAlignment = VerticalAlignment.Top;
                }

                // Render paragraphs in the cell
                RenderCellContent(doc, mCell, c, colWidths, colIdx, scaleFactor, tempFiles, activeCondStyle);

                colIdx += gridSpan;
            }

            // Handle rows that use DrawingML cells (<a:tc>) instead of WordprocessingML (<w:tc>)
            if (colIdx == 0 && cols > 0)
            {
                RenderDrawingMLRow(doc, rows[ri], mRow, cols);
            }
        }
    }

    /// <summary>
    /// Renders content from a table row that contains DrawingML cells (a:tc) instead of standard w:tc.
    /// Extracts text and hyperlinks from DrawingML paragraphs.
    /// </summary>
    private static void RenderDrawingMLRow(WordprocessingDocument doc, W.TableRow row, Row mRow, int cols)
    {
        // Namespace for DrawingML
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        int cellIdx = 0;
        foreach (var child in row.ChildElements)
        {
            // Look for a:tc elements (DrawingML table cells)
            if (child.LocalName != "tc" || child.NamespaceUri != aNs) continue;
            if (cellIdx >= cols) break;

            var mCell = mRow.Cells[cellIdx];
            var para = mCell.AddParagraph();

            foreach (var pChild in child.ChildElements)
            {
                if (pChild.LocalName == "p") // a:p (DrawingML paragraph)
                {
                    foreach (var rChild in pChild.ChildElements)
                    {
                        // w:hyperlink inside a:p is parsed as OpenXmlUnknownElement — match by name/namespace
                        if (rChild.LocalName == "hyperlink" && rChild.NamespaceUri == wNs)
                        {
                            // Extract r:id attribute manually
                            const string rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                            var rId = rChild.GetAttribute("id", rNs).Value;
                            string? url = null;
                            if (!string.IsNullOrEmpty(rId))
                            {
                                try { url = doc.MainDocumentPart?.HyperlinkRelationships
                                    .FirstOrDefault(r => r.Id == rId)?.Uri?.ToString(); } catch { }
                            }

                            // Get text from a:t elements inside the hyperlink
                            var hlText = string.Join("", rChild.Descendants()
                                .Where(d => d.LocalName == "t")
                                .Select(d => d.InnerText));

                            if (string.IsNullOrEmpty(hlText)) hlText = url ?? rId ?? "link";

                            if (!string.IsNullOrEmpty(url))
                            {
                                var hl2 = para.AddHyperlink(url, HyperlinkType.Web);
                                var ft = hl2.AddFormattedText(hlText);
                                ft.Color = Color.Parse("#0000FF");
                                ft.Underline = Underline.Single;
                            }
                            else
                            {
                                var ft = para.AddFormattedText(hlText);
                                ft.Color = Color.Parse("#0000FF");
                                ft.Underline = Underline.Single;
                            }
                        }
                        else if (rChild.LocalName == "r" && rChild.NamespaceUri == aNs)
                        {
                            // a:r — DrawingML run: extract a:t text
                            var text = string.Join("", rChild.Descendants()
                                .Where(d => d.LocalName == "t")
                                .Select(d => d.InnerText));
                            if (!string.IsNullOrEmpty(text))
                                para.AddFormattedText(text);
                        }
                    }
                }
            }
            cellIdx++;
        }
    }

    /// <summary>
    /// Find the Word cell that covers a given column index in a row, accounting for gridSpan.
    /// </summary>
    private static W.TableCell? FindCellAtColumn(List<W.TableCell> rowCells, int targetCol)
    {
        int col = 0;
        foreach (var c in rowCells)
        {
            var tcPr = c.GetFirstChild<W.TableCellProperties>();
            var span = (int)(tcPr?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1);
            if (col == targetCol) return c;
            col += span;
            if (col > targetCol) return null;
        }
        return null;
    }

    private static void RenderCellContent(WordprocessingDocument doc, Cell mCell, W.TableCell wordCell,
        List<double> colWidths, int colIdx, double scaleFactor, List<string> tempFiles,
        W.TableStyleProperties? condStyle = null)
    {
        // Extract conditional run formatting (bold, color) from the table style
        var condRPr = condStyle?.RunPropertiesBaseStyle;
        bool condBold = condRPr?.Bold != null && (condRPr.Bold.Val == null || condRPr.Bold.Val.Value);
        string? condColor = condRPr?.Color?.Val?.Value;

        var paragraphs = wordCell.Elements<W.Paragraph>().ToList();
        bool firstPara = true;

        foreach (var wp in paragraphs)
        {
            var para = mCell.AddParagraph();

            var pPr = wp.ParagraphProperties;
            var pFmt = WordHelpers.GetParagraphFormatting(pPr);

            para.Format.Alignment = pFmt.Alignment;
            if (pFmt.LeftIndent > 0) para.Format.LeftIndent = Unit.FromPoint(pFmt.LeftIndent);
            if (pFmt.FirstLineIndent != 0) para.Format.FirstLineIndent = Unit.FromPoint(pFmt.FirstLineIndent);

            // Cap spacing in table cells to keep layout compact
            if (pFmt.SpacingBefore > 0) para.Format.SpaceBefore = Unit.FromPoint(Math.Min(pFmt.SpacingBefore, 6));
            if (pFmt.SpacingAfter > 0) para.Format.SpaceAfter = Unit.FromPoint(Math.Min(pFmt.SpacingAfter, 6));

            bool hasContent = false;

            foreach (var run in wp.Elements<W.Run>())
            {
                var fmt = WordHelpers.ResolveRunFormatting(doc.MainDocumentPart, run, wp);

                // Merge conditional formatting from table style (lower priority than inline)
                if (condStyle != null)
                    fmt = MergeConditionalRunFormat(fmt, condBold, condColor);

                bool foundText = false;
                foreach (var textEl in run.Elements<W.Text>())
                {
                    var txt = textEl.Text;
                    if (string.IsNullOrEmpty(txt)) continue;

                    if (ConverterExtensions.ContainsEmoji(txt))
                    {
                        foreach (var (seg, isEmoji) in ConverterExtensions.SplitEmojiSegments(txt))
                        {
                            var formatted = para.AddFormattedText(seg);
                            fmt.ApplyTo(formatted);
                            if (isEmoji) formatted.Font.Name = "Noto Emoji";
                        }
                    }
                    else
                    {
                        var formatted = para.AddFormattedText(txt);
                        fmt.ApplyTo(formatted);
                    }
                    hasContent = true;
                    foundText = true;
                }

                if (!foundText)
                {
                    var innerText = run.InnerText;
                    if (!string.IsNullOrEmpty(innerText))
                    {
                        var formatted = para.AddFormattedText(innerText);
                        hasContent = true;
                        fmt.ApplyTo(formatted);
                    }
                }

                foreach (var child in run.ChildElements)
                {
                    if (child is W.Break) para.AddLineBreak();
                    else if (child is W.TabChar) para.AddTab();
                }
            }

            // Handle hyperlinks inside table cells
            foreach (var hyperlink in wp.Elements<W.Hyperlink>())
            {
                foreach (var run in hyperlink.Elements<W.Run>())
                {
                    var fmt = WordHelpers.ResolveRunFormatting(doc.MainDocumentPart, run, wp);
                    if (condStyle != null)
                        fmt = MergeConditionalRunFormat(fmt, condBold, condColor);
                    string? txt = null;
                    foreach (var textEl in run.Elements<W.Text>())
                    {
                        txt = textEl.Text;
                        if (!string.IsNullOrEmpty(txt)) break;
                    }
                    txt ??= run.InnerText;
                    if (!string.IsNullOrEmpty(txt))
                    {
                        var formatted = para.AddFormattedText(txt);
                        hasContent = true;
                        fmt.ApplyTo(formatted);
                    }
                }
            }

            // Handle images in the paragraph
            try
            {
                var infos = ConverterExtensions.GetImageInfosFromParagraph(doc, wp);
                foreach (var info in infos)
                {
                    if (info.Bytes == null || info.Bytes.Length == 0) continue;
                    var imgPath = ConverterExtensions.SaveTempImage(info.Bytes);
                    tempFiles?.Add(imgPath);
                    try
                    {
                        if (!System.IO.File.Exists(imgPath)) continue;
                        var image = para.AddImage(imgPath);
                        image.LockAspectRatio = true;

                        double maxWidth = colIdx < colWidths.Count ? 
                            colWidths[colIdx] * scaleFactor * 0.9 : 100;

                        if (info.ExtentCxEmu.HasValue)
                        {
                            double imgWidth = info.ExtentCxEmu.Value / 12700.0;
                            image.Width = Unit.FromPoint(Math.Min(imgWidth, maxWidth));
                        }
                        else
                        {
                            image.Width = Unit.FromPoint(Math.Min(100, maxWidth));
                        }
                        hasContent = true;
                    }
                    catch { }
                }
            }
            catch { }

            if (!hasContent && !firstPara)
                para.AddText(" ");

            firstPara = false;
        }

        if (paragraphs.Count == 0)
            mCell.AddParagraph();
    }

    /// <summary>
    /// Merges conditional formatting from table style into a RunFormat,
    /// using the conditional values only when the run doesn't already specify them.
    /// </summary>
    private static RunFormat MergeConditionalRunFormat(RunFormat fmt, bool condBold, string? condColor)
    {
        return fmt with
        {
            Bold = fmt.Bold || condBold,
            Color = fmt.Color ?? condColor
        };
    }

    private static void ApplyTableLevelBorders(Table migTable, W.TableBorders tblBorders)
    {
        ApplyBorderEdge(migTable.Borders.Top, tblBorders.TopBorder);
        ApplyBorderEdge(migTable.Borders.Bottom, tblBorders.BottomBorder);
        ApplyBorderEdge(migTable.Borders.Left, tblBorders.LeftBorder);
        ApplyBorderEdge(migTable.Borders.Right, tblBorders.RightBorder);

        // InsideH and InsideV are the internal gridlines
        var insideH = tblBorders.InsideHorizontalBorder;
        var insideV = tblBorders.InsideVerticalBorder;

        if (insideH != null && insideH.Val != null && insideH.Val.Value != W.BorderValues.None)
        {
            double w = 0.5;
            if (insideH.Size != null && insideH.Size.HasValue)
                w = insideH.Size.Value / 8.0;
            if (migTable.Borders.Top.Width.Point == 0)
            {
                migTable.Borders.Top.Width = Unit.FromPoint(w);
                migTable.Borders.Top.Color = Colors.Black;
            }
            if (migTable.Borders.Bottom.Width.Point == 0)
            {
                migTable.Borders.Bottom.Width = Unit.FromPoint(w);
                migTable.Borders.Bottom.Color = Colors.Black;
            }
        }
        if (insideV != null && insideV.Val != null && insideV.Val.Value != W.BorderValues.None)
        {
            double w = 0.5;
            if (insideV.Size != null && insideV.Size.HasValue)
                w = insideV.Size.Value / 8.0;
            if (migTable.Borders.Left.Width.Point == 0)
            {
                migTable.Borders.Left.Width = Unit.FromPoint(w);
                migTable.Borders.Left.Color = Colors.Black;
            }
            if (migTable.Borders.Right.Width.Point == 0)
            {
                migTable.Borders.Right.Width = Unit.FromPoint(w);
                migTable.Borders.Right.Color = Colors.Black;
            }
        }
    }

    private static void ApplyBorderEdge(Border border, W.BorderType? source)
    {
        if (source == null) return;
        if (source.Val != null && source.Val.Value == W.BorderValues.None)
        {
            border.Width = 0;
            return;
        }

        if (source.Size != null && source.Size.HasValue)
            border.Width = Unit.FromPoint(source.Size.Value / 8.0);
        else
            border.Width = Unit.FromPoint(0.5);

        var col = source.Color?.Value;
        if (!string.IsNullOrEmpty(col) && !string.Equals(col, "auto", StringComparison.OrdinalIgnoreCase))
        {
            try { border.Color = MigraDoc.DocumentObjectModel.Color.Parse("#" + col); } catch { }
        }
        else
        {
            border.Color = Colors.Black;
        }
    }

    private static void ApplyCellBorders(Cell cell, BorderInfo borders)
    {
        try
        {
            if (borders.TopWidth > 0)
            {
                cell.Borders.Top.Width = Unit.FromPoint(borders.TopWidth);
                if (!string.IsNullOrEmpty(borders.TopColor))
                    cell.Borders.Top.Color = MigraDoc.DocumentObjectModel.Color.Parse(borders.TopColor);
                else
                    cell.Borders.Top.Color = Colors.Black;
            }

            if (borders.BottomWidth > 0)
            {
                cell.Borders.Bottom.Width = Unit.FromPoint(borders.BottomWidth);
                if (!string.IsNullOrEmpty(borders.BottomColor))
                    cell.Borders.Bottom.Color = MigraDoc.DocumentObjectModel.Color.Parse(borders.BottomColor);
                else
                    cell.Borders.Bottom.Color = Colors.Black;
            }

            if (borders.LeftWidth > 0)
            {
                cell.Borders.Left.Width = Unit.FromPoint(borders.LeftWidth);
                if (!string.IsNullOrEmpty(borders.LeftColor))
                    cell.Borders.Left.Color = MigraDoc.DocumentObjectModel.Color.Parse(borders.LeftColor);
                else
                    cell.Borders.Left.Color = Colors.Black;
            }

            if (borders.RightWidth > 0)
            {
                cell.Borders.Right.Width = Unit.FromPoint(borders.RightWidth);
                if (!string.IsNullOrEmpty(borders.RightColor))
                    cell.Borders.Right.Color = MigraDoc.DocumentObjectModel.Color.Parse(borders.RightColor);
                else
                    cell.Borders.Right.Color = Colors.Black;
            }
        }
        catch { }
    }
}
