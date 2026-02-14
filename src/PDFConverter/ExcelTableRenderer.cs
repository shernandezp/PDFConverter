using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;

namespace PDFConverter;

internal static class ExcelTableRenderer
{
    internal static void RenderTable(Section section, WorksheetPart wsPart, List<string>? tempFiles = null)
    {
        if (section == null || wsPart == null) return;
        var sheetData = wsPart.Worksheet.Elements<S.SheetData>().FirstOrDefault();
        if (sheetData == null) return;
        var rows = sheetData.Elements<S.Row>().ToList();
        if (rows.Count == 0) return;

        var wbPart = wsPart.GetParentParts().OfType<WorkbookPart>().First();
        
        // Determine data bounds from cells that actually contain values
        int minDataRow = int.MaxValue, maxDataRow = 0;
        int minDataCol = int.MaxValue, maxDataCol = 0;
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 1) - 1;
            foreach (var cell in row.Elements<S.Cell>())
            {
                var colIdx = GetColumnIndex(cell.CellReference?.Value);
                if (rowIdx < minDataRow) minDataRow = rowIdx;
                if (rowIdx > maxDataRow) maxDataRow = rowIdx;
                if (colIdx < minDataCol) minDataCol = colIdx;
                if (colIdx > maxDataCol) maxDataCol = colIdx;
            }
        }
        
        if (minDataRow == int.MaxValue) return; // no data

        // Include merge ranges in bounds (they reference data cells)
        var mergeRanges = ExcelHelpers.GetMergeCellRanges(wsPart.Worksheet);
        foreach (var mr in mergeRanges)
        {
            if (mr.startRow < minDataRow) minDataRow = mr.startRow;
            if (mr.endRow > maxDataRow) maxDataRow = mr.endRow;
            if (mr.startCol < minDataCol) minDataCol = mr.startCol;
            if (mr.endCol > maxDataCol) maxDataCol = mr.endCol;
        }

        // Images expand row bounds but NOT column bounds — images anchored
        // in spacer columns before data get placed at the first data column
        var imageInfos = ExcelHelpers.GetImagesWithPositionFromWorksheet(wsPart).ToList();
        foreach (var img in imageInfos)
        {
            if (img.FromRow.HasValue && img.FromRow.Value < minDataRow) minDataRow = img.FromRow.Value;
            if (img.FromRow.HasValue && img.FromRow.Value > maxDataRow) maxDataRow = img.FromRow.Value;
            // Only expand column bounds if image is within or after data columns
            if (img.FromCol.HasValue && img.FromCol.Value >= minDataCol && img.FromCol.Value > maxDataCol)
                maxDataCol = img.FromCol.Value;
        }

        // Extract connector lines (horizontal lines used as signature underlines etc.)
        var connectorLines = ExcelHelpers.GetConnectorLines(wsPart);
        // Build lookup: row offset → list of connector columns (offset from minDataCol)
        var connectorsByRow = new Dictionary<int, List<(int fromCol, int toCol)>>();

        int numCols = maxDataCol - minDataCol + 1;
        int numRows = maxDataRow - minDataRow + 1;

        // Populate connector line lookup
        foreach (var cl in connectorLines)
        {
            int rOff = cl.Row - minDataRow;
            if (rOff < 0 || rOff >= numRows) continue;
            int cFrom = Math.Max(cl.FromCol - minDataCol, 0);
            int cTo = Math.Min(cl.ToCol - minDataCol, numCols - 1);
            if (!connectorsByRow.ContainsKey(rOff))
                connectorsByRow[rOff] = new List<(int, int)>();
            connectorsByRow[rOff].Add((cFrom, cTo));
        }

        // Get column widths for the full sheet, then slice to our data range
        var allColWidths = ExcelHelpers.GetWorksheetColumnWidths(wsPart, maxDataCol + 1);
        var colWidths = new List<double>();
        for (int i = minDataCol; i <= maxDataCol; i++)
        {
            colWidths.Add(i < allColWidths.Count ? allColWidths[i] : 48);
        }

        // Calculate page content width
        double pageContentWidth = section.PageSetup.PageWidth.Point - 
            section.PageSetup.LeftMargin.Point - section.PageSetup.RightMargin.Point;

        // Scale columns only if they exceed page width; never scale up
        double totalWidth = colWidths.Sum();
        double scaleFactor = 1.0;
        if (totalWidth > pageContentWidth && totalWidth > 0)
        {
            scaleFactor = pageContentWidth / totalWidth;
        }

        var table = section.AddTable();
        table.Borders.Width = Unit.FromPoint(0);
        
        // Center the table on the page when it's narrower than the content area
        double actualTableWidth = totalWidth * scaleFactor;
        if (actualTableWidth < pageContentWidth)
        {
            double indent = (pageContentWidth - actualTableWidth) / 2;
            table.Rows.LeftIndent = Unit.FromPoint(indent);
        }
        
        for (int i = 0; i < numCols; i++)
        {
            double w = i < colWidths.Count ? colWidths[i] : 48;
            table.AddColumn(Unit.FromPoint(w * scaleFactor));
        }

        // Index row data by original row index
        var rowDataByIndex = new Dictionary<int, S.Row>();
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 1) - 1;
            rowDataByIndex[rowIdx] = row;
        }

        // Build a sparse matrix from row data
        var cellMatrix = new S.Cell?[numRows, numCols];
        var rowHeights = new double?[numRows];
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 1) - 1;
            if (rowIdx < minDataRow || rowIdx > maxDataRow) continue;
            int rOff = rowIdx - minDataRow;
            
            // Capture row heights (use ht attribute if present)
            if (row.Height?.Value != null)
                rowHeights[rOff] = row.Height.Value;

            foreach (var cell in row.Elements<S.Cell>())
            {
                var colIdx = GetColumnIndex(cell.CellReference?.Value);
                if (colIdx < minDataCol || colIdx > maxDataCol) continue;
                int cOff = colIdx - minDataCol;
                cellMatrix[rOff, cOff] = cell;
            }
        }

        // Track which cells are covered by horizontal merges only
        // MergeDown is avoided due to MigraDoc GetMinMergedCell bug
        var mergedCells = new HashSet<(int row, int col)>();
        var offsetMerges = new List<(int startRow, int startCol, int endRow, int endCol)>();
        // Detect merges where MigraDoc would extend partial borders from the row
        // above across the full merged width. These merges skip MergeRight but
        // simulate centering via left indent instead.
        var borderExtensionMerges = new HashSet<(int row, int col)>();
        foreach (var mr in mergeRanges)
        {
            int sRow = mr.startRow - minDataRow;
            int eRow = mr.endRow - minDataRow;
            int sCol = mr.startCol - minDataCol;
            int eCol = mr.endCol - minDataCol;
            offsetMerges.Add((sRow, sCol, eRow, eCol));
            for (int c = sCol + 1; c <= eCol; c++)
            {
                if (sRow >= 0 && c >= 0 && sRow < numRows && c < numCols)
                    mergedCells.Add((sRow, c));
            }

            // Check for border extension: merged cell has no top border but
            // row above has inconsistent bottom borders (some present, some not)
            int colSpan = Math.Min(eCol, numCols - 1) - sCol;
            if (colSpan > 0 && sRow > 0 && sRow < numRows)
            {
                var anchorCell = cellMatrix[sRow, sCol];
                var anchorStyle = ExcelHelpers.GetCellStyleInfo(wbPart, anchorCell?.StyleIndex?.Value);
                if (anchorStyle.Borders.TopWidth == 0)
                {
                    bool hasAbove = false, missingAbove = false;
                    for (int mc = sCol; mc <= sCol + colSpan; mc++)
                    {
                        if (mc < 0 || mc >= numCols) continue;
                        var aboveCell = cellMatrix[sRow - 1, mc];
                        var aboveStyle = ExcelHelpers.GetCellStyleInfo(wbPart,
                            aboveCell?.StyleIndex?.Value);
                        if (aboveStyle.Borders.BottomWidth > 0) hasAbove = true;
                        else missingAbove = true;
                    }
                    // Only skip merge if the anchor column is wide enough to hold
                    // text with centering indent (otherwise text wraps unusably)
                    double anchorColW = (sCol >= 0 && sCol < colWidths.Count)
                        ? colWidths[sCol] * scaleFactor : 0;
                    if (hasAbove && missingAbove && anchorColW >= 40)
                        borderExtensionMerges.Add((sRow, sCol));
                }
            }
        }

        // Build image lookup by row (using offset coordinates)
        // Group images that overlap vertically into the same row for side-by-side rendering
        var imagesByRow = new Dictionary<int, List<(ExcelHelpers.ExcelImageInfo info, string path)>>();
        var processedImages = new List<(ExcelHelpers.ExcelImageInfo info, string path, int assignedRow)>();
        foreach (var imgInfo in imageInfos)
        {
            if (imgInfo.Bytes == null || imgInfo.Bytes.Length == 0) continue;
            int imgRow = (imgInfo.FromRow ?? 0) - minDataRow;
            if (imgRow < 0) imgRow = 0;
            if (imgRow >= numRows) imgRow = numRows - 1;
            
            var imgPath = ConverterExtensions.SaveTempImage(imgInfo.Bytes);
            tempFiles?.Add(imgPath);
            processedImages.Add((imgInfo, imgPath, imgRow));
        }
        
        // Check for images in different columns that overlap vertically — merge to earliest row
        for (int i = 0; i < processedImages.Count; i++)
        {
            var (infoA, _, rowA) = processedImages[i];
            int fromRowA = infoA.FromRow ?? 0;
            int toRowA = infoA.ToRow ?? (fromRowA + 1);
            for (int j = i + 1; j < processedImages.Count; j++)
            {
                var (infoB, pathB, rowB) = processedImages[j];
                int fromRowB = infoB.FromRow ?? 0;
                int toRowB = infoB.ToRow ?? (fromRowB + 1);
                // Different columns, overlapping rows
                if (infoA.FromCol != infoB.FromCol && fromRowA < toRowB && fromRowB < toRowA)
                {
                    int mergedRow = Math.Min(rowA, rowB);
                    processedImages[i] = (processedImages[i].info, processedImages[i].path, mergedRow);
                    processedImages[j] = (processedImages[j].info, processedImages[j].path, mergedRow);
                }
            }
        }
        
        foreach (var (info, path, row) in processedImages)
        {
            if (!imagesByRow.ContainsKey(row))
                imagesByRow[row] = new List<(ExcelHelpers.ExcelImageInfo, string)>();
            imagesByRow[row].Add((info, path));
        }

        // Collapse consecutive empty rows to minimal height.
        // A row is "empty" if it has no cell data, no images, no connectors, and is not part of a merge.
        var mergeRowSet = new HashSet<int>();
        foreach (var mr in offsetMerges)
            for (int rr = mr.startRow; rr <= mr.endRow; rr++) mergeRowSet.Add(rr);

        // Protect the row where each image renders from collapsing.
        // For spacer images (before data range), also protect their full span rows
        // since they'll be rendered as floating images and need the vertical space.
        var imageSpanRows = new HashSet<int>();
        foreach (var (info, path, assignedRow) in processedImages)
        {
            imageSpanRows.Add(assignedRow);
            // Protect full span for spacer images
            if ((info.FromCol ?? 0) < minDataCol && info.FromRow.HasValue && info.ToRow.HasValue)
            {
                for (int sr = info.FromRow.Value; sr < info.ToRow.Value; sr++)
                {
                    int rOff = sr - minDataRow;
                    if (rOff >= 0 && rOff < numRows)
                        imageSpanRows.Add(rOff);
                }
            }
        }

        // Save original row heights before collapsing (for image span calculations)
        var originalRowHeights = new double?[numRows];
        Array.Copy(rowHeights, originalRowHeights, numRows);

        for (int r = 0; r < numRows; r++)
        {
            if (imagesByRow.ContainsKey(r) || connectorsByRow.ContainsKey(r) 
                || mergeRowSet.Contains(r) || imageSpanRows.Contains(r))
                continue;
            bool hasData = false;
            for (int c = 0; c < numCols; c++)
            {
                if (cellMatrix[r, c] != null) { hasData = true; break; }
            }
            if (!hasData && rowHeights[r] == null)
                rowHeights[r] = 6.0; // collapse to 6pt (minimal spacer, preserves some spacing)
        }

        // Identify spacer images that share their row with text — these will be
        // rendered as absolutely positioned section-level images so they don't
        // expand the row height (MigraDoc can't overflow images like Excel does).
        var floatingImages = new List<(ExcelHelpers.ExcelImageInfo info, string path, int row)>();
        var floatingImageSet = new HashSet<(int row, string path)>();
        foreach (var kvp in imagesByRow)
        {
            int imgR = kvp.Key;
            bool hasSpacerImg = false;
            foreach (var (imgInf, _) in kvp.Value)
                if ((imgInf.FromCol ?? 0) < minDataCol) { hasSpacerImg = true; break; }
            if (!hasSpacerImg) continue;
            bool hasRowText = false;
            for (int tc = 0; tc < numCols; tc++)
            {
                var tcCell = cellMatrix[imgR, tc];
                if (tcCell != null && !string.IsNullOrEmpty(GetCellValue(tcCell, wbPart)))
                { hasRowText = true; break; }
            }
            if (!hasRowText) continue;
            foreach (var (imgInf, imgPath) in kvp.Value)
            {
                if ((imgInf.FromCol ?? 0) < minDataCol)
                {
                    floatingImages.Add((imgInf, imgPath, imgR));
                    floatingImageSet.Add((imgR, imgPath));
                }
            }
        }

        // Track cumulative Y offset per row (for positioning floating images later)
        var rowYOffsets = new double[numRows];
        double cumulativeY = 0;
        for (int ri = 0; ri < numRows; ri++)
        {
            rowYOffsets[ri] = cumulativeY;
            cumulativeY += rowHeights[ri] ?? 14.5;
        }

        for (int r = 0; r < numRows; r++)
        {
            // If this row has both connectors and text data, render connectors in a separate row first
            bool hasConnectors = connectorsByRow.ContainsKey(r);
            bool hasTextData = false;
            if (hasConnectors)
            {
                for (int tc = 0; tc < numCols; tc++)
                {
                    var tcCell = cellMatrix[r, tc];
                    if (tcCell != null && !string.IsNullOrEmpty(GetCellValue(tcCell, wbPart)))
                    { hasTextData = true; break; }
                }
            }
            
            if (hasConnectors && hasTextData)
            {
                // Add a dedicated row for connector underscore lines
                var connRow = table.AddRow();
                connRow.Height = Unit.FromPoint(14);
                foreach (var (cFrom, cTo) in connectorsByRow[r])
                {
                    if (cFrom >= 0 && cFrom < numCols)
                    {
                        var connCell = connRow.Cells[cFrom];
                        double lineW = colWidths[cFrom] * scaleFactor;
                        // Use ~65% of cell width so the two lines are clearly separate
                        int underscoreCount = Math.Max(1, (int)(lineW * 0.65 / 4.5));
                        var linePara = connCell.AddParagraph();
                        var lineText = linePara.AddFormattedText(new string('_', underscoreCount));
                        lineText.Size = 10;
                        linePara.Format.Alignment = ParagraphAlignment.Center;
                    }
                }
            }

            var prow = table.AddRow();
            
            // Apply row height
            double effectiveRowHeight = rowHeights[r] ?? 14.5;
            
            prow.Height = Unit.FromPoint(effectiveRowHeight);
            
            for (int c = 0; c < numCols; c++)
            {
                bool isMergedAway = mergedCells.Contains((r, c));
                
                var cell = cellMatrix[r, c];
                var target = prow.Cells[c];
                
                // For merged-away cells, only apply borders (skip content/alignment/images)
                if (isMergedAway)
                {
                    if (cell != null)
                    {
                        var mergedStyle = ExcelHelpers.GetCellStyleInfo(wbPart, cell.StyleIndex?.Value);
                        ApplyCellBorders(target, mergedStyle.Borders);
                    }
                    continue;
                }
                
                string text = "";
                uint? styleIndex = null;
                
                if (cell != null)
                {
                    text = GetCellValue(cell, wbPart);
                    styleIndex = cell.StyleIndex?.Value;
                }

                var cellStyle = ExcelHelpers.GetCellStyleInfo(wbPart, styleIndex);

                if (!string.IsNullOrEmpty(cellStyle.FillColor))
                {
                    try { target.Shading.Color = MigraDoc.DocumentObjectModel.Color.Parse(cellStyle.FillColor); } catch { }
                }

                if (!string.IsNullOrEmpty(cellStyle.HorizontalAlignment))
                {
                    target.Format.Alignment = cellStyle.HorizontalAlignment.ToLowerInvariant() switch
                    {
                        "center" => ParagraphAlignment.Center,
                        "right" => ParagraphAlignment.Right,
                        "justify" => ParagraphAlignment.Justify,
                        _ => ParagraphAlignment.Left
                    };
                }

                if (!string.IsNullOrEmpty(cellStyle.VerticalAlignment))
                {
                    target.VerticalAlignment = cellStyle.VerticalAlignment.ToLowerInvariant() switch
                    {
                        "center" => VerticalAlignment.Center,
                        "bottom" => VerticalAlignment.Bottom,
                        _ => VerticalAlignment.Top
                    };
                }

                ApplyCellBorders(target, cellStyle.Borders);

                // Check for merge (MergeRight only, MergeDown avoided)
                foreach (var mrItem in offsetMerges)
                {
                    if (r == mrItem.startRow && c == mrItem.startCol)
                    {
                        var colSpan = Math.Min(mrItem.endCol - mrItem.startCol, numCols - 1 - c);
                        if (colSpan > 0)
                        {
                            if (borderExtensionMerges.Contains((r, c)))
                            {
                                // Skip MergeRight to prevent MigraDoc border extension.
                                // Simulate centering by computing the full merge width
                                // and applying a left indent to the anchor cell.
                                // With Center alignment, MigraDoc centers text within
                                // (cellWidth - leftIndent), so to center across the
                                // full merge: leftIndent = mergeWidth - anchorWidth
                                double mergeWidth = 0;
                                for (int mc = c; mc <= c + colSpan && mc < colWidths.Count; mc++)
                                    mergeWidth += colWidths[mc] * scaleFactor;
                                double anchorWidth = colWidths[c] * scaleFactor;
                                if (target.Format.Alignment == ParagraphAlignment.Center && mergeWidth > anchorWidth)
                                    target.Format.LeftIndent = Unit.FromPoint(mergeWidth - anchorWidth);
                            }
                            else
                            {
                                target.MergeRight = colSpan;
                            }
                        }
                        break;
                    }
                }

                // Add image if one is anchored to this cell (or within its merge range)
                if (imagesByRow.TryGetValue(r, out var imgs))
                {
                    int actualCol = c + minDataCol;
                    // Determine merge range end column for this cell (if it's a merge anchor)
                    int mergeEndActualCol = actualCol;
                    foreach (var mrItem in offsetMerges)
                    {
                        if (r == mrItem.startRow && c == mrItem.startCol && mrItem.endCol > mrItem.startCol)
                        {
                            mergeEndActualCol = mrItem.endCol + minDataCol;
                            break;
                        }
                    }
                    foreach (var (imgInf, imgPath) in imgs)
                    {
                        int imgFromCol = imgInf.FromCol ?? 0;
                        // Images from before the data range go to column 0
                        bool fromSpacer = imgFromCol < minDataCol && c == 0;
                        // Skip images that will be rendered as floating (absolute positioned)
                        if (floatingImageSet.Contains((r, imgPath))) continue;
                        // Match if image column equals this cell's column OR falls within its merge range
                        bool match = (imgFromCol == actualCol) || fromSpacer
                            || (imgFromCol > actualCol && imgFromCol <= mergeEndActualCol);
                        if (match && System.IO.File.Exists(imgPath))
                        {
                            try
                            {
                                // Calculate max width for the image
                                double maxImgW;
                                if (fromSpacer)
                                {
                                    // Spacer-column images: use full table width
                                    maxImgW = actualTableWidth;
                                }
                                else
                                {
                                    // Cell-anchored images: use cell/merged width
                                    maxImgW = colWidths[c] * scaleFactor;
                                    foreach (var mrItem in offsetMerges)
                                    {
                                        if (r == mrItem.startRow && c == mrItem.startCol && mrItem.endCol > mrItem.startCol)
                                        {
                                            maxImgW = 0;
                                            for (int mc = mrItem.startCol; mc <= mrItem.endCol && mc < colWidths.Count; mc++)
                                                maxImgW += colWidths[mc] * scaleFactor;
                                            break;
                                        }
                                    }
                                    maxImgW = Math.Min(maxImgW, actualTableWidth);
                                    // When multiple images share a row, reduce width and add
                                    // left indent on non-first images for visual separation
                                    if (imgs.Count > 1)
                                        maxImgW = Math.Max(10, maxImgW - 12);
                                }

                                var imgPara = target.AddParagraph();
                                // Shift non-first images right for spacing in multi-image rows
                                if (imgs.Count > 1 && c > 0)
                                    imgPara.Format.LeftIndent = Unit.FromPoint(12);
                                var image = imgPara.AddImage(imgPath);
                                
                                // Calculate row span height for this image using original row heights
                                double imgRowSpanHeight = effectiveRowHeight;
                                if (imgInf.ToRow.HasValue && imgInf.FromRow.HasValue)
                                {
                                    imgRowSpanHeight = 0;
                                    for (int sr = imgInf.FromRow.Value; sr < imgInf.ToRow.Value; sr++)
                                    {
                                        int rOff = sr - minDataRow;
                                        if (rOff >= 0 && rOff < numRows)
                                            imgRowSpanHeight += originalRowHeights[rOff] ?? 14.5;
                                        else
                                            imgRowSpanHeight += 14.5;
                                    }
                                }
                                if (imgRowSpanHeight <= 0) imgRowSpanHeight = effectiveRowHeight;

                                // When the image's cell has text, clamp to single row height
                                bool cellHasText = !string.IsNullOrEmpty(text);
                                image.LockAspectRatio = false;
                                double maxImgH = cellHasText ? effectiveRowHeight : imgRowSpanHeight;
                                if (imgInf.WidthEmu.HasValue && imgInf.WidthEmu.Value > 0)
                                {
                                    double widthPts = imgInf.WidthEmu.Value / 12700.0;
                                    image.Width = Unit.FromPoint(Math.Min(widthPts, maxImgW));
                                    double heightPts = imgInf.HeightEmu.HasValue && imgInf.HeightEmu.Value > 0
                                        ? imgInf.HeightEmu.Value / 12700.0
                                        : maxImgH;
                                    image.Height = Unit.FromPoint(Math.Min(heightPts, maxImgH));
                                }
                                else if (imgInf.ToCol.HasValue && imgInf.FromCol.HasValue)
                                {
                                    // Zero extent: calculate size from anchor col/row span
                                    double spanW = 0;
                                    for (int sc = imgInf.FromCol.Value; sc < imgInf.ToCol.Value && sc < allColWidths.Count; sc++)
                                        spanW += allColWidths[sc];
                                    if (spanW <= 0) spanW = colWidths[c] * scaleFactor;
                                    image.Width = Unit.FromPoint(Math.Min(spanW, maxImgW));
                                    image.Height = Unit.FromPoint(maxImgH);
                                }
                                else
                                {
                                    double defaultW = Math.Min(colWidths[c] * scaleFactor, Unit.FromCentimeter(8).Point);
                                    image.Width = Unit.FromPoint(defaultW);
                                    image.Height = Unit.FromPoint(maxImgH);
                                }
                            }
                            catch { }
                        }
                    }
                }

                var para = target.AddParagraph();

                // Propagate cell alignment to the paragraph explicitly.
                // When connectors were split into a dedicated row above, center the
                // text paragraph so "Firma" aligns under the centered underscore lines.
                bool isConnectorCol = false;
                if (hasConnectors && hasTextData)
                    foreach (var (cF, _) in connectorsByRow[r])
                        if (cF == c) { isConnectorCol = true; break; }
                para.Format.Alignment = isConnectorCol ? ParagraphAlignment.Center : target.Format.Alignment;

                // Render connector lines inline only when there's no text in this row
                // (otherwise they were already rendered in a dedicated row above)
                if (!hasTextData && connectorsByRow.TryGetValue(r, out var connectors))
                {
                    foreach (var (cFrom, cTo) in connectors)
                    {
                        // Only render in the starting cell of each connector's range
                        if (c == cFrom)
                        {
                            // Use only the starting cell's width for the underscores
                            double lineW = colWidths[c] * scaleFactor;
                            int underscoreCount = Math.Max(1, (int)(lineW / 4.5));
                            var linePara = target.AddParagraph();
                            var lineText = linePara.AddFormattedText(new string('_', underscoreCount));
                            lineText.Size = 10;
                            linePara.Format.Alignment = ParagraphAlignment.Center;
                            break;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(text))
                {
                    var formatted = para.AddFormattedText(text);
                    double fsize = cellStyle.FontSize ?? 10;
                    formatted.Size = fsize;
                    if (!string.IsNullOrEmpty(cellStyle.FontFamily))
                    {
                        try { formatted.Font.Name = cellStyle.FontFamily; } catch { }
                    }
                    if (!string.IsNullOrEmpty(cellStyle.FontColor))
                    {
                        try { formatted.Color = MigraDoc.DocumentObjectModel.Color.Parse("#" + cellStyle.FontColor); } catch { }
                    }
                    if (cellStyle.Bold) formatted.Bold = true;
                    if (cellStyle.Italic) formatted.Italic = true;
                }
                else
                {
                    para.Format.Font.Size = 1;
                    para.Format.SpaceBefore = 0;
                    para.Format.SpaceAfter = 0;
                    para.Format.LineSpacing = Unit.FromPoint(1);
                }
            }
        }

        // Render floating images (spacer images that share rows with text)
        // as absolutely positioned section-level images
        double tableIndent = 0;
        if (actualTableWidth < pageContentWidth)
            tableIndent = (pageContentWidth - actualTableWidth) / 2;
        double topMargin = section.PageSetup.TopMargin.Point;
        double leftMargin = section.PageSetup.LeftMargin.Point;

        foreach (var (fImgInf, fImgPath, fRow) in floatingImages)
        {
            if (!System.IO.File.Exists(fImgPath)) continue;
            try
            {
                // Calculate image dimensions — prefer EMU extents if available
                double fImgH;
                if (fImgInf.HeightEmu.HasValue && fImgInf.HeightEmu.Value > 0)
                {
                    fImgH = fImgInf.HeightEmu.Value / 12700.0;
                }
                else
                {
                    fImgH = rowHeights[fRow] ?? 14.5;
                    if (fImgInf.ToRow.HasValue && fImgInf.FromRow.HasValue)
                    {
                        fImgH = 0;
                        for (int sr = fImgInf.FromRow.Value; sr < fImgInf.ToRow.Value; sr++)
                        {
                            int rOff = sr - minDataRow;
                            fImgH += (rOff >= 0 && rOff < numRows)
                                ? originalRowHeights[rOff] ?? 14.5 : 14.5;
                        }
                    }
                    if (fImgH <= 0) fImgH = 14.5;
                }

                double fImgW;
                if (fImgInf.WidthEmu.HasValue && fImgInf.WidthEmu.Value > 0)
                {
                    fImgW = fImgInf.WidthEmu.Value / 12700.0;
                }
                else
                {
                    fImgW = 0;
                    if (fImgInf.ToCol.HasValue && fImgInf.FromCol.HasValue)
                    {
                        for (int sc = fImgInf.FromCol.Value; sc < fImgInf.ToCol.Value && sc < allColWidths.Count; sc++)
                            fImgW += allColWidths[sc];
                    }
                    if (fImgW <= 0) fImgW = 80;
                }

                // X position: align to the left edge of the table (column 0)
                double imgX = leftMargin + tableIndent;

                // Y position: top margin + cumulative row heights up to this row
                // Account for any connector rows added before this row
                double imgY = topMargin + rowYOffsets[fRow];

                var floatImg = section.AddImage(fImgPath);
                floatImg.Width = Unit.FromPoint(fImgW);
                floatImg.Height = Unit.FromPoint(fImgH);
                floatImg.LockAspectRatio = false;
                floatImg.RelativeVertical = MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Page;
                floatImg.RelativeHorizontal = MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Page;
                floatImg.Top = MigraDoc.DocumentObjectModel.Shapes.TopPosition.Parse(imgY.ToString("F1") + "pt");
                floatImg.Left = MigraDoc.DocumentObjectModel.Shapes.LeftPosition.Parse(imgX.ToString("F1") + "pt");
                floatImg.WrapFormat.Style = MigraDoc.DocumentObjectModel.Shapes.WrapStyle.None;
            }
            catch { }
        }
    }

    private static void ApplyCellBorders(Cell target, BorderInfo borders)
    {
        try
        {
            if (borders.TopWidth > 0)
            {
                target.Borders.Top.Width = Unit.FromPoint(borders.TopWidth);
                if (!string.IsNullOrEmpty(borders.TopColor))
                    target.Borders.Top.Color = MigraDoc.DocumentObjectModel.Color.Parse(borders.TopColor);
            }

            if (borders.BottomWidth > 0)
            {
                target.Borders.Bottom.Width = Unit.FromPoint(borders.BottomWidth);
                if (!string.IsNullOrEmpty(borders.BottomColor))
                    target.Borders.Bottom.Color = MigraDoc.DocumentObjectModel.Color.Parse(borders.BottomColor);
            }

            if (borders.LeftWidth > 0)
            {
                target.Borders.Left.Width = Unit.FromPoint(borders.LeftWidth);
                if (!string.IsNullOrEmpty(borders.LeftColor))
                    target.Borders.Left.Color = MigraDoc.DocumentObjectModel.Color.Parse(borders.LeftColor);
            }

            if (borders.RightWidth > 0)
            {
                target.Borders.Right.Width = Unit.FromPoint(borders.RightWidth);
                if (!string.IsNullOrEmpty(borders.RightColor))
                    target.Borders.Right.Color = MigraDoc.DocumentObjectModel.Color.Parse(borders.RightColor);
            }
        }
        catch { }
    }

    private static int GetColumnIndex(string? cellRef)
    {
        if (string.IsNullOrEmpty(cellRef)) return 0;
        var letters = new string(cellRef.TakeWhile(char.IsLetter).ToArray());
        int index = 0;
        foreach (char c in letters.ToUpperInvariant())
        {
            index = index * 26 + (c - 'A' + 1);
        }
        return index - 1;
    }

    private static string GetCellValue(S.Cell cell, WorkbookPart wbPart)
    {
        // Handle inline strings first
        if (cell.DataType != null && cell.DataType == S.CellValues.InlineString)
            return cell.InlineString?.Text?.Text ?? cell.InnerText ?? string.Empty;
        
        // Shared strings
        if (cell.DataType != null && cell.DataType == S.CellValues.SharedString)
        {
            var rawRef = cell.CellValue?.Text ?? cell.InnerText;
            if (int.TryParse(rawRef, out var si))
            {
                var sst = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sst != null)
                {
                    var items = sst.SharedStringTable.Elements<S.SharedStringItem>().ToList();
                    if (si < items.Count)
                        return items[si].InnerText;
                }
            }
            return rawRef ?? string.Empty;
        }
        
        // Boolean
        if (cell.DataType != null && cell.DataType == S.CellValues.Boolean)
        {
            var raw = cell.CellValue?.Text ?? cell.InnerText ?? string.Empty;
            return raw == "0" ? "FALSE" : "TRUE";
        }
        
        // Formula: prefer cached value
        if (cell.CellFormula != null)
        {
            var cached = cell.CellValue?.Text;
            if (!string.IsNullOrEmpty(cached))
            {
                // Format the cached value
                var styleIndex = cell.StyleIndex?.Value;
                var styleInfo = ExcelHelpers.GetCellStyleInfo(wbPart, styleIndex);
                var numFmtId = styleInfo.NumberFormatId;
                var fmt = ExcelHelpers.GetNumberFormatString(wbPart, numFmtId);
                
                if (!string.IsNullOrEmpty(fmt) && double.TryParse(cached, out var d))
                {
                    return FormatNumber(d, fmt);
                }
                return cached;
            }
            return ""; // Don't show _formula_ - just empty
        }

        var rawValue = cell.CellValue?.Text ?? cell.InnerText ?? string.Empty;

        // Apply number formatting
        var cellStyleIndex = cell.StyleIndex?.Value;
        var cellStyleInfo = ExcelHelpers.GetCellStyleInfo(wbPart, cellStyleIndex);
        var formatId = cellStyleInfo.NumberFormatId;
        var format = ExcelHelpers.GetNumberFormatString(wbPart, formatId);

        if (!string.IsNullOrEmpty(format) && double.TryParse(rawValue, out var num))
        {
            return FormatNumber(num, format);
        }

        return rawValue;
    }

    private static string FormatNumber(double value, string format)
    {
        // Check for date-like format
        if (format.IndexOf('M', StringComparison.OrdinalIgnoreCase) >= 0 || 
            format.IndexOf('y', StringComparison.OrdinalIgnoreCase) >= 0 || 
            format.IndexOf('d', StringComparison.OrdinalIgnoreCase) >= 0 || 
            format.IndexOf('H', StringComparison.OrdinalIgnoreCase) >= 0)
        {
            try
            {
                var dt = DateTime.FromOADate(value);
                return dt.ToString(format);
            }
            catch { }
        }
        
        try 
        { 
            return value.ToString(format); 
        } 
        catch 
        { 
            return value.ToString(); 
        }
    }
}
