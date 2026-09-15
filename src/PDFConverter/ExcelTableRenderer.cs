using DocumentFormat.OpenXml.Packaging;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace PDFConverter;

internal static class ExcelTableRenderer
{
    public static void RenderSheet(Section section, WorksheetPart wsPart, TempImageStore images)
    {
        var workbookPart = wsPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
        if (workbookPart == null) return;

        var styles = ExcelStyles.For(workbookPart);
        var sheet = SheetGrid.Build(wsPart, styles);
        if (sheet == null) return;

        var contentWidth = section.PageSetup.PageWidth.Point
            - section.PageSetup.LeftMargin.Point - section.PageSetup.RightMargin.Point;
        sheet.ScaleToWidth(contentWidth);

        var placement = ImagePlacement.Build(sheet, images);
        sheet.CollapseEmptyRows(placement);

        var table = section.AddTable();
        table.Borders.Width = 0;
        table.LeftPadding = 0;
        table.RightPadding = 0;
        foreach (var width in sheet.ColumnWidths) table.AddColumn(Unit.FromPoint(width));

        var indent = Math.Max(0, (contentWidth - sheet.TableWidth) / 2);
        if (indent > 0) table.Rows.LeftIndent = Unit.FromPoint(indent);

        for (var row = 0; row < sheet.RowCount; row++)
            RenderRow(table, sheet, styles, placement, row);

        RenderFloatingImages(section, sheet, placement, indent);
    }

    sealed class SheetGrid
    {
        public required int MinRow { get; init; }
        public required int MinColumn { get; init; }
        public required int RowCount { get; init; }
        public required int ColumnCount { get; init; }
        public required S.Cell?[,] Cells { get; init; }
        public required double?[] RowHeights { get; init; }
        public required bool[] ClipsToHeight { get; init; }
        public required double[] OriginalRowHeights { get; init; }
        public required List<double> ColumnWidths { get; init; }
        public required List<double> SheetColumnWidths { get; init; }
        public required List<MergeRange> Merges { get; init; }
        public required HashSet<(int Row, int Column)> CoveredByMerge { get; init; }
        public required HashSet<(int Row, int Column)> SkipHorizontalMerge { get; init; }
        public required Dictionary<int, List<Connector>> Connectors { get; init; }
        public double ScaleFactor { get; private set; } = 1.0;
        public required List<ExcelHelpers.ExcelImageInfo> Images { get; init; }
        public required ExcelStyles Styles { get; init; }

        public double TableWidth => ColumnWidths.Sum();

        public double RowHeight(int row) => RowHeights[row] ?? LayoutDefaults.ExcelRowHeightPoints;

        public double ColumnOffset(int column)
        {
            double offset = 0;
            for (var i = 0; i < column && i < ColumnWidths.Count; i++) offset += ColumnWidths[i];
            return offset;
        }

        static double SheetOffset(List<double> widths, int column)
        {
            double offset = 0;
            for (var i = 0; i < column && i < widths.Count; i++) offset += widths[i];
            return offset;
        }

        public double SheetStart(int sheetColumn, double offsetPoints) =>
            SheetOffset(SheetColumnWidths, sheetColumn) + offsetPoints
                - SheetOffset(SheetColumnWidths, MinColumn);

        public int ColumnAt(double offset)
        {
            double running = 0;
            for (var i = 0; i < ColumnCount; i++)
            {
                running += ColumnWidths[i];
                if (offset < running) return i;
            }
            return ColumnCount - 1;
        }

        public bool HidesText(int row, double fontSize) =>
            ClipsToHeight[row] && RowHeight(row) < fontSize * LayoutDefaults.ExcelLineHeightRatio;

        public S.Cell? CellAt(int row, int column) =>
            row >= 0 && row < RowCount && column >= 0 && column < ColumnCount ? Cells[row, column] : null;

        public string CellText(int row, int column)
        {
            var cell = CellAt(row, column);
            return cell == null ? string.Empty : ReadCellValue(cell, Styles);
        }

        public bool RowHasText(int row)
        {
            for (var column = 0; column < ColumnCount; column++)
                if (!string.IsNullOrEmpty(CellText(row, column))) return true;
            return false;
        }

        public static SheetGrid? Build(WorksheetPart wsPart, ExcelStyles styles)
        {
            var worksheet = wsPart.Worksheet;
            var sheetData = worksheet?.Elements<S.SheetData>().FirstOrDefault();
            var rows = sheetData?.Elements<S.Row>().ToList() ?? [];
            if (rows.Count == 0) return null;

            int minRow = int.MaxValue, maxRow = 0, minColumn = int.MaxValue, maxColumn = 0;
            foreach (var row in rows)
            {
                var rowIndex = (int)(row.RowIndex?.Value ?? 1) - 1;
                foreach (var cell in row.Elements<S.Cell>())
                {
                    var column = ExcelHelpers.GetColumnIndex(cell.CellReference?.Value);
                    if (column < 0) continue;
                    minRow = Math.Min(minRow, rowIndex);
                    maxRow = Math.Max(maxRow, rowIndex);
                    minColumn = Math.Min(minColumn, column);
                    maxColumn = Math.Max(maxColumn, column);
                }
            }
            if (minRow == int.MaxValue) return null;

            var mergeRanges = ExcelHelpers.GetMergeCellRanges(worksheet!);
            foreach (var merge in mergeRanges)
            {
                minRow = Math.Min(minRow, merge.startRow);
                maxRow = Math.Max(maxRow, merge.endRow);
                minColumn = Math.Min(minColumn, merge.startCol);
                maxColumn = Math.Max(maxColumn, merge.endCol);
            }

            var connectorLines = ExcelHelpers.GetConnectorLines(wsPart);

            var images = ExcelHelpers.GetImagesWithPositionFromWorksheet(wsPart);
            foreach (var image in images)
            {
                if (image.FromRow is not { } imageRow) continue;
                minRow = Math.Min(minRow, imageRow);
                maxRow = Math.Max(maxRow, imageRow);
                if (image.FromCol is { } imageColumn && imageColumn >= minColumn)
                    maxColumn = Math.Max(maxColumn, imageColumn);
            }

            var rowCount = maxRow - minRow + 1;
            var columnCount = maxColumn - minColumn + 1;

            var cells = new S.Cell?[rowCount, columnCount];
            var rowHeights = new double?[rowCount];
            var clipsToHeight = new bool[rowCount];
            foreach (var row in rows)
            {
                var rowIndex = (int)(row.RowIndex?.Value ?? 1) - 1 - minRow;
                if (rowIndex < 0 || rowIndex >= rowCount) continue;
                if (row.Height?.Value is { } height) rowHeights[rowIndex] = height;
                clipsToHeight[rowIndex] = row.CustomHeight?.Value == true;

                foreach (var cell in row.Elements<S.Cell>())
                {
                    var column = ExcelHelpers.GetColumnIndex(cell.CellReference?.Value) - minColumn;
                    if (column >= 0 && column < columnCount) cells[rowIndex, column] = cell;
                }
            }

            var sheetColumnWidths = ExcelHelpers.GetWorksheetColumnWidths(wsPart, maxColumn + 1);
            var columnWidths = new List<double>(columnCount);
            for (var i = minColumn; i <= maxColumn; i++)
                columnWidths.Add(i < sheetColumnWidths.Count
                    ? sheetColumnWidths[i]
                    : LayoutDefaults.ExcelColumnWidthPoints);

            var connectors = new Dictionary<int, List<Connector>>();
            foreach (var line in connectorLines)
            {
                var row = line.Row - minRow;
                if (row < 0 || row >= rowCount) continue;
                if (!connectors.TryGetValue(row, out var list)) connectors[row] = list = [];
                var origin = SheetOffset(sheetColumnWidths, minColumn);
                list.Add(new Connector(row,
                    SheetOffset(sheetColumnWidths, line.FromCol) + line.FromOffsetPoints - origin,
                    SheetOffset(sheetColumnWidths, line.ToCol) + line.ToOffsetPoints - origin));
            }

            var grid = new SheetGrid
            {
                MinRow = minRow,
                MinColumn = minColumn,
                RowCount = rowCount,
                ColumnCount = columnCount,
                Cells = cells,
                RowHeights = rowHeights,
                ClipsToHeight = clipsToHeight,
                OriginalRowHeights = Enumerable.Range(0, rowCount)
                    .Select(r => rowHeights[r] ?? LayoutDefaults.ExcelRowHeightPoints).ToArray(),
                ColumnWidths = columnWidths,
                SheetColumnWidths = sheetColumnWidths,
                Merges = [],
                CoveredByMerge = [],
                SkipHorizontalMerge = [],
                Connectors = connectors,
                Images = images,
                Styles = styles,
            };

            grid.IndexMerges(mergeRanges);
            return grid;
        }

        void IndexMerges(List<(int startRow, int startCol, int endRow, int endCol)> ranges)
        {
            foreach (var range in ranges)
            {
                var merge = new MergeRange(range.startRow - MinRow, range.startCol - MinColumn,
                    range.endRow - MinRow, range.endCol - MinColumn);
                Merges.Add(merge);

                for (var column = merge.StartColumn + 1; column <= merge.EndColumn; column++)
                    if (merge.StartRow >= 0 && merge.StartRow < RowCount && column >= 0 && column < ColumnCount)
                        CoveredByMerge.Add((merge.StartRow, column));

                if (WouldExtendBordersFromRowAbove(merge)) SkipHorizontalMerge.Add((merge.StartRow, merge.StartColumn));
            }
        }

        // MigraDoc stretches a partial bottom border from the row above across a merged cell. When
        // the row above has mixed borders the merge is skipped and centring simulated with an indent.
        bool WouldExtendBordersFromRowAbove(MergeRange merge)
        {
            var span = Math.Min(merge.EndColumn, ColumnCount - 1) - merge.StartColumn;
            if (span <= 0 || merge.StartRow <= 0 || merge.StartRow >= RowCount) return false;
            if (merge.StartColumn < 0 || merge.StartColumn >= ColumnWidths.Count) return false;
            if (ColumnWidths[merge.StartColumn] < 40) return false;

            var anchor = Styles.GetCellStyle(CellAt(merge.StartRow, merge.StartColumn)?.StyleIndex?.Value);
            if (anchor.Borders.TopWidth > 0) return false;

            bool anyBorder = false, anyMissing = false;
            for (var column = merge.StartColumn; column <= merge.StartColumn + span; column++)
            {
                var above = Styles.GetCellStyle(CellAt(merge.StartRow - 1, column)?.StyleIndex?.Value);
                if (above.Borders.BottomWidth > 0) anyBorder = true;
                else anyMissing = true;
            }
            return anyBorder && anyMissing;
        }

        public MergeRange? MergeAnchoredAt(int row, int column)
        {
            foreach (var merge in Merges)
                if (merge.StartRow == row && merge.StartColumn == column) return merge;
            return null;
        }

        public double MergeWidth(MergeRange merge)
        {
            double width = 0;
            for (var column = merge.StartColumn; column <= merge.EndColumn && column < ColumnWidths.Count; column++)
                width += ColumnWidths[column];
            return width;
        }

        public void ScaleToWidth(double contentWidth)
        {
            var total = ColumnWidths.Sum();
            if (total <= contentWidth || total <= 0) return;

            var factor = contentWidth / total;
            ScaleFactor = factor;
            for (var i = 0; i < ColumnWidths.Count; i++) ColumnWidths[i] *= factor;
        }

        public void CollapseEmptyRows(ImagePlacement placement)
        {
            var mergedRows = new HashSet<int>();
            foreach (var merge in Merges)
                for (var row = merge.StartRow; row <= merge.EndRow; row++) mergedRows.Add(row);

            for (var row = 0; row < RowCount; row++)
            {
                if (RowHeights[row] != null) continue;
                if (mergedRows.Contains(row) || Connectors.ContainsKey(row) || placement.OccupiesRow(row)) continue;

                var hasData = false;
                for (var column = 0; column < ColumnCount && !hasData; column++)
                    hasData = Cells[row, column] != null;

                if (!hasData) RowHeights[row] = LayoutDefaults.ExcelEmptyRowHeightPoints;
            }
        }

        public double RowOffset(int row)
        {
            double offset = 0;
            for (var i = 0; i < row && i < RowCount; i++) offset += RowHeight(i);
            return offset;
        }
    }

    readonly record struct MergeRange(int StartRow, int StartColumn, int EndRow, int EndColumn);

    readonly record struct Connector(int Row, double StartPoints, double EndPoints);

    sealed class ImagePlacement
    {
        public Dictionary<int, List<(ExcelHelpers.ExcelImageInfo Info, string Path)>> ByRow { get; } = [];
        public List<(ExcelHelpers.ExcelImageInfo Info, string Path, int Row)> Floating { get; } = [];

        readonly HashSet<(int Row, string Path)> _floatingKeys = [];
        readonly HashSet<int> _occupiedRows = [];

        public bool OccupiesRow(int row) => _occupiedRows.Contains(row);

        public bool IsFloating(int row, string path) => _floatingKeys.Contains((row, path));

        public static ImagePlacement Build(SheetGrid sheet, TempImageStore images)
        {
            var placement = new ImagePlacement();
            var placed = new List<(ExcelHelpers.ExcelImageInfo Info, string Path, int Row)>();

            foreach (var info in sheet.Images)
            {
                var path = images.Save(info.Bytes);
                if (path == null) continue;
                var row = Math.Clamp((info.FromRow ?? 0) - sheet.MinRow, 0, sheet.RowCount - 1);
                placed.Add((info, path, row));
            }

            MergeOverlappingRows(placed);

            foreach (var (info, path, row) in placed)
            {
                if (!placement.ByRow.TryGetValue(row, out var list)) placement.ByRow[row] = list = [];
                list.Add((info, path));
                placement._occupiedRows.Add(row);

                if (!IsInSpacerColumn(info, sheet)) continue;
                for (var spanned = info.FromRow ?? 0; spanned < (info.ToRow ?? 0); spanned++)
                {
                    var offset = spanned - sheet.MinRow;
                    if (offset >= 0 && offset < sheet.RowCount) placement._occupiedRows.Add(offset);
                }
            }

            // A picture anchored left of the data shares its row with text; MigraDoc cannot let a
            // cell picture overflow, so those are positioned absolutely instead.
            foreach (var (row, entries) in placement.ByRow)
            {
                if (!sheet.RowHasText(row)) continue;
                foreach (var (info, path) in entries)
                {
                    if (!IsInSpacerColumn(info, sheet)) continue;
                    placement.Floating.Add((info, path, row));
                    placement._floatingKeys.Add((row, path));
                }
            }

            return placement;
        }

        static bool IsInSpacerColumn(ExcelHelpers.ExcelImageInfo info, SheetGrid sheet) =>
            (info.FromCol ?? 0) < sheet.MinColumn;

        static void MergeOverlappingRows(List<(ExcelHelpers.ExcelImageInfo Info, string Path, int Row)> placed)
        {
            for (var i = 0; i < placed.Count; i++)
            {
                for (var j = i + 1; j < placed.Count; j++)
                {
                    var (first, second) = (placed[i], placed[j]);
                    if (first.Info.FromCol == second.Info.FromCol) continue;

                    var firstEnd = first.Info.ToRow ?? (first.Info.FromRow ?? 0) + 1;
                    var secondEnd = second.Info.ToRow ?? (second.Info.FromRow ?? 0) + 1;
                    if ((first.Info.FromRow ?? 0) >= secondEnd || (second.Info.FromRow ?? 0) >= firstEnd) continue;

                    var row = Math.Min(first.Row, second.Row);
                    placed[i] = first with { Row = row };
                    placed[j] = second with { Row = row };
                }
            }
        }
    }

    static void RenderRow(Table table, SheetGrid sheet, ExcelStyles styles, ImagePlacement placement, int rowIndex)
    {
        var row = table.AddRow();
        row.Height = Unit.FromPoint(sheet.RowHeight(rowIndex));
        // Excel clips a row given an explicit height; MigraDoc would grow it to fit instead.
        if (sheet.ClipsToHeight[rowIndex]) row.HeightRule = RowHeightRule.Exactly;

        if (sheet.Connectors.TryGetValue(rowIndex, out var connectors))
            foreach (var connector in connectors)
                AddConnectorLine(sheet, row, connector);

        for (var column = 0; column < sheet.ColumnCount; column++)
        {
            var cell = sheet.CellAt(rowIndex, column);
            var target = row.Cells[column];

            if (sheet.CoveredByMerge.Contains((rowIndex, column)))
            {
                if (cell != null) ApplyCellBorders(target, styles.GetCellStyle(cell.StyleIndex?.Value).Borders);
                continue;
            }

            var style = styles.GetCellStyle(cell?.StyleIndex?.Value);
            var text = cell == null ? string.Empty : ReadCellValue(cell, styles);

            ApplyCellStyle(target, style);
            ApplyHorizontalMerge(sheet, target, rowIndex, column);
            AddCellImages(sheet, placement, target, rowIndex, column, text.Length > 0);
            AddCellText(sheet, target, style, text, rowIndex, column);
        }

    }

    // A negative right indent stretches the rule past its starting cell, because two connectors
    // often share a column at different offsets, which merged cells cannot express.
    static void AddConnectorLine(SheetGrid sheet, Row row, Connector connector)
    {
        var start = Math.Max(connector.StartPoints * sheet.ScaleFactor, 0);
        var end = Math.Min(connector.EndPoints * sheet.ScaleFactor, sheet.TableWidth);
        if (end <= start) return;

        var host = sheet.ColumnAt(start);
        var indent = start - sheet.ColumnOffset(host);

        var paragraph = row.Cells[host].AddParagraph();
        paragraph.Format.Font.Size = 1;
        paragraph.Format.SpaceBefore = 0;
        paragraph.Format.SpaceAfter = 0;
        paragraph.Format.LeftIndent = Unit.FromPoint(Math.Max(indent, 0));
        paragraph.Format.RightIndent = Unit.FromPoint(sheet.ColumnWidths[host] - indent - (end - start));
        paragraph.Format.Borders.Bottom.Width = Unit.FromPoint(LayoutDefaults.BorderWidthPoints);
        paragraph.Format.Borders.Bottom.Color = Colors.Black;
    }

    static void ApplyCellStyle(Cell target, ExcelCellStyleInfo style)
    {
        if (ColorUtils.TryParse(style.FillColor, out var fill)) target.Shading.Color = fill;

        if (!string.IsNullOrEmpty(style.HorizontalAlignment))
            target.Format.Alignment = style.HorizontalAlignment.ToLowerInvariant() switch
            {
                "center" or "centercontinuous" => ParagraphAlignment.Center,
                "right" => ParagraphAlignment.Right,
                "justify" or "distributed" => ParagraphAlignment.Justify,
                _ => ParagraphAlignment.Left,
            };

        if (!string.IsNullOrEmpty(style.VerticalAlignment))
            target.VerticalAlignment = style.VerticalAlignment.ToLowerInvariant() switch
            {
                "center" => VerticalAlignment.Center,
                "bottom" => VerticalAlignment.Bottom,
                _ => VerticalAlignment.Top,
            };

        ApplyCellBorders(target, style.Borders);
    }

    static void ApplyHorizontalMerge(SheetGrid sheet, Cell target, int row, int column)
    {
        if (sheet.MergeAnchoredAt(row, column) is not { } merge) return;

        var span = Math.Min(merge.EndColumn - merge.StartColumn, sheet.ColumnCount - 1 - column);
        if (span <= 0) return;

        if (!sheet.SkipHorizontalMerge.Contains((row, column)))
        {
            target.MergeRight = span;
            return;
        }

        var mergeWidth = sheet.MergeWidth(merge);
        var anchorWidth = sheet.ColumnWidths[column];
        if (target.Format.Alignment == ParagraphAlignment.Center && mergeWidth > anchorWidth)
            target.Format.LeftIndent = Unit.FromPoint(mergeWidth - anchorWidth);
    }

    static void AddCellImages(SheetGrid sheet, ImagePlacement placement, Cell target,
        int row, int column, bool cellHasText)
    {
        if (!placement.ByRow.TryGetValue(row, out var entries)) return;

        var actualColumn = column + sheet.MinColumn;
        var merge = sheet.MergeAnchoredAt(row, column);
        var mergeEndColumn = merge is { } m && m.EndColumn > m.StartColumn
            ? m.EndColumn + sheet.MinColumn
            : actualColumn;

        foreach (var (info, path) in entries)
        {
            if (placement.IsFloating(row, path)) continue;

            var imageColumn = info.FromCol ?? 0;
            var fromSpacer = imageColumn < sheet.MinColumn && column == 0;
            var matches = imageColumn == actualColumn || fromSpacer
                || (imageColumn > actualColumn && imageColumn <= mergeEndColumn);
            if (!matches) continue;

            var start = fromSpacer
                ? 0
                : sheet.SheetStart(imageColumn, info.FromColumnOffsetPoints) * sheet.ScaleFactor;
            var indent = Math.Max(start - sheet.ColumnOffset(column), 0);
            var maxWidth = Math.Max(10, sheet.TableWidth - Math.Max(start, 0));
            var maxHeight = cellHasText ? sheet.RowHeight(row) : SpannedHeight(sheet, info, row);

            var paragraph = target.AddParagraph();
            if (indent > 0) paragraph.Format.LeftIndent = Unit.FromPoint(indent);

            var image = paragraph.AddImage(path);
            image.LockAspectRatio = false;
            image.Width = Unit.FromPoint(Math.Min(ImageWidth(sheet, info, column), maxWidth));
            image.Height = Unit.FromPoint(Math.Min(ImageHeight(info, maxHeight), maxHeight));
        }
    }

    static double ImageWidth(SheetGrid sheet, ExcelHelpers.ExcelImageInfo info, int column)
    {
        if (info.WidthEmu is > 0) return Units.EmuToPoints(info.WidthEmu.Value);

        if (info.FromCol is { } from && info.ToCol is { } to)
        {
            double width = 0;
            for (var i = from; i < to && i < sheet.SheetColumnWidths.Count; i++) width += sheet.SheetColumnWidths[i];
            if (width > 0) return width;
        }

        return column < sheet.ColumnWidths.Count
            ? Math.Min(sheet.ColumnWidths[column], Unit.FromCentimeter(LayoutDefaults.ExcelImageWidthCentimeters).Point)
            : LayoutDefaults.ExcelColumnWidthPoints;
    }

    static double ImageHeight(ExcelHelpers.ExcelImageInfo info, double fallback) =>
        info.HeightEmu is > 0 ? Units.EmuToPoints(info.HeightEmu.Value) : fallback;

    static double SpannedHeight(SheetGrid sheet, ExcelHelpers.ExcelImageInfo info, int row)
    {
        if (info.FromRow is not { } from || info.ToRow is not { } to || to <= from) return sheet.RowHeight(row);

        double height = 0;
        for (var i = from; i < to; i++)
        {
            var offset = i - sheet.MinRow;
            height += offset >= 0 && offset < sheet.RowCount
                ? sheet.OriginalRowHeights[offset]
                : LayoutDefaults.ExcelRowHeightPoints;
        }
        return height > 0 ? height : sheet.RowHeight(row);
    }

    static void AddCellText(SheetGrid sheet, Cell target, ExcelCellStyleInfo style, string text,
        int row, int column)
    {
        var fontSize = style.FontSize ?? LayoutDefaults.ExcelFontSizePoints;
        if (sheet.HidesText(row, fontSize)) text = string.Empty;

        var paragraph = target.AddParagraph();
        paragraph.Format.Alignment = target.Format.Alignment;

        if (text.Length == 0)
        {
            paragraph.Format.Font.Size = 1;
            paragraph.Format.SpaceBefore = 0;
            paragraph.Format.SpaceAfter = 0;
            paragraph.Format.LineSpacing = Unit.FromPoint(1);
            return;
        }

        var formatted = paragraph.AddFormattedText(text);
        formatted.Size = fontSize;
        if (!string.IsNullOrEmpty(style.FontFamily)) formatted.Font.Name = style.FontFamily;
        if (ColorUtils.TryParse(style.FontColor, out var color)) formatted.Color = color;
        if (style.Bold) formatted.Bold = true;
        if (style.Italic) formatted.Italic = true;
    }

    static void RenderFloatingImages(Section section, SheetGrid sheet, ImagePlacement placement, double tableIndent)
    {
        foreach (var (info, path, row) in placement.Floating)
        {
            try
            {
                var image = section.AddImage(path);
                image.LockAspectRatio = false;
                image.Width = Unit.FromPoint(ImageWidth(sheet, info, 0));
                image.Height = Unit.FromPoint(ImageHeight(info, SpannedHeight(sheet, info, row)));
                image.RelativeHorizontal = RelativeHorizontal.Page;
                image.RelativeVertical = RelativeVertical.Page;
                image.Left = Unit.FromPoint(section.PageSetup.LeftMargin.Point + tableIndent);
                image.Top = Unit.FromPoint(section.PageSetup.TopMargin.Point + sheet.RowOffset(row));
                image.WrapFormat.Style = WrapStyle.None;
            }
            catch (Exception ex)
            {
                OpenXmlHelpers.ImageLoadLogger?.Invoke($"Failed placing floating image '{path}': {ex.Message}");
            }
        }
    }

    static void ApplyCellBorders(Cell target, BorderInfo borders)
    {
        ApplyEdge(target.Borders.Top, borders.TopWidth, borders.TopColor);
        ApplyEdge(target.Borders.Bottom, borders.BottomWidth, borders.BottomColor);
        ApplyEdge(target.Borders.Left, borders.LeftWidth, borders.LeftColor);
        ApplyEdge(target.Borders.Right, borders.RightWidth, borders.RightColor);
    }

    static void ApplyEdge(Border border, double width, string? colorValue)
    {
        if (width <= 0) return;
        border.Width = Unit.FromPoint(width);
        if (ColorUtils.TryParse(colorValue, out var color)) border.Color = color;
    }

    static string ReadCellValue(S.Cell cell, ExcelStyles styles)
    {
        var dataType = cell.DataType?.Value;

        if (dataType == S.CellValues.InlineString)
            return cell.InlineString?.Text?.Text ?? cell.InnerText;

        if (dataType == S.CellValues.SharedString)
            return ResolveSharedString(cell);

        if (dataType == S.CellValues.Boolean)
            return (cell.CellValue?.Text ?? cell.InnerText) == "0" ? "FALSE" : "TRUE";

        if (dataType == S.CellValues.Error)
            return cell.CellValue?.Text ?? cell.InnerText;

        if (dataType == S.CellValues.String)
            return cell.CellValue?.Text ?? cell.InnerText;

        if (dataType == S.CellValues.Date)
            return cell.CellValue?.Text ?? cell.InnerText;

        var raw = cell.CellValue?.Text;
        if (string.IsNullOrEmpty(raw)) return cell.CellFormula != null ? string.Empty : cell.InnerText;

        var numberFormat = styles.GetNumberFormat(styles.GetCellStyle(cell.StyleIndex?.Value).NumberFormatId);
        if (!string.IsNullOrEmpty(numberFormat)) return ExcelNumberFormat.Apply(raw, numberFormat);

        return Units.TryParseDouble(raw, out var number) ? GeneralFormat(number) : raw;
    }

    // Excel's "General" format carries at most 11 significant digits, so a value stored as
    // 4.099999999998545 displays as 4.1 rather than exposing its binary-rounding tail.
    const int GeneralSignificantDigits = 11;

    static string GeneralFormat(double value)
    {
        var culture = System.Globalization.CultureInfo.CurrentCulture;
        if (value == 0 || double.IsNaN(value) || double.IsInfinity(value))
            return value.ToString(culture);

        var decimals = GeneralSignificantDigits - 1 - (int)Math.Floor(Math.Log10(Math.Abs(value)));
        if (decimals is < 0 or > 15) return value.ToString("G15", culture);

        return Math.Round(value, decimals).ToString("G15", culture);
    }

    static string ResolveSharedString(S.Cell cell)
    {
        var raw = cell.CellValue?.Text ?? cell.InnerText;
        if (!Units.TryParseInt(raw, out var index)) return raw;

        var workbookPart = cell.Ancestors<S.Worksheet>().FirstOrDefault()?.WorksheetPart
            ?.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
        var table = workbookPart?.SharedStringTablePart?.SharedStringTable;
        if (table == null) return raw;

        var item = table.Elements<S.SharedStringItem>().ElementAtOrDefault(index);
        return item?.InnerText ?? raw;
    }
}
