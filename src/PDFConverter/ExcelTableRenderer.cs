using DocumentFormat.OpenXml.Packaging;
using PDFConverter.Layout;
using SkiaSharp;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace PDFConverter;

internal static class ExcelTableRenderer
{
    const string DefaultFontFamily = "Calibri";

    public static void RenderSheet(Section section, WorksheetPart wsPart, ImageStore images)
    {
        var workbookPart = wsPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
        if (workbookPart == null) return;

        var styles = ExcelStyles.For(workbookPart);
        var sheet = SheetGrid.Build(wsPart, styles);
        if (sheet == null) return;

        var contentWidth = section.ContentWidth;
        sheet.ScaleToWidth(contentWidth);

        var placement = ImagePlacement.Build(sheet, images);
        sheet.CollapseEmptyRows(placement);

        var table = section.AddTable();
        table.ColumnWidths.AddRange(sheet.ColumnWidths);
        table.LeftIndent = Math.Max(0, (contentWidth - sheet.TableWidth) / 2);

        for (var row = 0; row < sheet.RowCount; row++)
            RenderRow(table, sheet, styles, placement, row);
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

        public bool HidesText(int row, double fontSize) =>
            ClipsToHeight[row] && RowHeight(row) < fontSize * LayoutDefaults.ExcelLineHeightRatio;

        public S.Cell? CellAt(int row, int column) =>
            row >= 0 && row < RowCount && column >= 0 && column < ColumnCount ? Cells[row, column] : null;

        bool IsFree(int row, int column)
        {
            if (column < 0 || column >= ColumnCount || CoveredByMerge.Contains((row, column))) return false;
            var cell = CellAt(row, column);
            return cell == null || ReadCellValue(cell, Styles).Length == 0;
        }

        public double FreeWidthRight(int row, int column)
        {
            double width = 0;
            for (var c = column; IsFree(row, c); c++) width += ColumnWidths[c];
            return width;
        }

        public double FreeWidthLeft(int row, int column)
        {
            double width = 0;
            for (var c = column; IsFree(row, c); c--) width += ColumnWidths[c];
            return width;
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
                    SheetOffset(sheetColumnWidths, line.ToCol) + line.ToOffsetPoints - origin,
                    line.RowOffsetPoints));
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

                for (var row = merge.StartRow; row <= merge.EndRow; row++)
                    for (var column = merge.StartColumn; column <= merge.EndColumn; column++)
                        if ((row != merge.StartRow || column != merge.StartColumn)
                            && row >= 0 && row < RowCount && column >= 0 && column < ColumnCount)
                            CoveredByMerge.Add((row, column));
            }
        }

        public MergeRange? MergeAnchoredAt(int row, int column)
        {
            foreach (var merge in Merges)
                if (merge.StartRow == row && merge.StartColumn == column) return merge;
            return null;
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
    }

    readonly record struct MergeRange(int StartRow, int StartColumn, int EndRow, int EndColumn);

    readonly record struct Connector(int Row, double StartPoints, double EndPoints, double OffsetPoints);

    sealed class ImagePlacement
    {
        public Dictionary<int, List<(ExcelHelpers.ExcelImageInfo Info, ImageSource Image)>> ByRow { get; } = [];

        readonly HashSet<int> _occupiedRows = [];

        public bool OccupiesRow(int row) => _occupiedRows.Contains(row);

        public static ImagePlacement Build(SheetGrid sheet, ImageStore images)
        {
            var placement = new ImagePlacement();
            var placed = new List<(ExcelHelpers.ExcelImageInfo Info, ImageSource Image, int Row)>();

            foreach (var info in sheet.Images)
            {
                var image = images.Get(info.Bytes);
                if (image == null) continue;
                var row = Math.Clamp((info.FromRow ?? 0) - sheet.MinRow, 0, sheet.RowCount - 1);
                placed.Add((info, image, row));
            }

            MergeOverlappingRows(placed);

            foreach (var (info, image, row) in placed)
            {
                if (!placement.ByRow.TryGetValue(row, out var list)) placement.ByRow[row] = list = [];
                list.Add((info, image));
                placement._occupiedRows.Add(row);

                for (var spanned = info.FromRow ?? 0; spanned <= (info.ToRow ?? 0); spanned++)
                {
                    var offset = spanned - sheet.MinRow;
                    if (offset >= 0 && offset < sheet.RowCount) placement._occupiedRows.Add(offset);
                }
            }

            return placement;
        }

        public static bool IsInSpacerColumn(ExcelHelpers.ExcelImageInfo info, SheetGrid sheet) =>
            (info.FromCol ?? 0) < sheet.MinColumn;

        static void MergeOverlappingRows(List<(ExcelHelpers.ExcelImageInfo Info, ImageSource Image, int Row)> placed)
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
        row.MinHeight = sheet.RowHeight(rowIndex);
        row.ExactHeight = sheet.ClipsToHeight[rowIndex];

        if (sheet.Connectors.TryGetValue(rowIndex, out var connectors))
            foreach (var connector in connectors)
                AddConnectorLine(sheet, row, connector, rowIndex);

        for (var column = 0; column < sheet.ColumnCount; column++)
        {
            var cell = sheet.CellAt(rowIndex, column);
            var style = styles.GetCellStyle(cell?.StyleIndex?.Value);

            if (sheet.CoveredByMerge.Contains((rowIndex, column)))
            {
                if (cell != null) ApplyCellBorders(row.Cell(column), style.Borders);
                continue;
            }

            var target = row.Cell(column);
            target.BordersSpanColumns = false;
            ApplyCellStyle(target, style);
            ApplyHorizontalMerge(sheet, target, rowIndex, column);

            var text = cell == null ? string.Empty : ReadCellValue(cell, styles);
            AddCellText(sheet, target, style, text, rowIndex, column, IsNumeric(cell));
        }

        AddRowImages(sheet, placement, row, rowIndex);
    }

    static void AddConnectorLine(SheetGrid sheet, TableRow row, Connector connector, int rowIndex)
    {
        var start = Math.Max(connector.StartPoints * sheet.ScaleFactor, 0);
        var end = Math.Min(connector.EndPoints * sheet.ScaleFactor, sheet.TableWidth);
        if (end <= start) return;

        var offset = Math.Clamp(connector.OffsetPoints, 0, sheet.RowHeight(rowIndex));
        row.Lines.Add(new RowLine(start, end, offset, SKColors.Black, LayoutDefaults.BorderWidthPoints));
    }

    static void ApplyCellStyle(TableCell target, ExcelCellStyleInfo style)
    {
        if (ColorUtils.TryParse(style.FillColor, out var fill)) target.Shading = fill;

        target.VerticalAlignment = style.VerticalAlignment?.ToLowerInvariant() switch
        {
            "center" => VerticalAlignment.Center,
            "top" => VerticalAlignment.Top,
            _ => VerticalAlignment.Bottom,
        };

        ApplyCellBorders(target, style.Borders);
    }

    static ParagraphAlignment Alignment(ExcelCellStyleInfo style, bool numeric) =>
        style.HorizontalAlignment?.ToLowerInvariant() switch
        {
            "center" or "centercontinuous" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            "justify" or "distributed" => ParagraphAlignment.Justify,
            "left" or "fill" => ParagraphAlignment.Left,
            _ => numeric ? ParagraphAlignment.Right : ParagraphAlignment.Left,
        };

    static bool IsNumeric(S.Cell? cell)
    {
        if (cell == null) return false;
        var type = cell.DataType?.Value;
        if (type != null && type != S.CellValues.Number) return false;
        return Units.TryParseDouble(cell.CellValue?.Text, out _);
    }

    static void ApplyHorizontalMerge(SheetGrid sheet, TableCell target, int row, int column)
    {
        if (sheet.MergeAnchoredAt(row, column) is not { } merge) return;

        var span = Math.Min(merge.EndColumn - merge.StartColumn, sheet.ColumnCount - 1 - column);
        if (span > 0) target.ColumnSpan = span + 1;
        var rowSpan = Math.Min(merge.EndRow - merge.StartRow, sheet.RowCount - 1 - row);
        if (rowSpan > 0) target.RowSpan = rowSpan + 1;
    }

    static void AddRowImages(SheetGrid sheet, ImagePlacement placement, TableRow row, int rowIndex)
    {
        if (!placement.ByRow.TryGetValue(rowIndex, out var entries)) return;

        foreach (var (info, image) in entries)
        {
            var start = ImagePlacement.IsInSpacerColumn(info, sheet)
                ? 0
                : Math.Max(sheet.SheetStart(info.FromCol ?? 0, info.FromColumnOffsetPoints) * sheet.ScaleFactor, 0);
            var maxWidth = Math.Max(10, sheet.TableWidth - start);
            var column = Math.Max(0, Math.Min(sheet.ColumnCount - 1, (info.FromCol ?? 0) - sheet.MinColumn));

            var width = Math.Min(ImageWidth(sheet, info, column), maxWidth);
            var height = info.HeightEmu is > 0 ? Units.EmuToPoints(info.HeightEmu.Value) : SpannedHeight(sheet, info, rowIndex);
            row.Images.Add(new RowImage(image, start, info.FromRowOffsetPoints, width, height));
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
            ? Math.Min(sheet.ColumnWidths[column], LayoutDefaults.ExcelImageWidthCentimeters / 2.54 * 72)
            : LayoutDefaults.ExcelColumnWidthPoints;
    }

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

    static void AddCellText(SheetGrid sheet, TableCell target, ExcelCellStyleInfo style, string text, int row, int column,
        bool numeric)
    {
        var fontSize = style.FontSize ?? LayoutDefaults.ExcelFontSizePoints;
        if (text.Length == 0 || sheet.HidesText(row, fontSize)) return;

        var paragraph = target.AddParagraph();
        paragraph.Alignment = Alignment(style, numeric);

        // Excel lets unwrapped text run into neighbouring empty cells; when there is no room to
        // overflow it would truncate, and wrapping reads better than a cut-off value.
        if (!style.WrapText)
        {
            var right = paragraph.Alignment == ParagraphAlignment.Right
                ? 0
                : sheet.FreeWidthRight(row, column + target.ColumnSpan);
            var left = paragraph.Alignment is ParagraphAlignment.Right or ParagraphAlignment.Center
                ? sheet.FreeWidthLeft(row, column - 1)
                : 0;
            if (left + right > 0)
            {
                paragraph.NoWrap = true;
                target.ClipContent = true;
                target.OverflowLeft = left;
                target.OverflowRight = right;
            }
        }
        paragraph.DefaultStyle = new TextStyle(DefaultFontFamily, fontSize);

        var textStyle = new TextStyle(
            string.IsNullOrEmpty(style.FontFamily) ? DefaultFontFamily : style.FontFamily,
            fontSize,
            style.Bold,
            style.Italic,
            ColorUtils.TryParse(style.FontColor, out var color) ? color : null);
        paragraph.Add(new TextRun(text, textStyle));
    }

    static void ApplyCellBorders(TableCell target, BorderInfo borders)
    {
        target.Borders.Top = Edge(borders.TopWidth, borders.TopColor, borders.TopStyle);
        target.Borders.Bottom = Edge(borders.BottomWidth, borders.BottomColor, borders.BottomStyle);
        target.Borders.Left = Edge(borders.LeftWidth, borders.LeftColor, borders.LeftStyle);
        target.Borders.Right = Edge(borders.RightWidth, borders.RightColor, borders.RightStyle);
    }

    static BorderEdge? Edge(double width, string? colorValue, string? style)
    {
        if (width <= 0) return null;
        return new BorderEdge(width, ColorUtils.TryParse(colorValue, out var color) ? color : SKColors.Black,
            BorderPatterns.From(style));
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
