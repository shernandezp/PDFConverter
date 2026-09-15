using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal static class WordTableRenderer
{
    public static void RenderTable(WordRenderContext ctx, Section section, W.Table table)
    {
        var rows = table.Elements<W.TableRow>().ToList();
        if (rows.Count == 0) return;

        var contentWidth = section.PageSetup.PageWidth.Point
            - section.PageSetup.LeftMargin.Point - section.PageSetup.RightMargin.Point;

        var layout = TableLayout.Measure(table, rows, contentWidth);
        var migTable = section.AddTable();
        foreach (var width in layout.ScaledWidths)
            migTable.AddColumn(Unit.FromPoint(width));

        var tblPr = table.GetFirstChild<W.TableProperties>();
        ApplyTableBorders(migTable, WordHelpers.ResolveTableBorders(ctx.MainPart, tblPr));
        var margins = ApplyCellMargins(migTable, ctx, tblPr);
        ApplyTableIndent(migTable, tblPr, layout, contentWidth);

        var firstRowStyle = WordHelpers.GetTableStyleConditionalFormatting(
            ctx.MainPart, tblPr, W.TableStyleOverrideValues.FirstRow);
        var firstColumnStyle = WordHelpers.GetTableStyleConditionalFormatting(
            ctx.MainPart, tblPr, W.TableStyleOverrideValues.FirstColumn);

        var verticalSpans = MeasureVerticalMerges(rows, layout.ColumnCount);

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var sourceRow = rows[rowIndex];
            var row = migTable.AddRow();
            ApplyRowProperties(row, sourceRow);

            var isFirstRow = WordHelpers.IsConditionalRow(
                sourceRow, tblPr, rowIndex, rows.Count, W.TableStyleOverrideValues.FirstRow);

            var column = 0;
            foreach (var sourceCell in sourceRow.Elements<W.TableCell>())
            {
                if (column >= layout.ColumnCount) break;

                var tcPr = sourceCell.GetFirstChild<W.TableCellProperties>();
                var gridSpan = (int)(tcPr?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1);
                var verticalMerge = tcPr?.GetFirstChild<W.VerticalMerge>();

                if (IsMergeContinuation(verticalMerge))
                {
                    column += gridSpan;
                    continue;
                }

                var cell = row.Cells[column];
                var conditionalStyle = isFirstRow ? firstRowStyle : column == 0 ? firstColumnStyle : null;

                ApplyShading(cell, tcPr, conditionalStyle);
                var borders = WordHelpers.GetWordCellBorders(tcPr, tblPr);
                ApplyCellBorders(cell, borders);
                ApplyCellPadding(cell, borders);
                ApplyVerticalAlignment(cell, tcPr);

                if (gridSpan > 1)
                {
                    var mergeRight = Math.Min(gridSpan - 1, layout.ColumnCount - column - 1);
                    if (mergeRight > 0) cell.MergeRight = mergeRight;
                }
                if (verticalSpans.TryGetValue((rowIndex, column), out var mergeDown) && mergeDown > 0)
                    cell.MergeDown = mergeDown;

                var textWidth = layout.WidthOf(column, gridSpan) - margins.Left - margins.Right;
                RenderCellContent(ctx, cell, sourceCell, textWidth, conditionalStyle);
                column += gridSpan;
            }

            if (column == 0) RenderDrawingMlRow(ctx, sourceRow, row, layout.ColumnCount);
        }
    }

    const string DrawingMlNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const string WordprocessingMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Word occasionally emits DrawingML cells (a:tc) inside an ordinary w:tr; they carry text and
    // hyperlinks but none of the WordprocessingML cell structure.
    static void RenderDrawingMlRow(WordRenderContext ctx, W.TableRow sourceRow, Row row, int columnCount)
    {
        var column = 0;
        foreach (var sourceCell in sourceRow.ChildElements)
        {
            if (sourceCell.LocalName != "tc" || sourceCell.NamespaceUri != DrawingMlNamespace) continue;
            if (column >= columnCount) break;

            var target = row.Cells[column].AddParagraph();
            foreach (var paragraph in sourceCell.ChildElements.Where(c => c.LocalName == "p"))
            {
                foreach (var child in paragraph.ChildElements)
                {
                    if (child.LocalName == "hyperlink" && child.NamespaceUri == WordprocessingMlNamespace)
                        AddDrawingMlHyperlink(ctx, child, target);
                    else if (child.LocalName == "r" && child.NamespaceUri == DrawingMlNamespace)
                        target.AddText(DrawingMlText(child));
                }
            }
            column++;
        }
    }

    static void AddDrawingMlHyperlink(WordRenderContext ctx, DocumentFormat.OpenXml.OpenXmlElement source, Paragraph target)
    {
        const string relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var relationshipId = source.GetAttributes()
            .FirstOrDefault(a => a.LocalName == "id" && a.NamespaceUri == relationships).Value;

        string? url = null;
        if (!string.IsNullOrEmpty(relationshipId))
        {
            try
            {
                url = ctx.MainPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relationshipId)?.Uri?.ToString();
            }
            catch (Exception ex)
            {
                OpenXmlHelpers.ImageLoadLogger?.Invoke($"Unresolved hyperlink '{relationshipId}': {ex.Message}");
            }
        }

        var text = DrawingMlText(source);
        if (string.IsNullOrEmpty(text)) text = url ?? string.Empty;
        if (string.IsNullOrEmpty(text)) return;

        var formatted = url == null
            ? target.AddFormattedText(text)
            : target.AddHyperlink(url, HyperlinkType.Web).AddFormattedText(text);
        formatted.Color = Colors.Blue;
        formatted.Underline = Underline.Single;
    }

    static string DrawingMlText(DocumentFormat.OpenXml.OpenXmlElement source) =>
        string.Concat(source.Descendants().Where(d => d.LocalName == "t").Select(d => d.InnerText));

    public static void RenderNestedTable(WordRenderContext ctx, W.Table table, Cell parentCell)
    {
        var rows = table.Elements<W.TableRow>().ToList();
        if (rows.Count == 0) return;

        var parentWidth = parentCell.Column?.Width.Point ?? 0;
        var layout = TableLayout.Measure(table, rows, parentWidth);

        var nested = new Table();
        foreach (var width in layout.ScaledWidths)
            nested.AddColumn(Unit.FromPoint(width));

        var tblPr = table.GetFirstChild<W.TableProperties>();
        ApplyTableBorders(nested, WordHelpers.ResolveTableBorders(ctx.MainPart, tblPr));
        var margins = ApplyCellMargins(nested, ctx, tblPr);

        foreach (var sourceRow in rows)
        {
            var row = nested.AddRow();
            var column = 0;
            foreach (var sourceCell in sourceRow.Elements<W.TableCell>())
            {
                if (column >= layout.ColumnCount) break;
                var tcPr = sourceCell.GetFirstChild<W.TableCellProperties>();
                ApplyShading(row.Cells[column], tcPr, null);
                RenderCellContent(ctx, row.Cells[column], sourceCell,
                    layout.WidthOf(column, 1) - margins.Left - margins.Right, null);
                column += (int)(tcPr?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1);
            }
        }

        parentCell.Elements.Add(nested);
    }

    sealed class TableLayout
    {
        public required IReadOnlyList<double> ScaledWidths { get; init; }
        public int ColumnCount => ScaledWidths.Count;
        public double TotalWidth => ScaledWidths.Sum();

        public double WidthOf(int column, int span)
        {
            double width = 0;
            for (var i = column; i < Math.Min(column + span, ScaledWidths.Count); i++)
                width += ScaledWidths[i];
            return width;
        }

        public static TableLayout Measure(W.Table table, List<W.TableRow> rows, double availableWidth)
        {
            var columnCount = rows
                .Select(row => row.Elements<W.TableCell>()
                    .Sum(cell => (int)(cell.GetFirstChild<W.TableCellProperties>()
                        ?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1)))
                .DefaultIfEmpty(0)
                .Max();
            if (columnCount == 0) columnCount = 1;

            var widths = WordHelpers.GetTableGridColumnWidths(table, availableWidth);
            if (widths.Count > columnCount) widths.RemoveRange(columnCount, widths.Count - columnCount);

            var fallback = availableWidth > 0
                ? availableWidth / columnCount
                : LayoutDefaults.TableColumnWidthPoints;
            while (widths.Count < columnCount) widths.Add(fallback);

            var total = widths.Sum();
            var target = WordHelpers.GetTableWidth(table.GetFirstChild<W.TableProperties>(), availableWidth);
            if (target is > 0 && total > 0) total = Rescale(widths, target.Value / total);
            if (availableWidth > 0 && total > availableWidth) Rescale(widths, availableWidth / total);

            return new TableLayout { ScaledWidths = widths };
        }

        static double Rescale(List<double> widths, double factor)
        {
            for (var i = 0; i < widths.Count; i++) widths[i] *= factor;
            return widths.Sum();
        }
    }

    static Dictionary<(int Row, int Column), int> MeasureVerticalMerges(List<W.TableRow> rows, int columnCount)
    {
        var spans = new Dictionary<(int, int), int>();
        var cellsPerRow = rows.Select(row => row.Elements<W.TableCell>().ToList()).ToList();

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var column = 0;
            foreach (var cell in cellsPerRow[rowIndex])
            {
                var tcPr = cell.GetFirstChild<W.TableCellProperties>();
                var gridSpan = (int)(tcPr?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1);
                var verticalMerge = tcPr?.GetFirstChild<W.VerticalMerge>();

                if (verticalMerge?.Val?.Value == W.MergedCellValues.Restart && column < columnCount)
                {
                    var span = 0;
                    for (var below = rowIndex + 1; below < rows.Count; below++)
                    {
                        var belowCell = CellAtColumn(cellsPerRow[below], column);
                        var belowMerge = belowCell?.GetFirstChild<W.TableCellProperties>()
                            ?.GetFirstChild<W.VerticalMerge>();
                        if (!IsMergeContinuation(belowMerge)) break;
                        span++;
                    }
                    if (span > 0) spans[(rowIndex, column)] = span;
                }

                column += gridSpan;
            }
        }
        return spans;
    }

    static bool IsMergeContinuation(W.VerticalMerge? verticalMerge) =>
        verticalMerge != null
        && (verticalMerge.Val == null || verticalMerge.Val.Value == W.MergedCellValues.Continue);

    static W.TableCell? CellAtColumn(List<W.TableCell> cells, int targetColumn)
    {
        var column = 0;
        foreach (var cell in cells)
        {
            if (column == targetColumn) return cell;
            column += (int)(cell.GetFirstChild<W.TableCellProperties>()
                ?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1);
            if (column > targetColumn) return null;
        }
        return null;
    }

    static void RenderCellContent(WordRenderContext ctx, Cell cell, W.TableCell sourceCell,
        double cellWidth, W.TableStyleProperties? conditionalStyle)
    {
        var conditionalRun = conditionalStyle?.RunPropertiesBaseStyle;
        var options = new InlineOptions(
            MaxImageWidthPoints: cellWidth * 0.9,
            ConditionalBold: conditionalRun?.Bold != null
                && (conditionalRun.Bold.Val == null || conditionalRun.Bold.Val.Value),
            ConditionalColor: conditionalRun?.Color?.Val?.Value,
            MaxTextWidthPoints: cellWidth);

        var paragraphs = new List<W.Paragraph>();
        foreach (var child in sourceCell.ChildElements)
        {
            if (child is W.Paragraph paragraph) paragraphs.Add(paragraph);
            else if (child is W.SdtBlock { SdtContentBlock: not null } sdt)
                paragraphs.AddRange(sdt.SdtContentBlock.Elements<W.Paragraph>());
        }

        if (paragraphs.Count == 0)
        {
            cell.AddParagraph();
            return;
        }

        for (var i = 0; i < paragraphs.Count; i++)
        {
            var source = paragraphs[i];
            var target = cell.AddParagraph();
            var format = WordHelpers.GetParagraphFormatting(source.ParagraphProperties);

            target.Format.Alignment = format.Alignment;
            if (format.LeftIndent > 0) target.Format.LeftIndent = Unit.FromPoint(format.LeftIndent);
            if (format.FirstLineIndent != 0) target.Format.FirstLineIndent = Unit.FromPoint(format.FirstLineIndent);
            if (format.SpacingBefore > 0)
                target.Format.SpaceBefore = Unit.FromPoint(
                    Math.Min(format.SpacingBefore, LayoutDefaults.TableCellSpacingCapPoints));
            if (format.SpacingAfter > 0)
                target.Format.SpaceAfter = Unit.FromPoint(
                    Math.Min(format.SpacingAfter, LayoutDefaults.TableCellSpacingCapPoints));

            var added = WordContentRenderer.Render(ctx, source, source, target, options);
            if (!added && i > 0) target.AddText(" ");
        }
    }

    static void ApplyRowProperties(Row row, W.TableRow sourceRow)
    {
        var rowProperties = sourceRow.GetFirstChild<W.TableRowProperties>();
        if (rowProperties == null) return;

        var height = rowProperties.GetFirstChild<W.TableRowHeight>()?.Val;
        if (height is { HasValue: true }) row.Height = Unit.FromPoint(Units.TwipsToPoints(height.Value));

        if (rowProperties.GetFirstChild<W.TableHeader>() != null) row.HeadingFormat = true;
    }

    static void ApplyShading(Cell cell, W.TableCellProperties? tcPr, W.TableStyleProperties? conditionalStyle)
    {
        var fill = tcPr?.GetFirstChild<W.Shading>()?.Fill?.Value;
        if (!ColorUtils.TryParse(fill, out var color))
        {
            var conditionalFill = conditionalStyle?.TableStyleConditionalFormattingTableCellProperties
                ?.GetFirstChild<W.Shading>()?.Fill?.Value;
            if (!ColorUtils.TryParse(conditionalFill, out color)) return;
        }
        cell.Shading.Color = color;
    }

    static void ApplyCellPadding(Cell cell, BorderInfo borders)
    {
        if (borders.PaddingTop > 0) cell.Format.SpaceBefore = Unit.FromPoint(borders.PaddingTop);
        if (borders.PaddingBottom > 0) cell.Format.SpaceAfter = Unit.FromPoint(borders.PaddingBottom);
    }

    static void ApplyVerticalAlignment(Cell cell, W.TableCellProperties? tcPr)
    {
        var alignment = tcPr?.GetFirstChild<W.TableCellVerticalAlignment>()?.Val?.Value;
        if (alignment == null) return;
        cell.VerticalAlignment =
            alignment == W.TableVerticalAlignmentValues.Center ? VerticalAlignment.Center :
            alignment == W.TableVerticalAlignmentValues.Bottom ? VerticalAlignment.Bottom :
            VerticalAlignment.Top;
    }

    static WordHelpers.CellMargins ApplyCellMargins(Table migTable, WordRenderContext ctx, W.TableProperties? tblPr)
    {
        var margins = WordHelpers.GetTableCellMargins(ctx.MainPart, tblPr);
        migTable.LeftPadding = Unit.FromPoint(margins.Left);
        migTable.RightPadding = Unit.FromPoint(margins.Right);
        migTable.TopPadding = Unit.FromPoint(margins.Top);
        migTable.BottomPadding = Unit.FromPoint(margins.Bottom);
        return margins;
    }

    static void ApplyTableIndent(Table migTable, W.TableProperties? tblPr, TableLayout layout, double contentWidth)
    {
        var indent = Units.TwipsToPoints(tblPr?.TableIndentation?.Width?.Value ?? 0);
        var justification = tblPr?.TableJustification?.Val?.Value;
        var slack = contentWidth - layout.TotalWidth;

        if (slack > 0)
        {
            if (justification == W.TableRowAlignmentValues.Center) indent = slack / 2;
            else if (justification == W.TableRowAlignmentValues.Right) indent = slack;
        }

        if (indent > 0) migTable.Rows.LeftIndent = Unit.FromPoint(indent);
    }

    static void ApplyTableBorders(Table migTable, W.TableBorders? borders)
    {
        if (borders == null)
        {
            migTable.Borders.Width = Unit.FromPoint(LayoutDefaults.BorderWidthPoints);
            migTable.Borders.Color = Colors.Black;
            return;
        }

        ApplyBorderEdge(migTable.Borders.Top, borders.TopBorder);
        ApplyBorderEdge(migTable.Borders.Bottom, borders.BottomBorder);
        ApplyBorderEdge(migTable.Borders.Left, borders.LeftBorder);
        ApplyBorderEdge(migTable.Borders.Right, borders.RightBorder);

        ApplyInsideBorder(migTable.Borders.Top, migTable.Borders.Bottom, borders.InsideHorizontalBorder);
        ApplyInsideBorder(migTable.Borders.Left, migTable.Borders.Right, borders.InsideVerticalBorder);
    }

    static void ApplyInsideBorder(Border first, Border second, W.BorderType? source)
    {
        if (source?.Val == null || source.Val.Value == W.BorderValues.None) return;

        var width = source.Size is { HasValue: true } size
            ? Units.EighthPointsToPoints(size.Value)
            : LayoutDefaults.BorderWidthPoints;

        foreach (var border in new[] { first, second })
        {
            if (border.Width.Point != 0) continue;
            border.Width = Unit.FromPoint(width);
            border.Color = ColorUtils.TryParse(source.Color?.Value, out var color) ? color : Colors.Black;
        }
    }

    static void ApplyBorderEdge(Border border, W.BorderType? source)
    {
        if (source == null) return;
        if (source.Val?.Value == W.BorderValues.None || source.Val?.Value == W.BorderValues.Nil)
        {
            border.Width = 0;
            return;
        }

        border.Width = Unit.FromPoint(source.Size is { HasValue: true } size
            ? Units.EighthPointsToPoints(size.Value)
            : LayoutDefaults.BorderWidthPoints);
        border.Color = ColorUtils.TryParse(source.Color?.Value, out var color) ? color : Colors.Black;
    }

    static void ApplyCellBorders(Cell cell, BorderInfo borders)
    {
        ApplyEdge(cell.Borders.Top, borders.TopWidth, borders.TopColor);
        ApplyEdge(cell.Borders.Bottom, borders.BottomWidth, borders.BottomColor);
        ApplyEdge(cell.Borders.Left, borders.LeftWidth, borders.LeftColor);
        ApplyEdge(cell.Borders.Right, borders.RightWidth, borders.RightColor);
    }

    static void ApplyEdge(Border border, double width, string? colorValue)
    {
        if (width <= 0) return;
        border.Width = Unit.FromPoint(width);
        border.Color = ColorUtils.TryParse(colorValue, out var color) ? color : Colors.Black;
    }
}
