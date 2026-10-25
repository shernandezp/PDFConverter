using PDFConverter.Layout;
using SkiaSharp;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal static class WordTableRenderer
{
    public static void RenderTable(WordRenderContext ctx, Section section, W.Table table)
    {
        var built = Build(ctx, table, section.ContentWidth, topLevel: true);
        if (built != null) section.Blocks.Add(built);
    }

    public static void RenderNestedTable(WordRenderContext ctx, W.Table table, TableCell parentCell, double parentWidth)
    {
        var built = Build(ctx, table, parentWidth, topLevel: false);
        if (built != null) parentCell.Blocks.Add(built);
    }

    static Table? Build(WordRenderContext ctx, W.Table table, double availableWidth, bool topLevel)
    {
        var rows = table.Elements<W.TableRow>().ToList();
        if (rows.Count == 0) return null;

        var layout = TableLayout.Measure(table, rows, availableWidth);
        var target = new Table();
        target.ColumnWidths.AddRange(layout.ScaledWidths);

        var tblPr = table.GetFirstChild<W.TableProperties>();
        ApplyTableBorders(target, WordHelpers.ResolveTableBorders(ctx.MainPart, tblPr));
        var margins = ApplyCellMargins(target, ctx, tblPr);
        if (topLevel) ApplyTableIndent(target, tblPr, layout, availableWidth);

        var firstRowStyle = topLevel
            ? WordHelpers.GetTableStyleConditionalFormatting(ctx.MainPart, tblPr, W.TableStyleOverrideValues.FirstRow)
            : null;
        var firstColumnStyle = topLevel
            ? WordHelpers.GetTableStyleConditionalFormatting(ctx.MainPart, tblPr, W.TableStyleOverrideValues.FirstColumn)
            : null;

        var verticalSpans = MeasureVerticalMerges(rows, layout.ColumnCount);

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var sourceRow = rows[rowIndex];
            var row = target.AddRow();
            ApplyRowProperties(row, sourceRow);

            var isFirstRow = topLevel && WordHelpers.IsConditionalRow(
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

                var cell = row.Cell(column);
                var conditionalStyle = isFirstRow ? firstRowStyle : column == 0 ? firstColumnStyle : null;

                ApplyShading(cell, tcPr, conditionalStyle);
                var borders = WordHelpers.GetWordCellBorders(tcPr, tblPr);
                ApplyCellBorders(cell, borders);
                if (borders.PaddingTop > 0) cell.ExtraPaddingTop = borders.PaddingTop;
                if (borders.PaddingBottom > 0) cell.ExtraPaddingBottom = borders.PaddingBottom;
                ApplyVerticalAlignment(cell, tcPr);

                cell.ColumnSpan = Math.Max(1, Math.Min(gridSpan, layout.ColumnCount - column));
                if (verticalSpans.TryGetValue((rowIndex, column), out var mergeDown) && mergeDown > 0)
                    cell.RowSpan = mergeDown + 1;

                var textWidth = layout.WidthOf(column, gridSpan) - margins.Left - margins.Right;
                RenderCellContent(ctx, cell, sourceCell, textWidth, conditionalStyle);
                column += gridSpan;
            }

            if (column == 0) RenderDrawingMlRow(ctx, sourceRow, row, layout.ColumnCount);
        }

        return target;
    }

    const string DrawingMlNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const string WordprocessingMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Word occasionally emits DrawingML cells (a:tc) inside an ordinary w:tr; they carry text and
    // hyperlinks but none of the WordprocessingML cell structure.
    static void RenderDrawingMlRow(WordRenderContext ctx, W.TableRow sourceRow, TableRow row, int columnCount)
    {
        var column = 0;
        foreach (var sourceCell in sourceRow.ChildElements)
        {
            if (sourceCell.LocalName != "tc" || sourceCell.NamespaceUri != DrawingMlNamespace) continue;
            if (column >= columnCount) break;

            var target = ctx.Prepare(row.Cell(column).AddParagraph());
            foreach (var paragraph in sourceCell.ChildElements.Where(c => c.LocalName == "p"))
            {
                foreach (var child in paragraph.ChildElements)
                {
                    if (child.LocalName == "hyperlink" && child.NamespaceUri == WordprocessingMlNamespace)
                        AddDrawingMlHyperlink(ctx, child, target);
                    else if (child.LocalName == "r" && child.NamespaceUri == DrawingMlNamespace)
                        target.Add(new TextRun(DrawingMlText(child), target.DefaultStyle));
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

        var style = target.DefaultStyle with { Color = SKColors.Blue, Underline = UnderlineStyle.Single };
        target.Add(new TextRun(text, style, url));
    }

    static string DrawingMlText(DocumentFormat.OpenXml.OpenXmlElement source) =>
        string.Concat(source.Descendants().Where(d => d.LocalName == "t").Select(d => d.InnerText));

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

    static void RenderCellContent(WordRenderContext ctx, TableCell cell, W.TableCell sourceCell,
        double cellWidth, W.TableStyleProperties? conditionalStyle)
    {
        var conditionalRun = conditionalStyle?.RunPropertiesBaseStyle;
        var options = new InlineOptions(
            MaxImageWidthPoints: cellWidth * 0.9,
            ConditionalBold: conditionalRun?.Bold != null
                && (conditionalRun.Bold.Val == null || conditionalRun.Bold.Val.Value),
            ConditionalColor: conditionalRun?.Color?.Val?.Value);

        var paragraphs = new List<W.Paragraph>();
        foreach (var child in sourceCell.ChildElements)
        {
            if (child is W.Paragraph paragraph) paragraphs.Add(paragraph);
            else if (child is W.SdtBlock { SdtContentBlock: not null } sdt)
                paragraphs.AddRange(sdt.SdtContentBlock.Elements<W.Paragraph>());
        }

        if (paragraphs.Count == 0)
        {
            ctx.Prepare(cell.AddParagraph());
            return;
        }

        foreach (var source in paragraphs)
        {
            var target = ctx.Prepare(cell.AddParagraph());
            var format = WordHelpers.GetParagraphFormatting(source.ParagraphProperties);

            target.Alignment = format.Bidi ? WordDocumentBuilder.MirrorForRightToLeft(format.Alignment) : format.Alignment;
            target.RightToLeft = format.Bidi;
            if (format.LeftIndent > 0) target.LeftIndent = format.LeftIndent;
            if (format.RightIndent > 0) target.RightIndent = format.RightIndent;
            if (format.FirstLineIndent != 0) target.FirstLineIndent = format.FirstLineIndent;
            if (format.SpacingBefore > 0)
                target.SpaceBefore = Math.Min(format.SpacingBefore, LayoutDefaults.TableCellSpacingCapPoints);
            if (format.SpacingAfter > 0)
                target.SpaceAfter = Math.Min(format.SpacingAfter, LayoutDefaults.TableCellSpacingCapPoints);

            WordContentRenderer.Render(ctx, source, source, target, options);
        }
    }

    static void ApplyRowProperties(TableRow row, W.TableRow sourceRow)
    {
        var rowProperties = sourceRow.GetFirstChild<W.TableRowProperties>();
        if (rowProperties == null) return;

        var rowHeight = rowProperties.GetFirstChild<W.TableRowHeight>();
        if (rowHeight?.Val is { HasValue: true } height)
        {
            row.MinHeight = Units.TwipsToPoints(height.Value);
            row.ExactHeight = rowHeight.HeightType?.Value == W.HeightRuleValues.Exact;
        }

        if (rowProperties.GetFirstChild<W.TableHeader>() != null) row.IsHeader = true;
    }

    static void ApplyShading(TableCell cell, W.TableCellProperties? tcPr, W.TableStyleProperties? conditionalStyle)
    {
        var fill = tcPr?.GetFirstChild<W.Shading>()?.Fill?.Value;
        if (!ColorUtils.TryParse(fill, out var color))
        {
            var conditionalFill = conditionalStyle?.TableStyleConditionalFormattingTableCellProperties
                ?.GetFirstChild<W.Shading>()?.Fill?.Value;
            if (!ColorUtils.TryParse(conditionalFill, out color)) return;
        }
        cell.Shading = color;
    }

    static void ApplyVerticalAlignment(TableCell cell, W.TableCellProperties? tcPr)
    {
        var alignment = tcPr?.GetFirstChild<W.TableCellVerticalAlignment>()?.Val?.Value;
        if (alignment == null) return;
        cell.VerticalAlignment =
            alignment == W.TableVerticalAlignmentValues.Center ? VerticalAlignment.Center :
            alignment == W.TableVerticalAlignmentValues.Bottom ? VerticalAlignment.Bottom :
            VerticalAlignment.Top;
    }

    static WordHelpers.CellMargins ApplyCellMargins(Table table, WordRenderContext ctx, W.TableProperties? tblPr)
    {
        var margins = WordHelpers.GetTableCellMargins(ctx.MainPart, tblPr);
        table.PaddingLeft = margins.Left;
        table.PaddingRight = margins.Right;
        table.PaddingTop = margins.Top;
        table.PaddingBottom = margins.Bottom;
        return margins;
    }

    static void ApplyTableIndent(Table table, W.TableProperties? tblPr, TableLayout layout, double contentWidth)
    {
        var indent = Units.TwipsToPoints(tblPr?.TableIndentation?.Width?.Value ?? 0);
        var justification = tblPr?.TableJustification?.Val?.Value;
        var slack = contentWidth - layout.TotalWidth;

        if (slack > 0)
        {
            if (justification == W.TableRowAlignmentValues.Center) indent = slack / 2;
            else if (justification == W.TableRowAlignmentValues.Right) indent = slack;
        }

        if (indent > 0) table.LeftIndent = indent;
    }

    static void ApplyTableBorders(Table table, W.TableBorders? borders)
    {
        var defaults = table.DefaultBorders;
        if (borders == null)
        {
            defaults.SetAll(new BorderEdge(LayoutDefaults.BorderWidthPoints, SKColors.Black));
            return;
        }

        defaults.Top = Edge(borders.TopBorder);
        defaults.Bottom = Edge(borders.BottomBorder);
        defaults.Left = Edge(borders.LeftBorder);
        defaults.Right = Edge(borders.RightBorder);

        var insideH = Edge(borders.InsideHorizontalBorder);
        if (insideH is { IsVisible: true })
        {
            defaults.Top ??= insideH;
            defaults.Bottom ??= insideH;
            if (!defaults.Top.Value.IsVisible) defaults.Top = insideH;
            if (!defaults.Bottom.Value.IsVisible) defaults.Bottom = insideH;
        }

        var insideV = Edge(borders.InsideVerticalBorder);
        if (insideV is { IsVisible: true })
        {
            defaults.Left ??= insideV;
            defaults.Right ??= insideV;
            if (!defaults.Left.Value.IsVisible) defaults.Left = insideV;
            if (!defaults.Right.Value.IsVisible) defaults.Right = insideV;
        }
    }

    static BorderEdge? Edge(W.BorderType? source)
    {
        if (source == null) return null;
        if (source.Val?.Value == W.BorderValues.None || source.Val?.Value == W.BorderValues.Nil)
            return BorderEdge.None;

        var width = source.Size is { HasValue: true } size
            ? Units.EighthPointsToPoints(size.Value)
            : LayoutDefaults.BorderWidthPoints;
        return new BorderEdge(width, ColorUtils.TryParse(source.Color?.Value, out var color) ? color : SKColors.Black,
            BorderPatterns.From(source.Val?.InnerText));
    }

    static void ApplyCellBorders(TableCell cell, BorderInfo borders)
    {
        cell.Borders.Top = Edge(borders.TopWidth, borders.TopColor, borders.TopStyle);
        cell.Borders.Bottom = Edge(borders.BottomWidth, borders.BottomColor, borders.BottomStyle);
        cell.Borders.Left = Edge(borders.LeftWidth, borders.LeftColor, borders.LeftStyle);
        cell.Borders.Right = Edge(borders.RightWidth, borders.RightColor, borders.RightStyle);
    }

    static BorderEdge? Edge(double width, string? colorValue, string? style)
    {
        if (width <= 0) return style is "nil" or "none" ? BorderEdge.None : null;
        return new BorderEdge(width, ColorUtils.TryParse(colorValue, out var color) ? color : SKColors.Black,
            BorderPatterns.From(style));
    }
}
