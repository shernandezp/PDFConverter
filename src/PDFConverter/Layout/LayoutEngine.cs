using SkiaSharp;

namespace PDFConverter.Layout;

internal sealed class LayoutEngine(FontRegistry fonts) : IDisposable
{
    readonly TextMeasurer _measurer = new(fonts);

    public void Dispose() => _measurer.Dispose();
    int _totalPagesHint = 1;

    sealed class Flow
    {
        public required Section Section;
        public required double X;
        public required double Width;
        public required List<DrawOp> Ops;
        public double Y;
        public double Bottom = double.MaxValue;
        public bool Paginated;
        public bool PageIsEmpty;
        public bool AtExplicitStart;
        public (double X, double Y)? Origin;
        public List<Page>? Pages;
        public Page? Page;
        public Action<Flow>? NewPage;
        public List<Exclusion> Exclusions = [];

        public bool EnsureRoom(double height)
        {
            if (!Paginated || PageIsEmpty || Y + height <= Bottom + 0.01) return true;
            NewPage!(this);
            return false;
        }

        public int PageNumber => Pages?.Count ?? 1;
    }

    public List<Page> Layout(Document document)
    {
        var pages = LayoutBody(document);
        if (pages.Count != _totalPagesHint && ContainsFields(document))
        {
            _totalPagesHint = pages.Count;
            pages = LayoutBody(document);
        }
        _totalPagesHint = pages.Count;
        AddHeadersAndFooters(pages);
        return pages;
    }

    static bool ContainsFields(Document document) =>
        document.Sections.SelectMany(s => s.Blocks).Any(ContainsFields);

    static bool ContainsFields(Block block) => block switch
    {
        Paragraph paragraph => paragraph.Inlines.Any(i => i is FieldRun),
        Table table => table.Rows.SelectMany(r => r.Cells).Any(c => c != null && c.Blocks.Any(ContainsFields)),
        _ => false,
    };

    List<Page> LayoutBody(Document document)
    {
        var pages = new List<Page>();
        foreach (var section in document.Sections)
        {
            var flow = new Flow
            {
                Section = section,
                X = section.LeftMargin,
                Width = section.ContentWidth,
                Ops = [],
                Paginated = true,
                Pages = pages,
                Origin = (0, 0),
                NewPage = StartPage,
            };
            StartPage(flow);
            flow.AtExplicitStart = true;
            LayoutBlocks(section.Blocks, flow);
        }
        return pages;
    }

    static void StartPage(Flow flow)
    {
        var section = flow.Section;
        var page = new Page(section, section.PageWidth, section.PageHeight)
        {
            IsFirstOfSection = flow.Page == null || flow.Page.Section != section,
        };
        foreach (var art in section.Backgrounds)
        {
            page.Background.Add(new ImageOp(art.Image, 0, 0,
                art.Width > 0 ? art.Width : section.PageWidth,
                art.Height > 0 ? art.Height : section.PageHeight));
        }
        flow.Pages!.Add(page);
        flow.Page = page;
        flow.Ops = page.Body;
        flow.Y = section.TopMargin;
        flow.Bottom = section.PageHeight - section.BottomMargin;
        flow.PageIsEmpty = true;
        flow.AtExplicitStart = false;
        flow.Exclusions.Clear();
    }

    readonly record struct Exclusion(double Left, double Top, double Right, double Bottom, WrapSide Side);

    const double MinimumWrapWidth = 36;

    static LineGeometry Geometry(Flow flow, double top)
    {
        if (flow.Exclusions.Count == 0) return (_, _) => new LineSegment(0, flow.Width);
        return (yOffset, height) =>
        {
            var y = top + yOffset;
            var gap = 0.0;
            for (var attempt = 0; attempt < 8; attempt++)
            {
                var left = flow.X;
                var right = flow.X + flow.Width;
                double pushTo = 0;
                foreach (var exclusion in flow.Exclusions)
                {
                    if (exclusion.Bottom <= y + gap || exclusion.Top >= y + gap + height) continue;
                    var leftRoom = exclusion.Left - left;
                    var rightRoom = right - exclusion.Right;
                    var useLeft = exclusion.Side switch
                    {
                        WrapSide.Left => true,
                        WrapSide.Right => false,
                        _ => leftRoom >= rightRoom,
                    };
                    if (useLeft) right = Math.Min(right, exclusion.Left);
                    else left = Math.Max(left, exclusion.Right);
                    pushTo = Math.Max(pushTo, exclusion.Bottom);
                }
                if (right - left >= MinimumWrapWidth || pushTo <= y + gap)
                    return new LineSegment(left - flow.X, Math.Max(right - left, 1), gap);
                gap = pushTo - y;
            }
            return new LineSegment(0, flow.Width, gap);
        };
    }

    string FieldText(FieldKind kind, int pageNumber) =>
        kind == FieldKind.Page ? pageNumber.ToString() : _totalPagesHint.ToString();

    void LayoutBlocks(List<Block> blocks, Flow flow)
    {
        for (var i = 0; i < blocks.Count; i++)
        {
            switch (blocks[i])
            {
                case Paragraph { KeepWithNext: true } paragraph when flow.Paginated && !flow.PageIsEmpty && i + 1 < blocks.Count:
                    LayoutKeptParagraph(paragraph, blocks[i + 1], flow);
                    break;
                case Paragraph paragraph:
                    LayoutParagraph(paragraph, flow);
                    break;
                case Table table:
                    LayoutTable(table, flow);
                    break;
                case FloatingBlock floating:
                    LayoutFloating(floating, flow);
                    break;
            }
        }
    }

    void LayoutKeptParagraph(Paragraph paragraph, Block next, Flow flow)
    {
        var pageCount = flow.Pages!.Count;
        var ops = flow.Ops;
        var opsCount = ops.Count;
        var y = flow.Y;
        var explicitStart = flow.AtExplicitStart;

        LayoutParagraph(paragraph, flow);
        if (flow.Pages.Count != pageCount || flow.Y + FirstHeight(next, flow) <= flow.Bottom + 0.01) return;

        ops.RemoveRange(opsCount, ops.Count - opsCount);
        flow.Y = y;
        flow.AtExplicitStart = explicitStart;
        flow.NewPage!(flow);
        LayoutParagraph(paragraph, flow);
    }

    double FirstHeight(Block block, Flow flow) => block switch
    {
        Paragraph paragraph => paragraph.SpaceBefore
            + TextLayout.Layout(paragraph, flow.Width, _measurer, kind => FieldText(kind, flow.PageNumber)).Lines[0].Height,
        Table { Rows.Count: > 0, ColumnWidths.Count: > 0 } table => MeasureRows(table, Covered(table))[0].Height,
        _ => 0,
    };

    void LayoutParagraph(Paragraph paragraph, Flow flow)
    {
        if (paragraph.PageBreakBefore && flow.Paginated && !flow.PageIsEmpty)
        {
            flow.NewPage!(flow);
            flow.AtExplicitStart = true;
        }

        var spaceBefore = flow.Paginated && flow.PageIsEmpty && !flow.AtExplicitStart ? 0 : paragraph.SpaceBefore;
        var lines = TextLayout.Layout(paragraph, flow.Width, _measurer,
            kind => FieldText(kind, flow.PageNumber), Geometry(flow, flow.Y + spaceBefore)).Lines;
        if (lines.Count == 0) return;

        if (spaceBefore > 0 && !flow.EnsureRoom(spaceBefore + lines[0].Height))
            spaceBefore = 0;
        flow.Y += spaceBefore;

        var index = 0;
        var firstFragment = true;
        while (index < lines.Count)
        {
            var count = flow.Paginated ? FitCount(paragraph, lines, index, flow) : lines.Count - index;
            if (count == 0)
            {
                flow.NewPage!(flow);
                continue;
            }

            var fragmentStart = flow.Y;
            var fragmentIndex = flow.Ops.Count;
            var fragmentOps = flow.Ops;
            var pageNumber = flow.PageNumber;
            for (var i = index; i < index + count; i++)
            {
                TextLayout.Emit(lines[i], flow.X, flow.Y, kind => FieldText(kind, pageNumber), _measurer, flow.Ops);
                flow.Y += lines[i].Height;
            }
            flow.PageIsEmpty = false;
            flow.AtExplicitStart = false;
            index += count;

            Decorate(paragraph, flow, fragmentOps, fragmentIndex, fragmentStart, firstFragment, index == lines.Count);
            firstFragment = false;
        }

        flow.Y += paragraph.SpaceAfter;
    }

    static int FitCount(Paragraph paragraph, List<LaidOutLine> lines, int index, Flow flow)
    {
        var remaining = lines.Count - index;
        var available = flow.Bottom - flow.Y + 0.01;
        var fit = 0;
        double used = 0;
        while (fit < remaining && used + lines[index + fit].Height <= available)
        {
            used += lines[index + fit].Height;
            fit++;
        }

        if (fit >= remaining) return remaining;
        if (flow.PageIsEmpty) return Math.Max(fit, 1);
        if (paragraph.KeepTogether) return 0;
        if (paragraph.WidowControl && remaining >= 2)
        {
            if (fit == 1 && index == 0) return 0;
            if (remaining - fit == 1) fit--;
        }
        return Math.Max(fit, 0);
    }

    static void Decorate(Paragraph paragraph, Flow flow, List<DrawOp> ops, int insertAt, double top,
        bool first, bool last)
    {
        var bottom = flow.Y;
        if (bottom <= top) return;
        var left = flow.X + paragraph.LeftIndent;
        var width = Math.Max(flow.Width - paragraph.LeftIndent - paragraph.RightIndent, 0);

        if (paragraph.Shading is { } shading)
            ops.Insert(insertAt, new FillRectOp(left, top, width, bottom - top, shading));

        var borders = paragraph.Borders;
        if (borders == null) return;
        if (first) Edge(borders.Top, left, top, left + width, top, ops);
        if (last) Edge(borders.Bottom, left, bottom, left + width, bottom, ops);
        Edge(borders.Left, left, top, left, bottom, ops);
        Edge(borders.Right, left + width, top, left + width, bottom, ops);
    }

    void LayoutFloating(FloatingBlock floating, Flow flow)
    {
        var section = flow.Section;
        double x, y;
        if (flow.Origin is { } origin)
        {
            x = (floating.RelativeHorizontal == RelativeHorizontal.Page ? 0 : section.LeftMargin) + floating.Left - origin.X;
            y = floating.RelativeVertical switch
            {
                RelativeVertical.Page => floating.Top - origin.Y,
                RelativeVertical.Margin => section.TopMargin + floating.Top - origin.Y,
                _ => flow.Y + floating.Top,
            };
        }
        else
        {
            x = flow.X + floating.Left;
            y = flow.Y + floating.Top;
        }

        if (floating.Image != null)
        {
            flow.Ops.Add(new ImageOp(floating.Image, x, y, floating.Width, floating.Height));
            if (floating.Url != null)
                flow.Ops.Add(new LinkOp(x, y, floating.Width, floating.Height, floating.Url));
            if (floating.WrapSide != WrapSide.None && flow.Paginated)
                flow.Exclusions.Add(new Exclusion(x - floating.DistanceLeft, y - floating.DistanceTop,
                    x + floating.Width + floating.DistanceRight, y + floating.Height + floating.DistanceBottom, floating.WrapSide));
        }
        else if (floating.LineColor is { } color && floating.LineWidth > 0)
        {
            flow.Ops.Add(new LineOp(x, y, x + floating.Width, y, floating.LineWidth, color));
        }

        if (!floating.ReservesSpace) return;
        var reserved = Math.Max(floating.Top, 0) + floating.Height;
        flow.EnsureRoom(reserved);
        flow.Y += reserved;
        flow.PageIsEmpty = false;
    }

    sealed class CellMeasure
    {
        public required TableCell Cell;
        public required int Column;
        public required List<DrawOp> Ops;
        public required double ContentHeight;
    }

    sealed class RowMeasure
    {
        public double Height;
        public List<CellMeasure> Cells = [];
    }

    void LayoutTable(Table table, Flow flow)
    {
        if (table.Rows.Count == 0 || table.ColumnWidths.Count == 0) return;

        var covered = Covered(table);
        var rows = MeasureRows(table, covered);
        var x0 = flow.X + table.LeftIndent;

        var headerCount = 0;
        while (headerCount < table.Rows.Count && table.Rows[headerCount].IsHeader) headerCount++;

        for (var i = 0; i < table.Rows.Count; i++)
        {
            var height = rows[i].Height;
            if (!flow.EnsureRoom(height) && i >= headerCount)
            {
                for (var h = 0; h < headerCount; h++)
                {
                    PlaceRow(table, rows, covered, h, x0, flow.Y, flow.Ops);
                    flow.Y += rows[h].Height;
                }
            }

            PlaceRow(table, rows, covered, i, x0, flow.Y, flow.Ops);
            flow.Y += height;
            flow.PageIsEmpty = false;
            flow.AtExplicitStart = false;
        }
    }

    static HashSet<(int Row, int Column)> Covered(Table table)
    {
        var covered = new HashSet<(int, int)>();
        for (var r = 0; r < table.Rows.Count; r++)
        {
            var cells = table.Rows[r].Cells;
            for (var c = 0; c < cells.Length; c++)
            {
                var cell = cells[c];
                if (cell == null || covered.Contains((r, c))) continue;
                for (var dr = 0; dr < cell.RowSpan; dr++)
                    for (var dc = 0; dc < cell.ColumnSpan; dc++)
                        if ((dr > 0 || dc > 0) && r + dr < table.Rows.Count && c + dc < cells.Length)
                            covered.Add((r + dr, c + dc));
            }
        }
        return covered;
    }

    RowMeasure[] MeasureRows(Table table, HashSet<(int Row, int Column)> covered)
    {
        var rows = new RowMeasure[table.Rows.Count];
        var spanning = new List<(int Row, int Span, double Height)>();

        for (var r = 0; r < table.Rows.Count; r++)
        {
            var row = table.Rows[r];
            var measure = new RowMeasure { Height = row.MinHeight ?? 0 };
            rows[r] = measure;

            for (var c = 0; c < row.Cells.Length; c++)
            {
                var cell = row.Cells[c];
                if (cell == null || covered.Contains((r, c))) continue;

                var span = Math.Min(cell.ColumnSpan, row.Cells.Length - c);
                var contentWidth = Math.Max(table.WidthOf(c, span) - table.PaddingLeft - table.PaddingRight, 1);
                var (ops, contentHeight) = LayoutUnpaginated(cell.Blocks, contentWidth);
                measure.Cells.Add(new CellMeasure { Cell = cell, Column = c, Ops = ops, ContentHeight = contentHeight });

                var cellHeight = contentHeight + table.PaddingTop + table.PaddingBottom
                    + cell.ExtraPaddingTop + cell.ExtraPaddingBottom;
                if (cell.RowSpan > 1)
                    spanning.Add((r, Math.Min(cell.RowSpan, table.Rows.Count - r), cellHeight));
                else if (!row.ExactHeight)
                    measure.Height = Math.Max(measure.Height, cellHeight);
            }

            if (row.ExactHeight && row.MinHeight is { } exact) measure.Height = exact;
        }

        foreach (var (start, span, height) in spanning)
        {
            double total = 0;
            for (var r = start; r < start + span; r++) total += rows[r].Height;
            var last = start + span - 1;
            if (height > total && !table.Rows[last].ExactHeight) rows[last].Height += height - total;
        }
        return rows;
    }

    void PlaceRow(Table table, RowMeasure[] rows, HashSet<(int Row, int Column)> covered, int rowIndex,
        double x0, double y, List<DrawOp> ops)
    {
        var row = table.Rows[rowIndex];
        var rowHeight = rows[rowIndex].Height;
        var columnX = x0;

        for (var c = 0; c < row.Cells.Length; c++)
        {
            var cell = row.Cells[c];
            var width = table.ColumnWidths[c];

            if (covered.Contains((rowIndex, c)))
            {
                if (cell != null)
                    DrawBorders(cell.Borders, table.DefaultBorders, columnX, y, width, rowHeight, ops);
                columnX += width;
                continue;
            }

            if (cell == null)
            {
                DrawBorders(null, table.DefaultBorders, columnX, y, width, rowHeight, ops);
                columnX += width;
                continue;
            }

            var span = Math.Min(cell.ColumnSpan, row.Cells.Length - c);
            var cellWidth = table.WidthOf(c, span);
            var cellHeight = rowHeight;
            for (var r = rowIndex + 1; r < Math.Min(rowIndex + cell.RowSpan, rows.Length); r++)
                cellHeight += rows[r].Height;

            if (cell.Shading is { } shading)
                ops.Add(new FillRectOp(columnX, y, cellWidth, cellHeight, shading));

            var measure = rows[rowIndex].Cells.FirstOrDefault(m => m.Column == c);
            if (measure != null)
            {
                var available = cellHeight - table.PaddingTop - table.PaddingBottom
                    - cell.ExtraPaddingTop - cell.ExtraPaddingBottom;
                var offset = cell.VerticalAlignment switch
                {
                    VerticalAlignment.Center => Math.Max(0, (available - measure.ContentHeight) / 2),
                    VerticalAlignment.Bottom => Math.Max(0, available - measure.ContentHeight),
                    _ => 0,
                };
                var content = DrawOps.Translate(measure.Ops, columnX + table.PaddingLeft,
                    y + table.PaddingTop + cell.ExtraPaddingTop + offset);
                if (row.ExactHeight || cell.ClipContent)
                    ops.Add(new ClipOp(columnX - cell.OverflowLeft, y,
                        cellWidth + cell.OverflowLeft + cell.OverflowRight, cellHeight, content));
                else
                    ops.AddRange(content);
            }

            DrawBorders(cell.Borders, table.DefaultBorders, columnX, y,
                cell.BordersSpanColumns ? cellWidth : width, cellHeight, ops);
            columnX += width;
        }

        foreach (var line in row.Lines)
            ops.Add(new LineOp(x0 + line.StartX, y + line.OffsetY, x0 + line.EndX, y + line.OffsetY, line.Width, line.Color));

        foreach (var image in row.Images)
            ops.Add(new ImageOp(image.Image, x0 + image.OffsetX, y + image.OffsetY, image.Width, image.Height));
    }

    static void DrawBorders(Borders? cell, Borders defaults, double x, double y, double width, double height, List<DrawOp> ops)
    {
        Edge(cell?.Top ?? defaults.Top, x, y, x + width, y, ops);
        Edge(cell?.Bottom ?? defaults.Bottom, x, y + height, x + width, y + height, ops);
        Edge(cell?.Left ?? defaults.Left, x, y, x, y + height, ops);
        Edge(cell?.Right ?? defaults.Right, x + width, y, x + width, y + height, ops);
    }

    static void Edge(BorderEdge? edge, double x1, double y1, double x2, double y2, List<DrawOp> ops)
    {
        if (edge is not { IsVisible: true } visible) return;
        ops.Add(new LineOp(x1, y1, x2, y2, visible.Width, visible.Color, visible.Pattern));
    }

    (List<DrawOp> Ops, double Height) LayoutUnpaginated(List<Block> blocks, double width, Section? section = null,
        (double X, double Y)? origin = null, int pageNumber = 1)
    {
        var flow = new Flow
        {
            Section = section ?? new Section(),
            X = 0,
            Width = width,
            Ops = [],
            Origin = origin,
        };
        LayoutBlocksWithPage(blocks, flow, pageNumber);
        return (flow.Ops, flow.Y);
    }

    void LayoutBlocksWithPage(List<Block> blocks, Flow flow, int pageNumber)
    {
        if (pageNumber == 1)
        {
            LayoutBlocks(blocks, flow);
            return;
        }

        var pages = new List<Page>();
        for (var i = 0; i < pageNumber; i++) pages.Add(new Page(flow.Section, 0, 0));
        flow.Pages = pages;
        LayoutBlocks(blocks, flow);
        flow.Pages = null;
    }

    void AddHeadersAndFooters(List<Page> pages)
    {
        for (var i = 0; i < pages.Count; i++)
        {
            var page = pages[i];
            var section = page.Section;
            var pageNumber = i + 1;

            var header = Select(section, page.IsFirstOfSection, pageNumber,
                section.FirstPageHeader, section.EvenPageHeader, section.PrimaryHeader);
            if (!header.IsEmpty)
            {
                var (ops, _) = LayoutUnpaginated(header.Blocks, section.ContentWidth, section,
                    (section.LeftMargin, section.HeaderDistance), pageNumber);
                page.Overlay.AddRange(DrawOps.Translate(ops, section.LeftMargin, section.HeaderDistance));
            }

            var footer = Select(section, page.IsFirstOfSection, pageNumber,
                section.FirstPageFooter, section.EvenPageFooter, section.PrimaryFooter);
            if (!footer.IsEmpty)
            {
                var (ops, height) = LayoutUnpaginated(footer.Blocks, section.ContentWidth, section, null, pageNumber);
                var top = section.PageHeight - section.FooterDistance - height;
                page.Overlay.AddRange(DrawOps.Translate(ops, section.LeftMargin, top));
            }
        }
    }

    static HeaderFooter Select(Section section, bool firstOfSection, int pageNumber,
        HeaderFooter first, HeaderFooter even, HeaderFooter primary)
    {
        if (firstOfSection && section.DifferentFirstPage) return first;
        if (section.OddAndEvenPages && pageNumber % 2 == 0) return even;
        return primary;
    }
}
