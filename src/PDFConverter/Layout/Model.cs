using SkiaSharp;

namespace PDFConverter.Layout;

internal enum ParagraphAlignment { Left, Center, Right, Justify }

internal enum VerticalAlignment { Top, Center, Bottom }

internal enum UnderlineStyle { Single, Words, Dotted, Dash, DotDash, DotDotDash }

internal enum StrikeStyle { None, Single, Double }

internal enum TabAlignment { Left, Center, Right }

internal enum LineSpacingRule { Multiple, Exactly, AtLeast }

internal enum RelativeHorizontal { Page, Margin }

internal enum RelativeVertical { Page, Margin, Paragraph }

internal enum FieldKind { Page, NumPages }

internal enum WrapSide { None, Left, Right, Both, Largest }

internal sealed record TextStyle(
    string? FontFamily,
    double Size,
    bool Bold = false,
    bool Italic = false,
    SKColor? Color = null,
    UnderlineStyle? Underline = null,
    RunVerticalAlignment VerticalAlignment = RunVerticalAlignment.Baseline)
{
    public StrikeStyle Strike { get; init; }
    public SKColor? Highlight { get; init; }
    public double LetterSpacing { get; init; }
    public double WidthScale { get; init; } = 1;

    public static TextStyle Default { get; } = new(null, LayoutDefaults.FontSizePoints);
}

internal readonly record struct TabStop(double Position, TabAlignment Alignment = TabAlignment.Left);

internal readonly record struct BorderEdge(double Width, SKColor Color, UnderlineStyle Pattern = UnderlineStyle.Single)
{
    public static BorderEdge None { get; } = new(0, SKColors.Black);
    public bool IsVisible => Width > 0;
}

internal sealed class Borders
{
    public BorderEdge? Top, Bottom, Left, Right;

    public void SetAll(BorderEdge edge) => Top = Bottom = Left = Right = edge;
}

internal abstract class Inline;

internal sealed class TextRun(string text, TextStyle style, string? url = null) : Inline
{
    public string Text { get; } = text;
    public TextStyle Style { get; } = style;
    public string? Url { get; } = url;
}

internal sealed class FieldRun(FieldKind kind, TextStyle style) : Inline
{
    public FieldKind Kind { get; } = kind;
    public TextStyle Style { get; } = style;
}

internal sealed class TabInline : Inline;

internal sealed class LineBreakInline : Inline;

internal sealed class BookmarkInline(string name) : Inline
{
    public string Name { get; } = name;
}

internal sealed class InlineImage(ImageSource image, double width, double? height, string? url = null) : Inline
{
    public ImageSource Image { get; } = image;
    public double Width { get; } = width;
    public double? Height { get; } = height;
    public string? Url { get; } = url;

    public double ResolvedHeight => Height ?? Image.HeightFor(Width);
}

internal abstract class Block;

internal sealed class Paragraph : Block
{
    public List<Inline> Inlines { get; } = [];
    public ParagraphAlignment Alignment { get; set; }
    public double LeftIndent { get; set; }
    public double RightIndent { get; set; }
    public double FirstLineIndent { get; set; }
    public double SpaceBefore { get; set; }
    public double SpaceAfter { get; set; }
    public double? LineSpacing { get; set; }
    public LineSpacingRule LineSpacingRule { get; set; } = LineSpacingRule.Multiple;
    public bool PageBreakBefore { get; set; }
    public SKColor? Shading { get; set; }
    public List<TabStop>? TabStops { get; set; }
    public TextStyle DefaultStyle { get; set; } = TextStyle.Default;
    public bool NoWrap { get; set; }
    public Borders? Borders { get; set; }
    public bool KeepWithNext { get; set; }
    public bool KeepTogether { get; set; }
    public bool WidowControl { get; set; } = true;
    public bool RightToLeft { get; set; }

    public bool IsEmpty => Inlines.Count == 0;

    public Paragraph Add(Inline inline)
    {
        Inlines.Add(inline);
        return this;
    }
}

internal sealed class Table : Block
{
    public List<double> ColumnWidths { get; } = [];
    public List<TableRow> Rows { get; } = [];
    public double LeftIndent { get; set; }
    public double PaddingLeft { get; set; }
    public double PaddingRight { get; set; }
    public double PaddingTop { get; set; }
    public double PaddingBottom { get; set; }
    public Borders DefaultBorders { get; } = new();

    public double Width => ColumnWidths.Sum();

    public double WidthOf(int column, int span)
    {
        double width = 0;
        for (var i = column; i < Math.Min(column + span, ColumnWidths.Count); i++) width += ColumnWidths[i];
        return width;
    }

    public TableRow AddRow()
    {
        var row = new TableRow(ColumnWidths.Count);
        Rows.Add(row);
        return row;
    }
}

internal sealed class TableRow(int columnCount)
{
    public TableCell?[] Cells { get; } = new TableCell?[columnCount];
    public double? MinHeight { get; set; }
    public bool ExactHeight { get; set; }
    public bool IsHeader { get; set; }
    public List<RowLine> Lines { get; } = [];
    public List<RowImage> Images { get; } = [];

    public TableCell Cell(int column) => Cells[column] ??= new TableCell();
}

internal readonly record struct RowLine(double StartX, double EndX, double OffsetY, SKColor Color, double Width);

internal sealed record RowImage(ImageSource Image, double OffsetX, double OffsetY, double Width, double Height);

internal sealed class TableCell
{
    public List<Block> Blocks { get; } = [];
    public int ColumnSpan { get; set; } = 1;
    public int RowSpan { get; set; } = 1;
    public SKColor? Shading { get; set; }
    public Borders Borders { get; } = new();
    public VerticalAlignment VerticalAlignment { get; set; }
    public double ExtraPaddingTop { get; set; }
    public double ExtraPaddingBottom { get; set; }
    public bool BordersSpanColumns { get; set; } = true;
    public double OverflowLeft { get; set; }
    public double OverflowRight { get; set; }
    public bool ClipContent { get; set; }

    public Paragraph AddParagraph()
    {
        var paragraph = new Paragraph();
        Blocks.Add(paragraph);
        return paragraph;
    }
}

internal sealed class FloatingBlock : Block
{
    public ImageSource? Image { get; init; }
    public SKColor? LineColor { get; init; }
    public double LineWidth { get; init; }
    public string? Url { get; init; }
    public required double Left { get; init; }
    public required double Top { get; init; }
    public required double Width { get; init; }
    public double Height { get; init; }
    public RelativeHorizontal RelativeHorizontal { get; init; } = RelativeHorizontal.Margin;
    public RelativeVertical RelativeVertical { get; init; } = RelativeVertical.Paragraph;
    public bool ReservesSpace { get; init; }
    public WrapSide WrapSide { get; init; }
    public double DistanceTop { get; init; }
    public double DistanceBottom { get; init; }
    public double DistanceLeft { get; init; }
    public double DistanceRight { get; init; }
}

internal sealed class HeaderFooter
{
    public List<Block> Blocks { get; } = [];
    public bool IsEmpty => Blocks.Count == 0;
}

internal sealed record PageArt(ImageSource Image, double Width, double Height);

internal sealed class Section
{
    public double PageWidth { get; set; } = 595.276;
    public double PageHeight { get; set; } = 841.89;
    public double LeftMargin { get; set; } = 72;
    public double RightMargin { get; set; } = 72;
    public double TopMargin { get; set; } = 72;
    public double BottomMargin { get; set; } = 72;
    public double HeaderDistance { get; set; } = 36;
    public double FooterDistance { get; set; } = 36;
    public bool DifferentFirstPage { get; set; }
    public bool OddAndEvenPages { get; set; }

    public HeaderFooter PrimaryHeader { get; } = new();
    public HeaderFooter FirstPageHeader { get; } = new();
    public HeaderFooter EvenPageHeader { get; } = new();
    public HeaderFooter PrimaryFooter { get; } = new();
    public HeaderFooter FirstPageFooter { get; } = new();
    public HeaderFooter EvenPageFooter { get; } = new();

    public List<PageArt> Backgrounds { get; } = [];
    public List<Block> Blocks { get; } = [];

    public double ContentWidth => PageWidth - LeftMargin - RightMargin;

    public Paragraph AddParagraph()
    {
        var paragraph = new Paragraph();
        Blocks.Add(paragraph);
        return paragraph;
    }

    public Table AddTable()
    {
        var table = new Table();
        Blocks.Add(table);
        return table;
    }
}

internal sealed class Document
{
    public List<Section> Sections { get; } = [];
    public string? Title { get; set; }
    public string? Author { get; set; }
    public string? Subject { get; set; }
    public string? Keywords { get; set; }

    public Section AddSection()
    {
        var section = new Section();
        Sections.Add(section);
        return section;
    }
}
