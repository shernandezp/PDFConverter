using SkiaSharp;

namespace PDFConverter.Layout;

internal abstract record DrawOp;

internal sealed record TextOp(string Text, ShapedText Shaped, FontFace Face, double Size, SKColor Color, double X,
    double Baseline, double WidthScale = 1) : DrawOp;

internal sealed record ImageOp(ImageSource Image, double X, double Y, double Width, double Height) : DrawOp;

internal sealed record FillRectOp(double X, double Y, double Width, double Height, SKColor Color) : DrawOp;

internal sealed record LineOp(double X1, double Y1, double X2, double Y2, double Width, SKColor Color,
    UnderlineStyle Pattern = UnderlineStyle.Single) : DrawOp;

internal sealed record LinkOp(double X, double Y, double Width, double Height, string Url) : DrawOp;

internal sealed record DestinationOp(string Name, double X, double Y) : DrawOp;

internal sealed record ClipOp(double X, double Y, double Width, double Height, List<DrawOp> Children) : DrawOp;

internal sealed class Page(Section section, double width, double height)
{
    public Section Section { get; } = section;
    public double Width { get; } = width;
    public double Height { get; } = height;
    public bool IsFirstOfSection { get; init; }
    public List<DrawOp> Background { get; } = [];
    public List<DrawOp> Body { get; } = [];
    public List<DrawOp> Overlay { get; } = [];
}

internal static class DrawOps
{
    public static DrawOp Translate(DrawOp op, double dx, double dy) => op switch
    {
        TextOp text => text with { X = text.X + dx, Baseline = text.Baseline + dy },
        ImageOp image => image with { X = image.X + dx, Y = image.Y + dy },
        FillRectOp rect => rect with { X = rect.X + dx, Y = rect.Y + dy },
        LineOp line => line with { X1 = line.X1 + dx, Y1 = line.Y1 + dy, X2 = line.X2 + dx, Y2 = line.Y2 + dy },
        LinkOp link => link with { X = link.X + dx, Y = link.Y + dy },
        DestinationOp destination => destination with { X = destination.X + dx, Y = destination.Y + dy },
        ClipOp clip => clip with { X = clip.X + dx, Y = clip.Y + dy, Children = Translate(clip.Children, dx, dy) },
        _ => op,
    };

    public static List<DrawOp> Translate(IEnumerable<DrawOp> ops, double dx, double dy) =>
        ops.Select(op => Translate(op, dx, dy)).ToList();
}
