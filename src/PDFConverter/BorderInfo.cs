namespace PDFConverter;

internal sealed record BorderInfo(
    double TopWidth,
    string? TopColor,
    string? TopStyle,
    double BottomWidth,
    string? BottomColor,
    string? BottomStyle,
    double LeftWidth,
    string? LeftColor,
    string? LeftStyle,
    double RightWidth,
    string? RightColor,
    string? RightStyle,
    double PaddingTop = 0,
    double PaddingBottom = 0,
    double PaddingLeft = 0,
    double PaddingRight = 0)
{
    public static BorderInfo Empty { get; } = new(0, null, null, 0, null, null, 0, null, null, 0, null, null);
}
