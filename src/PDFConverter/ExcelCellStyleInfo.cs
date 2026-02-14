namespace PDFConverter;

internal sealed record ExcelCellStyleInfo(
    string? HorizontalAlignment,
    string? VerticalAlignment,
    string? FillColor,
    uint? NumberFormatId,
    BorderInfo Borders,
    string? FontFamily,
    double? FontSize,
    string? FontColor,
    bool Bold,
    bool Italic)
{
    public static ExcelCellStyleInfo Empty { get; } = new(null, null, null, null, BorderInfo.Empty, null, null, null, false, false);
}
