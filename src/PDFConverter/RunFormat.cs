using MigraDoc.DocumentObjectModel;

namespace PDFConverter;

internal sealed record RunFormat(
    string? FontFamily,
    string? Color,
    bool Bold,
    bool Italic,
    bool Underline,
    double? Size)
{
    internal void ApplyTo(FormattedText formatted)
    {
        if (Size.HasValue) formatted.Size = Size.Value;
        try { if (!string.IsNullOrEmpty(FontFamily)) formatted.Font.Name = FontFamily; } catch { }
        try { if (!string.IsNullOrEmpty(Color)) formatted.Color = MigraDoc.DocumentObjectModel.Color.Parse("#" + Color); } catch { }
        if (Bold) formatted.Bold = true;
        if (Italic) formatted.Italic = true;
        if (Underline) formatted.Underline = MigraDoc.DocumentObjectModel.Underline.Single;
    }
}
