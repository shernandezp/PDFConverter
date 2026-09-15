using MigraDoc.DocumentObjectModel;

namespace PDFConverter;

internal enum RunVerticalAlignment { Baseline, Superscript, Subscript }

internal sealed record RunFormat(
    string? FontFamily,
    string? Color,
    bool Bold,
    bool Italic,
    bool Underline,
    double? Size,
    bool BoldSpecified = false,
    Underline UnderlineStyle = MigraDoc.DocumentObjectModel.Underline.Single,
    RunVerticalAlignment VerticalAlignment = RunVerticalAlignment.Baseline,
    bool AllCaps = false)
{
    internal void ApplyTo(FormattedText formatted)
    {
        if (Size.HasValue) formatted.Size = Size.Value;
        if (!string.IsNullOrEmpty(FontFamily)) formatted.Font.Name = FontFamily;
        if (ColorUtils.TryParse(Color, out var color)) formatted.Color = color;
        if (Bold) formatted.Bold = true;
        if (Italic) formatted.Italic = true;
        if (Underline) formatted.Underline = UnderlineStyle;
        if (VerticalAlignment == RunVerticalAlignment.Superscript) formatted.Superscript = true;
        else if (VerticalAlignment == RunVerticalAlignment.Subscript) formatted.Subscript = true;
    }

    internal string TransformText(string text) => AllCaps ? text.ToUpperInvariant() : text;
}
