using PDFConverter.Layout;

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
    UnderlineStyle UnderlineStyle = UnderlineStyle.Single,
    RunVerticalAlignment VerticalAlignment = RunVerticalAlignment.Baseline,
    bool AllCaps = false,
    StrikeStyle Strike = StrikeStyle.None,
    string? Highlight = null,
    bool Hidden = false,
    double LetterSpacing = 0,
    double WidthScale = 1)
{
    internal TextStyle ToTextStyle(TextStyle defaults) => new(
        string.IsNullOrEmpty(FontFamily) ? defaults.FontFamily : FontFamily,
        Size ?? defaults.Size,
        Bold,
        Italic,
        ColorUtils.TryParse(Color, out var color) ? color : defaults.Color,
        Underline ? UnderlineStyle : null,
        VerticalAlignment)
    {
        Strike = Strike,
        Highlight = ColorUtils.TryParse(Highlight, out var highlight) ? highlight : null,
        LetterSpacing = LetterSpacing,
        WidthScale = WidthScale > 0 ? WidthScale : 1,
    };

    internal string TransformText(string text) => AllCaps ? text.ToUpperInvariant() : text;
}
