using PDFConverter.Layout;

namespace PDFConverter;

internal static class BorderPatterns
{
    public static UnderlineStyle From(string? style)
    {
        if (string.IsNullOrEmpty(style)) return UnderlineStyle.Single;
        var name = style.ToLowerInvariant();
        if (name.Contains("dashdotdot")) return UnderlineStyle.DotDotDash;
        if (name.Contains("dashdot") || name.Contains("dotdash")) return UnderlineStyle.DotDash;
        if (name.Contains("dash")) return UnderlineStyle.Dash;
        if (name.Contains("dot") || name == "hair") return UnderlineStyle.Dotted;
        return UnderlineStyle.Single;
    }
}
