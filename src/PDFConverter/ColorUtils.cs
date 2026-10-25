using SkiaSharp;

namespace PDFConverter;

internal static class ColorUtils
{
    /// <summary>Parses "RRGGBB", "#RRGGBB" or "AARRGGBB"; "auto", empty and malformed input yield false.</summary>
    public static bool TryParse(string? value, out SKColor color)
    {
        color = SKColors.Black;
        if (string.IsNullOrWhiteSpace(value)) return false;

        var hex = value.Trim().TrimStart('#');
        if (string.Equals(hex, "auto", StringComparison.OrdinalIgnoreCase)) return false;
        if (hex.Length == 8) hex = hex[2..];
        if (hex.Length != 6) return false;

        foreach (var c in hex)
            if (!Uri.IsHexDigit(c)) return false;

        var rgb = Convert.ToUInt32(hex, 16);
        color = new SKColor((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb);
        return true;
    }
}
