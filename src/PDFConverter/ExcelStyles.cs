using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace PDFConverter;

internal sealed class ExcelStyles
{
    static readonly ConditionalWeakTable<WorkbookPart, ExcelStyles> Caches = new();

    // Excel's theme colour order swaps the first two light/dark pairs relative to the theme part.
    static readonly int[] ThemeColorOrder = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];

    static readonly string[] IndexedPalette =
    [
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
        "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
        "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
        "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
        "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
        "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
    ];

    readonly S.Stylesheet? _stylesheet;
    readonly List<S.CellFormat> _cellFormats;
    readonly List<S.Fill> _fills;
    readonly List<S.Font> _fonts;
    readonly List<S.Border> _borders;
    readonly Dictionary<uint, string> _customNumberFormats = new();
    readonly List<string> _themeColors = [];
    readonly Dictionary<uint, ExcelCellStyleInfo> _resolved = new();

    ExcelStyles(WorkbookPart workbookPart)
    {
        _stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        _cellFormats = _stylesheet?.CellFormats?.Elements<S.CellFormat>().ToList() ?? [];
        _fills = _stylesheet?.Fills?.Elements<S.Fill>().ToList() ?? [];
        _fonts = _stylesheet?.Fonts?.Elements<S.Font>().ToList() ?? [];
        _borders = _stylesheet?.Borders?.Elements<S.Border>().ToList() ?? [];

        foreach (var format in _stylesheet?.NumberingFormats?.Elements<S.NumberingFormat>() ?? [])
        {
            if (format.NumberFormatId?.Value is { } id && format.FormatCode?.Value is { } code)
                _customNumberFormats[id] = code;
        }

        foreach (var color in workbookPart.ThemePart?.Theme?.ThemeElements?.ColorScheme?.Elements() ?? [])
        {
            var rgb = color.Descendants<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()?.Val?.Value
                ?? color.Descendants<DocumentFormat.OpenXml.Drawing.SystemColor>().FirstOrDefault()?.LastColor?.Value;
            _themeColors.Add(rgb ?? "000000");
        }
    }

    public static ExcelStyles For(WorkbookPart workbookPart) =>
        Caches.GetValue(workbookPart, part => new ExcelStyles(part));

    public string? GetNumberFormat(uint? numberFormatId)
    {
        if (numberFormatId is not { } id) return null;
        if (_customNumberFormats.TryGetValue(id, out var custom)) return custom;
        return BuiltInNumberFormat(id);
    }

    static string? BuiltInNumberFormat(uint id) => id switch
    {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "M/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "M/d/yyyy h:mm",
        37 => "#,##0;(#,##0)",
        38 => "#,##0;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        _ => null,
    };

    public ExcelCellStyleInfo GetCellStyle(uint? styleIndex)
    {
        if (styleIndex is not { } index || index >= _cellFormats.Count) return ExcelCellStyleInfo.Empty;
        if (_resolved.TryGetValue(index, out var cached)) return cached;

        var format = _cellFormats[(int)index];
        var (fontFamily, fontSize, fontColor, bold, italic) = ReadFont(format.FontId?.Value);

        var resolved = new ExcelCellStyleInfo(
            format.Alignment?.Horizontal?.InnerText,
            format.Alignment?.Vertical?.InnerText,
            ReadFill(format.FillId?.Value),
            format.NumberFormatId?.Value,
            ReadBorders(format.BorderId?.Value),
            fontFamily, fontSize, fontColor, bold, italic);

        _resolved[index] = resolved;
        return resolved;
    }

    string? ReadFill(uint? fillId)
    {
        if (fillId is not { } id || id >= _fills.Count) return null;
        var pattern = _fills[(int)id].PatternFill;
        if (pattern == null || pattern.PatternType?.Value == S.PatternValues.None) return null;
        return ResolveColor(pattern.ForegroundColor) ?? ResolveColor(pattern.BackgroundColor);
    }

    (string? Family, double? Size, string? Color, bool Bold, bool Italic) ReadFont(uint? fontId)
    {
        if (fontId is not { } id || id >= _fonts.Count) return (null, null, null, false, false);

        var font = _fonts[(int)id];
        var color = ResolveColor(font.Color)?.TrimStart('#');
        return (font.FontName?.Val?.Value, font.FontSize?.Val?.Value, color,
            font.Bold != null, font.Italic != null);
    }

    BorderInfo ReadBorders(uint? borderId)
    {
        if (borderId is not { } id || id >= _borders.Count) return BorderInfo.Empty;

        var border = _borders[(int)id];
        var (topWidth, topColor, topStyle) = ReadEdge(border.TopBorder);
        var (bottomWidth, bottomColor, bottomStyle) = ReadEdge(border.BottomBorder);
        var (leftWidth, leftColor, leftStyle) = ReadEdge(border.LeftBorder);
        var (rightWidth, rightColor, rightStyle) = ReadEdge(border.RightBorder);

        return new BorderInfo(topWidth, topColor, topStyle, bottomWidth, bottomColor, bottomStyle,
            leftWidth, leftColor, leftStyle, rightWidth, rightColor, rightStyle);
    }

    (double Width, string? Color, string? Style) ReadEdge(S.BorderPropertiesType? border)
    {
        if (border?.Style == null || border.Style.Value == S.BorderStyleValues.None) return (0, null, null);
        return (BorderWidth(border.Style.Value), ResolveColor(border.Color), border.Style.InnerText);
    }

    static double BorderWidth(S.BorderStyleValues style)
    {
        if (style == S.BorderStyleValues.Hair) return 0.25;
        if (style == S.BorderStyleValues.Double) return 1.5;
        if (style == S.BorderStyleValues.Thick) return 2.0;
        if (style == S.BorderStyleValues.Medium || style == S.BorderStyleValues.MediumDashed
            || style == S.BorderStyleValues.MediumDashDot || style == S.BorderStyleValues.MediumDashDotDot
            || style == S.BorderStyleValues.SlantDashDot) return 1.0;
        return LayoutDefaults.BorderWidthPoints;
    }

    string? ResolveColor(S.ColorType? color)
    {
        if (color == null || color.Auto?.Value == true) return null;

        var rgb = color.Rgb?.Value;
        if (!string.IsNullOrEmpty(rgb)) return "#" + (rgb.Length == 8 ? rgb[2..] : rgb);

        if (color.Indexed?.Value is { } indexed && indexed < IndexedPalette.Length)
            return "#" + IndexedPalette[indexed];

        if (color.Theme?.Value is not { } theme) return null;
        var schemeIndex = theme < ThemeColorOrder.Length ? ThemeColorOrder[theme] : (int)theme;
        if (schemeIndex >= _themeColors.Count) return null;

        return "#" + ApplyTint(_themeColors[schemeIndex], color.Tint?.Value ?? 0);
    }

    static string ApplyTint(string rgb, double tint)
    {
        if (tint == 0 || rgb.Length != 6) return rgb;

        Span<char> result = stackalloc char[6];
        for (var channel = 0; channel < 3; channel++)
        {
            var value = Convert.ToInt32(rgb.Substring(channel * 2, 2), 16);
            var tinted = tint > 0 ? value * (1 - tint) + 255 * tint : value * (1 + tint);
            var clamped = (int)Math.Round(Math.Clamp(tinted, 0, 255));
            clamped.ToString("X2").AsSpan().CopyTo(result[(channel * 2)..]);
        }
        return new string(result);
    }
}
