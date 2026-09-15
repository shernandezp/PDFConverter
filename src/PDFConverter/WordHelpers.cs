using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MigraDoc.DocumentObjectModel;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal static class WordHelpers
{
    public static IEnumerable<OpenXmlPart> ContentParts(MainDocumentPart? mainPart)
    {
        if (mainPart == null) yield break;
        yield return mainPart;
        foreach (var header in mainPart.HeaderParts) yield return header;
        foreach (var footer in mainPart.FooterParts) yield return footer;
        if (mainPart.FootnotesPart != null) yield return mainPart.FootnotesPart;
        if (mainPart.EndnotesPart != null) yield return mainPart.EndnotesPart;
    }

    public static byte[]? GetImageBytes(OpenXmlPart? owner, string? relationshipId)
    {
        if (owner == null || string.IsNullOrEmpty(relationshipId)) return null;
        try
        {
            if (owner.GetPartById(relationshipId) is not ImagePart image) return null;
            using var stream = image.GetStream();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }
        catch
        {
            return null;
        }
    }

    public static byte[]? GetImageBytesFromWord(WordprocessingDocument doc, string? relationshipId)
    {
        foreach (var part in ContentParts(doc?.MainDocumentPart))
        {
            var bytes = GetImageBytes(part, relationshipId);
            if (bytes != null) return bytes;
        }
        return null;
    }

    public static string? GetRunFontFamily(W.RunProperties? rPr) => GetFontFamily(rPr?.RunFonts);

    static string? GetFontFamily(W.RunFonts? fonts) =>
        fonts == null ? null : fonts.Ascii?.Value ?? fonts.HighAnsi?.Value ?? fonts.ComplexScript?.Value;

    public static string? GetRunColor(W.RunProperties? rPr) => rPr?.Color?.Val?.Value;

    public static W.ParagraphProperties? GetDocDefaultsParagraphProperties(MainDocumentPart? mainPart) =>
        mainPart == null ? null : WordStyleCache.For(mainPart).DocDefaultsParagraphProperties;

    public static W.RunProperties? GetDocDefaultsRunProperties(MainDocumentPart? mainPart) =>
        mainPart == null ? null : WordStyleCache.For(mainPart).DocDefaultsRunProperties;

    public static string? GetThemeFont(MainDocumentPart? mainPart) =>
        mainPart == null ? null : WordStyleCache.For(mainPart).ThemeFont;

    public static W.RunProperties? GetStyleRunProperties(MainDocumentPart? mainPart, string styleId) =>
        mainPart == null || string.IsNullOrEmpty(styleId)
            ? null
            : WordStyleCache.For(mainPart).GetStyleRunProperties(styleId);

    public static W.ParagraphProperties? GetStyleParagraphProperties(MainDocumentPart? mainPart, string styleId) =>
        mainPart == null || string.IsNullOrEmpty(styleId)
            ? null
            : WordStyleCache.For(mainPart).GetStyleParagraphProperties(styleId);

    public static ParagraphFormat GetParagraphFormatting(W.ParagraphProperties? pPr)
    {
        if (pPr == null)
            return new ParagraphFormat(ParagraphAlignment.Left, 0, 0, 0, 0, 0, null, null, false, false);

        var alignment = pPr.Justification?.Val?.Value switch
        {
            var j when j == W.JustificationValues.Center => ParagraphAlignment.Center,
            var j when j == W.JustificationValues.Right => ParagraphAlignment.Right,
            var j when j == W.JustificationValues.Both => ParagraphAlignment.Justify,
            _ => ParagraphAlignment.Left,
        };

        double leftIndent = 0, rightIndent = 0, firstLine = 0;
        if (pPr.Indentation is { } indent)
        {
            leftIndent = Units.TwipsToPoints(indent.Left?.Value) ?? 0;
            rightIndent = Units.TwipsToPoints(indent.Right?.Value) ?? 0;
            firstLine = Units.TwipsToPoints(indent.FirstLine?.Value) ?? 0;
            var hanging = Units.TwipsToPoints(indent.Hanging?.Value);
            if (hanging.HasValue) firstLine = -hanging.Value;
        }

        double before = 0, after = 0;
        double? lineSpacing = null;
        string? lineRule = null;
        bool hasExplicitBefore = false, hasExplicitAfter = false;

        if (pPr.SpacingBetweenLines is { } spacing)
        {
            var beforePoints = Units.TwipsToPoints(spacing.Before?.Value);
            if (beforePoints.HasValue) { before = beforePoints.Value; hasExplicitBefore = true; }

            var afterPoints = Units.TwipsToPoints(spacing.After?.Value);
            if (afterPoints.HasValue) { after = afterPoints.Value; hasExplicitAfter = true; }

            if (spacing.LineRule?.Value is { } rule)
            {
                if (rule == W.LineSpacingRuleValues.Exact) lineRule = "Exact";
                else if (rule == W.LineSpacingRuleValues.AtLeast) lineRule = "AtLeast";
                else lineRule = "Auto";
            }

            if (Units.TryParseDouble(spacing.Line?.Value, out var line))
            {
                if (lineRule is "Exact" or "AtLeast")
                {
                    lineSpacing = Units.TwipsToPoints(line);
                }
                else
                {
                    lineSpacing = line / Units.TwipsPerLine;
                    lineRule ??= "Auto";
                }
            }
        }

        var pageBreakBefore = IsToggleOn(pPr.PageBreakBefore) ?? false;
        var shading = pPr.Shading?.Fill?.Value;

        return new ParagraphFormat(alignment, leftIndent, rightIndent, firstLine, before, after,
            lineSpacing, lineRule, hasExplicitBefore, hasExplicitAfter, pageBreakBefore, shading);
    }

    /// <summary>Column widths in points. Percentage widths need <paramref name="availableWidthPoints"/> to resolve.</summary>
    public static List<double> GetTableGridColumnWidths(W.Table table, double availableWidthPoints = 0)
    {
        var widths = new List<double>();

        var grid = table.GetFirstChild<W.TableGrid>();
        if (grid != null)
        {
            foreach (var column in grid.Elements<W.GridColumn>())
                widths.Add(Units.TwipsToPoints(column.Width?.Value) ?? LayoutDefaults.TableColumnWidthPoints);
            if (widths.Count > 0) return widths;
        }

        var firstRow = table.Elements<W.TableRow>().FirstOrDefault();
        if (firstRow == null) return widths;

        foreach (var cell in firstRow.Elements<W.TableCell>())
        {
            var tcPr = cell.GetFirstChild<W.TableCellProperties>();
            var cellWidth = ResolveWidth(tcPr?.GetFirstChild<W.TableCellWidth>(), availableWidthPoints);
            var gridSpan = (int)(tcPr?.GetFirstChild<W.GridSpan>()?.Val?.Value ?? 1);
            for (var i = 0; i < gridSpan; i++)
                widths.Add(cellWidth / gridSpan);
        }

        return widths;
    }

    static double ResolveWidth(W.TableWidthType? width, double availableWidthPoints)
    {
        if (!Units.TryParseDouble(width?.Width?.Value, out var value))
            return LayoutDefaults.TableColumnWidthPoints;

        var type = width!.Type?.Value;
        if (type == W.TableWidthUnitValues.Dxa) return Units.TwipsToPoints(value);
        if (type == W.TableWidthUnitValues.Pct && availableWidthPoints > 0)
            return value / Units.FiftiethsOfPercentPerWhole * availableWidthPoints;
        return LayoutDefaults.TableColumnWidthPoints;
    }

    /// <summary>Total table width in points, or null when the table is auto-sized.</summary>
    public static double? GetTableWidth(W.TableProperties? tblPr, double availableWidthPoints)
    {
        var tblW = tblPr?.TableWidth;
        if (tblW == null) return null;
        var type = tblW.Type?.Value;
        if (type == W.TableWidthUnitValues.Auto || type == W.TableWidthUnitValues.Nil) return null;
        if (!Units.TryParseDouble(tblW.Width?.Value, out var value) || value <= 0) return null;
        if (type == W.TableWidthUnitValues.Pct)
            return value / Units.FiftiethsOfPercentPerWhole * availableWidthPoints;
        return Units.TwipsToPoints(value);
    }

    static bool? IsToggleOn(W.OnOffType? element) =>
        element == null ? null : element.Val == null || element.Val.Value;

    public static RunFormat ResolveRunFormatting(MainDocumentPart? mainPart, W.Run run, W.Paragraph? paragraph)
    {
        var cache = mainPart == null ? null : WordStyleCache.For(mainPart);
        var direct = run.RunProperties;

        // Lowest to highest priority, matching the OOXML formatting cascade.
        var layers = new List<W.RunProperties?>
        {
            cache?.DocDefaultsRunProperties,
            cache?.GetStyleRunPropertiesOrNull(paragraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value),
            cache?.GetStyleRunPropertiesOrNull(direct?.RunStyle?.Val?.Value),
            direct,
        };

        string? fontFamily = null, color = null;
        double? size = null;
        bool bold = false, italic = false, allCaps = false;
        W.Underline? underline = null;
        var verticalAlignment = RunVerticalAlignment.Baseline;

        foreach (var layer in layers)
        {
            if (layer == null) continue;
            fontFamily = ResolveFont(layer.RunFonts, cache) ?? fontFamily;
            color = layer.Color?.Val?.Value ?? color;
            size = Units.HalfPointsToPoints(layer.FontSize?.Val?.Value) ?? size;
            bold = IsToggleOn(layer.Bold) ?? bold;
            italic = IsToggleOn(layer.Italic) ?? italic;
            allCaps = IsToggleOn(layer.Caps) ?? allCaps;
            if (layer.Underline != null) underline = layer.Underline;
            if (layer.VerticalTextAlignment?.Val?.Value is { } vAlign)
            {
                verticalAlignment =
                    vAlign == W.VerticalPositionValues.Superscript ? RunVerticalAlignment.Superscript :
                    vAlign == W.VerticalPositionValues.Subscript ? RunVerticalAlignment.Subscript :
                    RunVerticalAlignment.Baseline;
            }
        }

        fontFamily ??= cache?.ThemeFont;

        var boldSpecified = direct?.Bold != null
            || (cache?.GetStyleRunPropertiesOrNull(direct?.RunStyle?.Val?.Value))?.Bold != null;

        var underlineOn = underline?.Val?.Value is { } style && style != W.UnderlineValues.None;

        return new RunFormat(fontFamily, color, bold, italic, underlineOn, size, boldSpecified,
            MapUnderline(underline?.Val?.Value), verticalAlignment, allCaps);
    }

    static string? ResolveFont(W.RunFonts? fonts, WordStyleCache? cache)
    {
        var explicitFont = GetFontFamily(fonts);
        if (!string.IsNullOrEmpty(explicitFont)) return explicitFont;
        if (fonts == null || cache == null) return null;

        var theme = fonts.AsciiTheme?.Value ?? fonts.HighAnsiTheme?.Value;
        if (theme == null) return null;
        return theme == W.ThemeFontValues.MajorHighAnsi || theme == W.ThemeFontValues.MajorAscii
            ? cache.ThemeMajorFont
            : cache.ThemeFont;
    }

    static Underline MapUnderline(W.UnderlineValues? value)
    {
        if (value is not { } v) return Underline.Single;
        if (v == W.UnderlineValues.Words) return Underline.Words;
        if (v == W.UnderlineValues.Dotted || v == W.UnderlineValues.DottedHeavy) return Underline.Dotted;
        if (v == W.UnderlineValues.Dash || v == W.UnderlineValues.DashedHeavy
            || v == W.UnderlineValues.DashLong || v == W.UnderlineValues.DashLongHeavy) return Underline.Dash;
        if (v == W.UnderlineValues.DotDash || v == W.UnderlineValues.DashDotHeavy) return Underline.DotDash;
        if (v == W.UnderlineValues.DotDotDash || v == W.UnderlineValues.DashDotDotHeavy) return Underline.DotDotDash;
        return Underline.Single;
    }

    public static (string numFmt, string lvlText, int? startAt) GetNumberingLevelFormat(
        MainDocumentPart? mainPart, string? numId, int ilvl)
    {
        var level = mainPart == null
            ? NumberingLevel.Default
            : WordStyleCache.For(mainPart).GetNumberingLevel(numId, ilvl);
        return (level.Format, level.LevelText, level.StartAt);
    }

    internal readonly record struct CellMargins(double Left, double Right, double Top, double Bottom)
    {
        public static CellMargins WordDefault { get; } =
            new(LayoutDefaults.TableCellMarginPoints, LayoutDefaults.TableCellMarginPoints, 0, 0);
    }

    /// <summary>Default cell margins for a table: inline tblCellMar over the table style's, over Word's own.</summary>
    internal static CellMargins GetTableCellMargins(MainDocumentPart? mainPart, W.TableProperties? tblPr)
    {
        var margins = CellMargins.WordDefault;

        var styleId = tblPr?.GetFirstChild<W.TableStyle>()?.Val?.Value;
        if (!string.IsNullOrEmpty(styleId) && mainPart != null)
        {
            var styleMargins = WordStyleCache.For(mainPart).WalkTableStyles(styleId,
                style => style.StyleTableProperties?.GetFirstChild<W.TableCellMarginDefault>());
            margins = Merge(margins, styleMargins);
        }

        return Merge(margins, tblPr?.GetFirstChild<W.TableCellMarginDefault>());
    }

    static CellMargins Merge(CellMargins margins, W.TableCellMarginDefault? source) => source == null
        ? margins
        : new CellMargins(
            MarginEdge(source, "left", "start") ?? margins.Left,
            MarginEdge(source, "right", "end") ?? margins.Right,
            MarginEdge(source, "top") ?? margins.Top,
            MarginEdge(source, "bottom") ?? margins.Bottom);

    // tblCellMar spells its edges left/right in the original schema and start/end in the strict one,
    // and types the width differently per edge, so the attribute is read by name.
    static double? MarginEdge(W.TableCellMarginDefault source, params string[] localNames)
    {
        foreach (var name in localNames)
        {
            var edge = source.ChildElements.FirstOrDefault(c => c.LocalName == name);
            var width = edge?.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value;
            if (Units.TryParseDouble(width, out var twips)) return Units.TwipsToPoints(twips);
        }
        return null;
    }

    public static BorderInfo GetWordCellBorders(W.TableCellProperties? tcPr) => GetWordCellBorders(tcPr, null);

    public static BorderInfo GetWordCellBorders(W.TableCellProperties? tcPr, W.TableProperties? tblPr)
    {
        var top = Edge.None; var bottom = Edge.None; var left = Edge.None; var right = Edge.None;

        var tblBorders = tblPr?.GetFirstChild<W.TableBorders>();
        if (tblBorders != null)
        {
            top = Edge.From(tblBorders.TopBorder);
            bottom = Edge.From(tblBorders.BottomBorder);
            left = Edge.From(tblBorders.LeftBorder);
            right = Edge.From(tblBorders.RightBorder);

            var insideH = Edge.From(tblBorders.InsideHorizontalBorder);
            if (insideH.Width > 0)
            {
                if (top.Width == 0) top = insideH;
                if (bottom.Width == 0) bottom = insideH;
            }
            var insideV = Edge.From(tblBorders.InsideVerticalBorder);
            if (insideV.Width > 0)
            {
                if (left.Width == 0) left = insideV;
                if (right.Width == 0) right = insideV;
            }
        }

        double padTop = 0, padBottom = 0, padLeft = 0, padRight = 0;

        if (tcPr != null)
        {
            var cellBorders = tcPr.GetFirstChild<W.TableCellBorders>();
            if (cellBorders != null)
            {
                if (cellBorders.TopBorder != null) top = Edge.From(cellBorders.TopBorder);
                if (cellBorders.BottomBorder != null) bottom = Edge.From(cellBorders.BottomBorder);
                if (cellBorders.LeftBorder != null) left = Edge.From(cellBorders.LeftBorder);
                if (cellBorders.RightBorder != null) right = Edge.From(cellBorders.RightBorder);
            }

            var margins = tcPr.GetFirstChild<W.TableCellMargin>();
            if (margins != null)
            {
                padTop = Units.TwipsToPoints(margins.TopMargin?.Width?.Value) ?? 0;
                padBottom = Units.TwipsToPoints(margins.BottomMargin?.Width?.Value) ?? 0;
                padLeft = Units.TwipsToPoints(margins.LeftMargin?.Width?.Value) ?? 0;
                padRight = Units.TwipsToPoints(margins.RightMargin?.Width?.Value) ?? 0;
            }
        }

        return new BorderInfo(
            top.Width, top.Color, top.Style,
            bottom.Width, bottom.Color, bottom.Style,
            left.Width, left.Color, left.Style,
            right.Width, right.Color, right.Style,
            padTop, padBottom, padLeft, padRight);
    }

    readonly record struct Edge(double Width, string? Color, string? Style)
    {
        public static Edge None { get; } = new(0, null, null);

        public static Edge From(W.BorderType? border)
        {
            if (border == null) return None;
            if (border.Val?.Value == W.BorderValues.None || border.Val?.Value == W.BorderValues.Nil) return None;

            var width = border.Size != null && border.Size.HasValue
                ? Units.EighthPointsToPoints(border.Size.Value)
                : LayoutDefaults.BorderWidthPoints;
            var color = ColorUtils.TryParse(border.Color?.Value, out _) ? "#" + border.Color!.Value!.TrimStart('#') : null;
            return new Edge(width, color, border.Val?.InnerText);
        }
    }

    /// <summary>Table borders from the referenced table style, overridden by inline tblPr borders.</summary>
    public static W.TableBorders? ResolveTableBorders(MainDocumentPart? mainPart, W.TableProperties? tblPr)
    {
        var inline = tblPr?.GetFirstChild<W.TableBorders>();
        if (inline != null) return inline;

        var styleId = tblPr?.GetFirstChild<W.TableStyle>()?.Val?.Value;
        if (string.IsNullOrEmpty(styleId) || mainPart == null) return null;

        return WordStyleCache.For(mainPart)
            .WalkTableStyles(styleId, style => style.StyleTableProperties?.GetFirstChild<W.TableBorders>());
    }

    public static W.TableStyleProperties? GetTableStyleConditionalFormatting(
        MainDocumentPart? mainPart, W.TableProperties? tblPr, W.TableStyleOverrideValues conditionType)
    {
        var styleId = tblPr?.GetFirstChild<W.TableStyle>()?.Val?.Value;
        if (string.IsNullOrEmpty(styleId) || mainPart == null) return null;

        return WordStyleCache.For(mainPart).WalkTableStyles(styleId,
            style => style.Elements<W.TableStyleProperties>().FirstOrDefault(p => p.Type?.Value == conditionType));
    }

    public static bool IsConditionalRow(W.TableRow row, W.TableProperties? tblPr,
        int rowIndex, int totalRows, W.TableStyleOverrideValues conditionType)
    {
        var cnf = row.GetFirstChild<W.TableRowProperties>()?.GetFirstChild<W.ConditionalFormatStyle>();
        if (cnf != null)
        {
            if (conditionType == W.TableStyleOverrideValues.FirstRow
                && (cnf.FirstRow?.Value == true || CnfStyleBit(cnf, 0))) return true;
            if (conditionType == W.TableStyleOverrideValues.LastRow
                && (cnf.LastRow?.Value == true || CnfStyleBit(cnf, 1))) return true;
        }

        var tblLook = tblPr?.GetFirstChild<W.TableLook>();
        if (tblLook == null) return false;

        if (conditionType == W.TableStyleOverrideValues.FirstRow && rowIndex == 0)
            return tblLook.FirstRow?.Value == true || TblLookBit(tblLook, 0x0020);
        if (conditionType == W.TableStyleOverrideValues.LastRow && rowIndex == totalRows - 1)
            return tblLook.LastRow?.Value == true || TblLookBit(tblLook, 0x0040);
        return false;
    }

    public static bool IsConditionalColumn(W.TableCell cell, W.TableProperties? tblPr, int colIndex, int totalCols)
    {
        var cnf = cell.GetFirstChild<W.TableCellProperties>()?.GetFirstChild<W.ConditionalFormatStyle>();
        if (cnf?.FirstColumn?.Value == true || (cnf != null && CnfStyleBit(cnf, 2))) return true;

        var tblLook = tblPr?.GetFirstChild<W.TableLook>();
        return tblLook != null && colIndex == 0
            && (tblLook.FirstColumn?.Value == true || TblLookBit(tblLook, 0x0080));
    }

    /// <summary>Reads the pre-2010 hexadecimal tblLook bitmask.</summary>
    static bool TblLookBit(W.TableLook tblLook, int bitMask)
    {
        var value = tblLook.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        if (string.IsNullOrEmpty(value)) return false;
        return int.TryParse(value, System.Globalization.NumberStyles.HexNumber,
            System.Globalization.CultureInfo.InvariantCulture, out var bits) && (bits & bitMask) != 0;
    }

    /// <summary>Reads the pre-2010 cnfStyle binary string; 0=firstRow, 1=lastRow, 2=firstColumn.</summary>
    static bool CnfStyleBit(W.ConditionalFormatStyle cnf, int position)
    {
        var value = cnf.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        return !string.IsNullOrEmpty(value) && position < value.Length && value[position] == '1';
    }
}
