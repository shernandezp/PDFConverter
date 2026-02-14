using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace PDFConverter;

internal static class ExcelHelpers
{
    /// <summary>
    /// Image info record containing bytes and position information
    /// </summary>
    public record ExcelImageInfo(
        byte[] Bytes,
        int? FromRow,
        int? FromCol,
        long? WidthEmu,
        long? HeightEmu,
        string? Name,
        int? ToRow = null,
        int? ToCol = null);

    /// <summary>
    /// Horizontal connector line info (used for signature lines etc.)
    /// </summary>
    public record ConnectorLineInfo(int Row, int FromCol, int ToCol);

    /// <summary>
    /// Extract horizontal connector lines from worksheet drawing.
    /// </summary>
    public static List<ConnectorLineInfo> GetConnectorLines(WorksheetPart wsPart)
    {
        var results = new List<ConnectorLineInfo>();
        if (wsPart?.DrawingsPart?.WorksheetDrawing == null) return results;

        foreach (var anchor in wsPart.DrawingsPart.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>())
        {
            try
            {
                if (anchor.Descendants<Xdr.ConnectionShape>().Any())
                {
                    var from = anchor.FromMarker;
                    var to = anchor.ToMarker;
                    if (from == null || to == null) continue;

                    int.TryParse(from.RowId?.Text, out var fromRow);
                    int.TryParse(to.RowId?.Text, out var toRow);
                    int.TryParse(from.ColumnId?.Text, out var fromCol);
                    int.TryParse(to.ColumnId?.Text, out var toCol);

                    // Only horizontal lines (same row)
                    if (fromRow == toRow && fromCol != toCol)
                        results.Add(new ConnectorLineInfo(fromRow, Math.Min(fromCol, toCol), Math.Max(fromCol, toCol)));
                }
            }
            catch { }
        }
        return results;
    }

    /// <summary>
    /// Get images from worksheet with position information
    /// </summary>
    public static IEnumerable<ExcelImageInfo> GetImagesWithPositionFromWorksheet(WorksheetPart wsPart)
    {
        var results = new List<ExcelImageInfo>();
        if (wsPart?.DrawingsPart == null) return results;

        var drawingsPart = wsPart.DrawingsPart;

        // Get the worksheet drawing
        var wsDrawing = drawingsPart.WorksheetDrawing;
        if (wsDrawing == null) return results;

        // Process TwoCellAnchor elements (most common for images)
        foreach (var anchor in wsDrawing.Elements<Xdr.TwoCellAnchor>())
        {
            try
            {
                var fromMarker = anchor.FromMarker;
                var toMarker = anchor.ToMarker;

                int? fromRow = null, fromCol = null;
                if (fromMarker != null)
                {
                    if (int.TryParse(fromMarker.RowId?.Text, out var r)) fromRow = r;
                    if (int.TryParse(fromMarker.ColumnId?.Text, out var c)) fromCol = c;
                }

                int? toRow = null, toCol = null;
                if (toMarker != null)
                {
                    if (int.TryParse(toMarker.RowId?.Text, out var r2)) toRow = r2;
                    if (int.TryParse(toMarker.ColumnId?.Text, out var c2)) toCol = c2;
                }

                // Calculate dimensions from to/from markers
                long? widthEmu = null, heightEmu = null;

                // Find the picture element
                var picture = anchor.Descendants<Xdr.Picture>().FirstOrDefault();
                if (picture != null)
                {
                    var blipFill = picture.BlipFill;
                    var blip = blipFill?.Blip;
                    var embed = blip?.Embed?.Value;

                    if (!string.IsNullOrEmpty(embed))
                    {
                        var imgBytes = GetImageBytesFromRelationship(drawingsPart, embed);
                        if (imgBytes != null && imgBytes.Length > 0)
                        {
                            // Try to get extent from picture properties
                            var spPr = picture.ShapeProperties;
                            var xfrm = spPr?.Transform2D;
                            if (xfrm?.Extents != null)
                            {
                                widthEmu = xfrm.Extents.Cx?.Value;
                                heightEmu = xfrm.Extents.Cy?.Value;
                            }

                            var name = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
                            results.Add(new ExcelImageInfo(imgBytes, fromRow, fromCol, widthEmu, heightEmu, name, toRow, toCol));
                        }
                    }
                }
            }
            catch { }
        }

        // Process OneCellAnchor elements
        foreach (var anchor in wsDrawing.Elements<Xdr.OneCellAnchor>())
        {
            try
            {
                var fromMarker = anchor.FromMarker;

                int? fromRow = null, fromCol = null;
                if (fromMarker != null)
                {
                    if (int.TryParse(fromMarker.RowId?.Text, out var r)) fromRow = r;
                    if (int.TryParse(fromMarker.ColumnId?.Text, out var c)) fromCol = c;
                }

                long? widthEmu = null, heightEmu = null;
                var extent = anchor.Extent;
                if (extent != null)
                {
                    widthEmu = extent.Cx?.Value;
                    heightEmu = extent.Cy?.Value;
                }

                var picture = anchor.Descendants<Xdr.Picture>().FirstOrDefault();
                if (picture != null)
                {
                    var blipFill = picture.BlipFill;
                    var blip = blipFill?.Blip;
                    var embed = blip?.Embed?.Value;

                    if (!string.IsNullOrEmpty(embed))
                    {
                        var imgBytes = GetImageBytesFromRelationship(drawingsPart, embed);
                        if (imgBytes != null && imgBytes.Length > 0)
                        {
                            var name = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
                            results.Add(new ExcelImageInfo(imgBytes, fromRow, fromCol, widthEmu, heightEmu, name));
                        }
                    }
                }
            }
            catch { }
        }

        // Process AbsoluteAnchor elements (less common)
        foreach (var anchor in wsDrawing.Elements<Xdr.AbsoluteAnchor>())
        {
            try
            {
                long? widthEmu = null, heightEmu = null;
                var extent = anchor.Extent;
                if (extent != null)
                {
                    widthEmu = extent.Cx?.Value;
                    heightEmu = extent.Cy?.Value;
                }

                var picture = anchor.Descendants<Xdr.Picture>().FirstOrDefault();
                if (picture != null)
                {
                    var blipFill = picture.BlipFill;
                    var blip = blipFill?.Blip;
                    var embed = blip?.Embed?.Value;

                    if (!string.IsNullOrEmpty(embed))
                    {
                        var imgBytes = GetImageBytesFromRelationship(drawingsPart, embed);
                        if (imgBytes != null && imgBytes.Length > 0)
                        {
                            var name = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
                            results.Add(new ExcelImageInfo(imgBytes, null, null, widthEmu, heightEmu, name));
                        }
                    }
                }
            }
            catch { }
        }

        // Fallback: If no images found via anchors, try ImageParts directly
        if (results.Count == 0)
        {
            foreach (var imagePart in drawingsPart.ImageParts)
            {
                try
                {
                    using var s = imagePart.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    var bytes = ms.ToArray();
                    if (bytes.Length > 0)
                    {
                        results.Add(new ExcelImageInfo(bytes, null, null, null, null, null));
                    }
                }
                catch { }
            }
        }

        return results;
    }

    private static byte[]? GetImageBytesFromRelationship(DrawingsPart drawingsPart, string relationshipId)
    {
        try
        {
            var part = drawingsPart.GetPartById(relationshipId);
            if (part is ImagePart imagePart)
            {
                using var s = imagePart.GetStream();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                return ms.ToArray();
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Legacy method - returns only bytes for backward compatibility
    /// </summary>
    public static IEnumerable<byte[]> GetImagesFromWorksheet(WorksheetPart wsPart)
    {
        return GetImagesWithPositionFromWorksheet(wsPart)
            .Where(i => i.Bytes != null && i.Bytes.Length > 0)
            .Select(i => i.Bytes);
    }

    public static List<(int startRow, int startCol, int endRow, int endCol)> GetMergeCellRanges(S.Worksheet ws)
    {
        var res = new List<(int, int, int, int)>();
        var merges = ws.Elements<S.MergeCells>().FirstOrDefault();
        if (merges == null) return res;
        foreach (var m in merges.Elements<S.MergeCell>())
        {
            var val = m.Reference?.Value;
            if (string.IsNullOrEmpty(val)) continue;
            var parts = val.Split(':');
            var start = parts[0];
            var end = parts.Length > 1 ? parts[1] : parts[0];
            var (sCol, sRow) = ParseCellReference(start);
            var (eCol, eRow) = ParseCellReference(end);
            res.Add((sRow, sCol, eRow, eCol));
        }
        return res;
    }

    /// <summary>
    /// Get column widths for a worksheet. Returns widths in points.
    /// If no column definitions exist, returns an empty list (caller should use defaults).
    /// </summary>
    public static List<double> GetWorksheetColumnWidths(WorksheetPart wsPart, int maxColumns = 0)
    {
        var res = new List<double>();
        var cols = wsPart.Worksheet.Elements<S.Columns>().FirstOrDefault();
        
        // Default width in Excel character units (typically 8.43 for Calibri 11)
        const double defaultExcelWidth = 8.43;
        
        if (cols == null)
        {
            // No column definitions - return list with default widths for requested columns
            for (int i = 0; i < maxColumns; i++)
            {
                res.Add(ConvertExcelWidthToPoints(defaultExcelWidth));
            }
            return res;
        }

        // Build a dictionary of column index -> width
        var colWidthMap = new Dictionary<uint, double>();
        
        foreach (var c in cols.Elements<S.Column>())
        {
            var min = c.Min?.Value ?? 1;
            var max = c.Max?.Value ?? min;
            var width = c.Width?.Value ?? defaultExcelWidth;

            for (uint i = min; i <= max; i++)
            {
                colWidthMap[i] = width;
            }
        }

        // Determine how many columns we need
        uint maxDefinedCol = colWidthMap.Count > 0 ? colWidthMap.Keys.Max() : 0;
        uint totalColumns = (uint)Math.Max(maxColumns, (int)maxDefinedCol);

        // Build the result list with proper widths for each column (1-indexed in Excel)
        for (uint i = 1; i <= totalColumns; i++)
        {
            double excelWidth = colWidthMap.TryGetValue(i, out var w) ? w : defaultExcelWidth;
            res.Add(ConvertExcelWidthToPoints(excelWidth));
        }
        
        return res;
    }

    /// <summary>
    /// Convert Excel column width (in character units) to points.
    /// Excel column width is measured in characters of the default font (typically Calibri 11pt).
    /// </summary>
    private static double ConvertExcelWidthToPoints(double excelWidth)
    {
        if (excelWidth <= 0) return 30; // minimum column width
        
        // Excel column width formula:
        // The width value represents the number of characters of the default font that fit in a cell.
        // For Calibri 11pt, one character is approximately 7 pixels wide.
        // Excel adds 5 pixels of padding/margin to each column.
        // At 96 DPI: 1 pixel = 0.75 points (72 points / 96 pixels)
        //
        // width_in_pixels = (excelWidth * 7) + 5
        // width_in_points = width_in_pixels * 0.75
        //
        // For default width of 8.43: (8.43 * 7 + 5) * 0.75 = 48 points ? 1.7 cm
        
        double widthInPixels = (excelWidth * 7.0) + 5.0;
        double widthInPoints = widthInPixels * 0.75;
        
        return Math.Max(widthInPoints, 20); // minimum 20 points
    }

    public static string? GetNumberFormatString(WorkbookPart wbPart, uint? numFmtId)
    {
        if (numFmtId == null) return null;
        var id = (int)numFmtId.Value;

        // Check custom number formats first
        var stylesPart = wbPart.WorkbookStylesPart;
        if (stylesPart?.Stylesheet?.NumberingFormats != null)
        {
            var customFmt = stylesPart.Stylesheet.NumberingFormats
                .Elements<S.NumberingFormat>()
                .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId);
            if (customFmt?.FormatCode?.Value != null)
            {
                return ConvertExcelFormatToNet(customFmt.FormatCode.Value);
            }
        }

        // Built-in formats
        return id switch
        {
            0 => "G",
            1 => "0",
            2 => "0.00",
            3 => "#,##0",
            4 => "#,##0.00",
            9 => "0%",
            10 => "0.00%",
            11 => "0.00E+00",
            12 => "# ?/?",
            13 => "# ??/??",
            14 => "MM/dd/yyyy",
            15 => "d-MMM-yy",
            16 => "d-MMM",
            17 => "MMM-yy",
            18 => "h:mm tt",
            19 => "h:mm:ss tt",
            20 => "H:mm",
            21 => "H:mm:ss",
            22 => "M/d/yyyy H:mm",
            37 => "#,##0 ;(#,##0)",
            38 => "#,##0 ;[Red](#,##0)",
            39 => "#,##0.00;(#,##0.00)",
            40 => "#,##0.00;[Red](#,##0.00)",
            45 => "mm:ss",
            46 => "[h]:mm:ss",
            47 => "mmss.0",
            48 => "##0.0E+0",
            49 => "@",
            _ => null,
        };
    }

    private static string ConvertExcelFormatToNet(string excelFormat)
    {
        // Basic conversion of Excel format codes to .NET format strings
        var result = excelFormat
            .Replace("yyyy", "yyyy")
            .Replace("yy", "yy")
            .Replace("mmmm", "MMMM")
            .Replace("mmm", "MMM")
            .Replace("mm", "MM")
            .Replace("dddd", "dddd")
            .Replace("ddd", "ddd")
            .Replace("dd", "dd")
            .Replace("d", "d")
            .Replace("hh", "HH")
            .Replace("h", "H")
            .Replace("ss", "ss")
            .Replace("AM/PM", "tt")
            .Replace("am/pm", "tt");

        return result;
    }

    public static ExcelCellStyleInfo GetCellStyleInfo(WorkbookPart wbPart, uint? styleIndex)
    {
        if (styleIndex == null) return ExcelCellStyleInfo.Empty;
        var stylesPart = wbPart.WorkbookStylesPart;
        if (stylesPart == null) return ExcelCellStyleInfo.Empty;
        var stylesheet = stylesPart.Stylesheet;
        if (stylesheet == null) return ExcelCellStyleInfo.Empty;

        var cfList = stylesheet.CellFormats?.Elements<S.CellFormat>().ToList();
        if (cfList == null || cfList.Count <= styleIndex) return ExcelCellStyleInfo.Empty;
        var cf = cfList[(int)styleIndex.Value];

        string? hAlignment = null;
        string? vAlignment = null;
        if (cf.Alignment != null)
        {
            if (cf.Alignment.Horizontal != null)
                hAlignment = cf.Alignment.Horizontal.InnerText;
            if (cf.Alignment.Vertical != null)
                vAlignment = cf.Alignment.Vertical.InnerText;
        }

        uint? numFmtId = cf.NumberFormatId?.Value;

        string? fillColor = null;
        if (cf.FillId != null && stylesheet.Fills != null)
        {
            var fills = stylesheet.Fills.Elements<S.Fill>().ToList();
            var fid = (int)cf.FillId.Value;
            if (fid < fills.Count)
            {
                var fill = fills[fid];
                var patternFill = fill.PatternFill;
                if (patternFill != null)
                {
                    var fg = patternFill.ForegroundColor?.Rgb?.Value;
                    if (!string.IsNullOrEmpty(fg))
                    {
                        if (fg.Length == 8) fg = fg[2..];
                        fillColor = "#" + fg;
                    }
                    if (string.IsNullOrEmpty(fillColor))
                    {
                        var bg = patternFill.BackgroundColor?.Rgb?.Value;
                        if (!string.IsNullOrEmpty(bg))
                        {
                            if (bg.Length == 8) bg = bg[2..];
                            fillColor = "#" + bg;
                        }
                    }
                }
            }
        }

        // Font information
        string? fontFamily = null;
        double? fontSize = null;
        string? fontColor = null;
        bool bold = false;
        bool italic = false;
        if (cf.FontId != null && stylesheet.Fonts != null)
        {
            var fonts = stylesheet.Fonts.Elements<S.Font>().ToList();
            var fid = (int)cf.FontId.Value;
            if (fid < fonts.Count)
            {
                var font = fonts[fid];
                fontFamily = font.FontName?.Val?.Value;
                if (font.FontSize?.Val != null)
                    fontSize = font.FontSize.Val.Value;
                var fc = font.Color?.Rgb?.Value;
                if (!string.IsNullOrEmpty(fc))
                {
                    if (fc.Length == 8) fc = fc[2..];
                    fontColor = fc;
                }
                bold = font.Bold != null;
                italic = font.Italic != null;
            }
        }

        var borders = BorderInfo.Empty;
        if (cf.BorderId != null && stylesheet.Borders != null)
        {
            var bordersList = stylesheet.Borders.Elements<S.Border>().ToList();
            var bid = (int)cf.BorderId.Value;
            if (bid < bordersList.Count)
            {
                var b = bordersList[bid];
                ReadExcelBorderEdge(b.TopBorder, out var tW, out var tC, out var tS);
                ReadExcelBorderEdge(b.BottomBorder, out var bW, out var bC, out var bS);
                ReadExcelBorderEdge(b.LeftBorder, out var lW, out var lC, out var lS);
                ReadExcelBorderEdge(b.RightBorder, out var rW, out var rC, out var rS);
                borders = new BorderInfo(tW, tC, tS, bW, bC, bS, lW, lC, lS, rW, rC, rS);
            }
        }

        return new ExcelCellStyleInfo(hAlignment, vAlignment, fillColor, numFmtId, borders, fontFamily, fontSize, fontColor, bold, italic);
    }

    private static void ReadExcelBorderEdge(S.BorderPropertiesType? border, out double width, out string? color, out string? style)
    {
        width = 0;
        color = null;
        style = null;
        if (border?.Style == null || border.Style.Value == S.BorderStyleValues.None) return;
        width = GetBorderWidth(border.Style.Value);
        style = border.Style.Value.ToString();
        var c = border.Color?.Rgb?.Value;
        if (!string.IsNullOrEmpty(c))
        {
            if (c.Length == 8) c = c[2..];
            color = "#" + c;
        }
    }

    private static double GetBorderWidth(S.BorderStyleValues style)
    {
        if (style == S.BorderStyleValues.Hair) return 0.25;
        if (style == S.BorderStyleValues.Thin) return 0.5;
        if (style == S.BorderStyleValues.Dotted) return 0.5;
        if (style == S.BorderStyleValues.Dashed) return 0.5;
        if (style == S.BorderStyleValues.DashDot) return 0.5;
        if (style == S.BorderStyleValues.DashDotDot) return 0.5;
        if (style == S.BorderStyleValues.Medium) return 1.0;
        if (style == S.BorderStyleValues.MediumDashed) return 1.0;
        if (style == S.BorderStyleValues.MediumDashDot) return 1.0;
        if (style == S.BorderStyleValues.MediumDashDotDot) return 1.0;
        if (style == S.BorderStyleValues.SlantDashDot) return 1.0;
        if (style == S.BorderStyleValues.Double) return 1.5;
        if (style == S.BorderStyleValues.Thick) return 2.0;
        return 0.5;
    }

    internal static int GetColumnIndexFromName(string name)
    {
        int index = 0;
        foreach (char c in name)
        {
            index = index * 26 + (c - 'A' + 1);
        }
        return index - 1;
    }

    internal static (int col, int row) ParseCellReference(string cellRef)
    {
        var letters = new string(cellRef.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(cellRef.SkipWhile(char.IsLetter).ToArray());
        int col = GetColumnIndexFromName(letters);
        int.TryParse(digits, out int row);
        return (col, row - 1);
    }
}
