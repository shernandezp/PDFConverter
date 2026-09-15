using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace PDFConverter;

internal static class ExcelHelpers
{
    public record ExcelImageInfo(
        byte[] Bytes,
        int? FromRow,
        int? FromCol,
        long? WidthEmu,
        long? HeightEmu,
        string? Name,
        int? ToRow = null,
        int? ToCol = null,
        double FromColumnOffsetPoints = 0);

    public record ConnectorLineInfo(int Row, int FromCol, int ToCol,
        double FromOffsetPoints, double ToOffsetPoints, double RowOffsetPoints);

    public static List<ConnectorLineInfo> GetConnectorLines(WorksheetPart wsPart)
    {
        var results = new List<ConnectorLineInfo>();
        var drawing = wsPart?.DrawingsPart?.WorksheetDrawing;
        if (drawing == null) return results;

        foreach (var anchor in drawing.Elements<Xdr.TwoCellAnchor>())
        {
            if (!anchor.Descendants<Xdr.ConnectionShape>().Any()) continue;
            var from = ReadMarker(anchor.FromMarker);
            var to = ReadMarker(anchor.ToMarker);
            if (from.Row == null || to.Row == null || from.Column == null || to.Column == null) continue;

            if (from.Row != to.Row || from.Column == to.Column) continue;

            var start = from.Column < to.Column ? from : to;
            var end = from.Column < to.Column ? to : from;
            results.Add(new ConnectorLineInfo(from.Row.Value, start.Column!.Value, end.Column!.Value,
                Units.EmuToPoints(start.Offset), Units.EmuToPoints(end.Offset),
                Units.EmuToPoints(from.RowOffset)));
        }
        return results;
    }

    public static List<ExcelImageInfo> GetImagesWithPositionFromWorksheet(WorksheetPart wsPart)
    {
        var results = new List<ExcelImageInfo>();
        var drawingsPart = wsPart?.DrawingsPart;
        var drawing = drawingsPart?.WorksheetDrawing;
        if (drawingsPart == null || drawing == null) return results;

        foreach (var anchor in drawing.Elements<Xdr.TwoCellAnchor>())
        {
            var from = ReadMarker(anchor.FromMarker);
            var to = ReadMarker(anchor.ToMarker);
            var picture = anchor.Descendants<Xdr.Picture>().FirstOrDefault();
            var extents = picture?.ShapeProperties?.Transform2D?.Extents;
            AddPicture(results, drawingsPart, picture, from, to, extents?.Cx?.Value, extents?.Cy?.Value);
        }

        foreach (var anchor in drawing.Elements<Xdr.OneCellAnchor>())
        {
            var from = ReadMarker(anchor.FromMarker);
            AddPicture(results, drawingsPart, anchor.Descendants<Xdr.Picture>().FirstOrDefault(),
                from, (null, null, 0, 0), anchor.Extent?.Cx?.Value, anchor.Extent?.Cy?.Value);
        }

        foreach (var anchor in drawing.Elements<Xdr.AbsoluteAnchor>())
        {
            AddPicture(results, drawingsPart, anchor.Descendants<Xdr.Picture>().FirstOrDefault(),
                (null, null, 0, 0), (null, null, 0, 0), anchor.Extent?.Cx?.Value, anchor.Extent?.Cy?.Value);
        }

        return results;
    }

    static void AddPicture(List<ExcelImageInfo> results, DrawingsPart drawingsPart, Xdr.Picture? picture,
        (int? Row, int? Column, long Offset, long RowOffset) from,
        (int? Row, int? Column, long Offset, long RowOffset) to,
        long? widthEmu, long? heightEmu)
    {
        var relationshipId = picture?.BlipFill?.Blip?.Embed?.Value;
        if (string.IsNullOrEmpty(relationshipId)) return;

        var bytes = WordHelpers.GetImageBytes(drawingsPart, relationshipId);
        if (bytes == null || bytes.Length == 0) return;

        var name = picture!.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
        results.Add(new ExcelImageInfo(bytes, from.Row, from.Column, widthEmu, heightEmu, name,
            to.Row, to.Column, Units.EmuToPoints(from.Offset)));
    }

    static (int? Row, int? Column, long Offset, long RowOffset) ReadMarker(Xdr.MarkerType? marker)
    {
        if (marker == null) return (null, null, 0, 0);
        return (Units.TryParseInt(marker.RowId?.Text, out var row) ? row : null,
            Units.TryParseInt(marker.ColumnId?.Text, out var column) ? column : null,
            Units.TryParseLong(marker.ColumnOffset?.Text, out var offset) ? offset : 0,
            Units.TryParseLong(marker.RowOffset?.Text, out var rowOffset) ? rowOffset : 0);
    }

    public static List<(int startRow, int startCol, int endRow, int endCol)> GetMergeCellRanges(S.Worksheet ws)
    {
        var ranges = new List<(int, int, int, int)>();
        var merges = ws.Elements<S.MergeCells>().FirstOrDefault();
        if (merges == null) return ranges;

        foreach (var merge in merges.Elements<S.MergeCell>())
        {
            var reference = merge.Reference?.Value;
            if (string.IsNullOrEmpty(reference)) continue;

            var parts = reference.Split(':');
            var (startColumn, startRow) = ParseCellReference(parts[0]);
            var (endColumn, endRow) = ParseCellReference(parts.Length > 1 ? parts[1] : parts[0]);
            if (startRow < 0 || startColumn < 0) continue;

            ranges.Add((startRow, startColumn, endRow, endColumn));
        }
        return ranges;
    }

    /// <summary>Column widths in points, padded to <paramref name="maxColumns"/> with the sheet default.</summary>
    public static List<double> GetWorksheetColumnWidths(WorksheetPart wsPart, int maxColumns = 0)
    {
        var worksheet = wsPart.Worksheet;
        if (worksheet == null) return [];

        var defaultWidth = worksheet.SheetFormatProperties?.DefaultColumnWidth?.Value
            ?? LayoutDefaults.ExcelColumnWidthChars;

        var widthByColumn = new Dictionary<uint, double>();
        foreach (var column in worksheet.Elements<S.Columns>().SelectMany(c => c.Elements<S.Column>()))
        {
            var min = column.Min?.Value ?? 1;
            var max = column.Max?.Value ?? min;
            var width = column.Hidden?.Value == true ? 0 : column.Width?.Value ?? defaultWidth;
            for (var i = min; i <= max && i - min < 16384; i++) widthByColumn[i] = width;
        }

        var lastDefined = widthByColumn.Count > 0 ? widthByColumn.Keys.Max() : 0;
        var total = (uint)Math.Max(maxColumns, (int)lastDefined);

        var widths = new List<double>((int)total);
        for (uint i = 1; i <= total; i++)
            widths.Add(ToPoints(widthByColumn.TryGetValue(i, out var width) ? width : defaultWidth));
        return widths;
    }

    // Excel stores column width as a count of digits in the default font: 7 pixels each plus 5
    // pixels of cell padding, at 96 dpi.
    static double ToPoints(double excelWidth)
    {
        if (excelWidth <= 0) return 0;
        var pixels = excelWidth * 7.0 + 5.0;
        return Math.Max(pixels * Units.PointsPerPixel, LayoutDefaults.ExcelMinColumnWidthPoints);
    }

    public static int GetColumnIndex(string? cellReference)
    {
        if (string.IsNullOrEmpty(cellReference)) return -1;

        var index = 0;
        var seen = false;
        foreach (var c in cellReference)
        {
            if (c == '$') continue;
            var upper = char.ToUpperInvariant(c);
            if (upper is < 'A' or > 'Z') break;
            index = index * 26 + (upper - 'A' + 1);
            seen = true;
        }
        return seen ? index - 1 : -1;
    }

    public static (int Column, int Row) ParseCellReference(string cellReference)
    {
        var column = GetColumnIndex(cellReference);
        var digits = new string(cellReference.Where(char.IsDigit).ToArray());
        return (column, Units.TryParseInt(digits, out var row) ? row - 1 : -1);
    }
}
