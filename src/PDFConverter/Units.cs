using System.Globalization;

namespace PDFConverter;

/// <summary>
/// Conversions between OpenXML measurement units and PDF points.
/// </summary>
internal static class Units
{
    public const double EmuPerPoint = 12700.0;
    public const double TwipsPerPoint = 20.0;
    public const double EighthPointsPerPoint = 8.0;
    public const double HalfPointsPerPoint = 2.0;
    public const double TwipsPerLine = 240.0;
    public const double PointsPerPixel = 0.75;

    /// <summary>Fiftieths of a percent, the unit used by w:tblW and w:tcW when type="pct".</summary>
    public const double FiftiethsOfPercentPerWhole = 5000.0;

    /// <summary>Thousandths of a percent, the unit used by a:srcRect crop offsets.</summary>
    public const double ThousandthsOfPercentPerWhole = 100000.0;

    public static double EmuToPoints(long emu) => emu / EmuPerPoint;

    public static double TwipsToPoints(double twips) => twips / TwipsPerPoint;

    public static double EighthPointsToPoints(double eighthPoints) => eighthPoints / EighthPointsPerPoint;

    public static double HalfPointsToPoints(double halfPoints) => halfPoints / HalfPointsPerPoint;

    /// <summary>
    /// Parses a number the way OpenXML writes it: invariant culture, no locale-specific separators.
    /// </summary>
    public static bool TryParseDouble(string? text, out double value) =>
        double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);

    public static bool TryParseInt(string? text, out int value) =>
        int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

    public static bool TryParseLong(string? text, out long value) =>
        long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

    public static double? TwipsToPoints(string? twips) =>
        TryParseDouble(twips, out var value) ? TwipsToPoints(value) : null;

    public static double? HalfPointsToPoints(string? halfPoints) =>
        TryParseDouble(halfPoints, out var value) ? HalfPointsToPoints(value) : null;
}
