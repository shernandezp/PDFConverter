namespace PDFConverter;

/// <summary>
/// Fallback measurements used when a document leaves a value unspecified.
/// Values mirror the defaults Word and Excel apply themselves.
/// </summary>
internal static class LayoutDefaults
{
    // --- Word ---
    public const double FontSizePoints = 11.0;
    public const double TabStopIntervalPoints = 36.0;
    public const double BorderWidthPoints = 0.5;
    public const double PageMarginCentimeters = 2.54;
    public const double TableColumnWidthPoints = 100.0;

    /// <summary>Word's default left and right cell margin (108 twips).</summary>
    public const double TableCellMarginPoints = 5.4;

    /// <summary>Size used for a picture whose drawing carries no extent (4 cm).</summary>
    public const double ImageSizePoints = 113.4;
    public const double FloatingImageWidthCentimeters = 4.0;
    public const double BlockImageWidthCentimeters = 16.0;
    public const double HeaderImageWidthCentimeters = 8.0;
    public const double FooterImageWidthCentimeters = 6.0;
    public const double TableCellImageWidthPoints = 100.0;

    /// <summary>Gap kept between header content and the first body line.</summary>
    public const double HeaderBodyGapPoints = 4.0;

    /// <summary>A header picture covering at least this share of the page is treated as page art.</summary>
    public const double HeaderBackgroundPageRatio = 0.6;

    /// <summary>An anchored header picture covering at least this share of the page is treated as page art.</summary>
    public const double FullPageImageRatio = 0.8;

    /// <summary>Upper bound on spacing inside table cells, which Word lays out more tightly than body text.</summary>
    public const double TableCellSpacingCapPoints = 6.0;

    /// <summary>Tolerance when matching a placed PDF image back to the picture it came from.</summary>
    public const double ImageLinkWidthTolerancePoints = 2.0;

    // --- Excel ---
    public const double ExcelRowHeightPoints = 14.5;
    public const double ExcelEmptyRowHeightPoints = 6.0;
    public const double ExcelColumnWidthChars = 8.43;
    public const double ExcelColumnWidthPoints = 48.0;
    public const double ExcelMinColumnWidthPoints = 20.0;
    public const double ExcelFontSizePoints = 10.0;
    public const double ExcelMinPageMarginInches = 0.2;
    public const double ExcelPageMarginCentimeters = 1.5;
    public const double ExcelImageWidthCentimeters = 8.0;

    /// <summary>Ratio of line height to font size, used to decide whether text fits an Excel row.</summary>
    public const double ExcelLineHeightRatio = 1.15;

    /// <summary>Width of one underscore character at <see cref="ExcelFontSizePoints"/>, used to draw connector lines.</summary>
    public const double UnderscoreWidthPoints = 4.5;
}
