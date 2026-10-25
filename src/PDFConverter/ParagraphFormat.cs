using PDFConverter.Layout;

namespace PDFConverter;

internal sealed record ParagraphFormat(
    ParagraphAlignment Alignment,
    double LeftIndent,
    double RightIndent,
    double FirstLineIndent,
    double SpacingBefore,
    double SpacingAfter,
    double? LineSpacing,
    string? LineRule,
    bool HasExplicitSpacingBefore,
    bool HasExplicitSpacingAfter,
    bool PageBreakBefore = false,
    string? ShadingColor = null,
    BorderInfo? Borders = null,
    bool? KeepNext = null,
    bool? KeepLines = null,
    bool? WidowControl = null,
    bool Bidi = false);
