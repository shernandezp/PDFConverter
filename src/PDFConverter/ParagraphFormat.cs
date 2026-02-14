using MigraDoc.DocumentObjectModel;

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
    bool HasExplicitSpacingAfter);
