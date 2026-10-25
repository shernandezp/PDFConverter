using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PDFConverter.Layout;
using SkiaSharp;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal static class WordDocumentBuilder
{
    public static Document Build(WordprocessingDocument word, ImageStore images)
    {
        var ctx = new WordRenderContext(word, images);
        var body = ctx.MainPart.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        var properties = word.PackageProperties;
        var document = new Document
        {
            Title = properties.Title, Author = properties.Creator, Subject = properties.Subject, Keywords = properties.Keywords,
        };

        var elements = Flatten(body).ToList();
        var sectionProperties = elements.OfType<W.Paragraph>()
            .Select(p => p.ParagraphProperties?.SectionProperties)
            .Where(s => s != null)
            .ToList();
        sectionProperties.Add(body.Elements<W.SectionProperties>().LastOrDefault());

        var sections = new SectionCursor(ctx, document, sectionProperties);
        var section = sections.Next();
        foreach (var element in elements)
        {
            if (element is W.Paragraph paragraph)
            {
                RenderBodyParagraph(ctx, section, paragraph);
                if (paragraph.ParagraphProperties?.SectionProperties != null && sections.HasNext)
                    section = sections.Next();
            }
            else if (element is W.Table table)
            {
                WordTableRenderer.RenderTable(ctx, section, table);
            }
        }
        return document;
    }

    // Each section carries its own page setup; headers and footers are inherited from the previous
    // section when the section break does not reference any of its own.
    sealed class SectionCursor(WordRenderContext ctx, Document document, List<W.SectionProperties?> properties)
    {
        int _index;
        W.SectionProperties? _headerSource;
        W.SectionProperties? _footerSource;

        public bool HasNext => _index < properties.Count;

        public Section Next()
        {
            var sectionProperties = properties[_index++];
            if (sectionProperties?.Elements<W.HeaderReference>().Any() == true) _headerSource = sectionProperties;
            if (sectionProperties?.Elements<W.FooterReference>().Any() == true) _footerSource = sectionProperties;

            var section = document.AddSection();
            ApplyPageSetup(section, sectionProperties);
            RenderHeaders(ctx, section, _headerSource);
            RenderFooters(ctx, section, _footerSource);
            return section;
        }
    }

    static void ApplyPageSetup(Section section, W.SectionProperties? sectionProperties)
    {
        section.LeftMargin = section.RightMargin = LayoutDefaults.PageMarginCentimeters / 2.54 * 72;

        var pageSize = sectionProperties?.GetFirstChild<W.PageSize>();
        if (pageSize?.Width is { HasValue: true } width)
            section.PageWidth = Units.TwipsToPoints(width.Value);
        if (pageSize?.Height is { HasValue: true } height)
            section.PageHeight = Units.TwipsToPoints(height.Value);

        var margin = sectionProperties?.GetFirstChild<W.PageMargin>();
        if (margin == null) return;

        if (margin.Top is { HasValue: true } top) section.TopMargin = Units.TwipsToPoints(top.Value);
        if (margin.Bottom is { HasValue: true } bottom) section.BottomMargin = Units.TwipsToPoints(bottom.Value);
        if (margin.Left is { HasValue: true } left) section.LeftMargin = Units.TwipsToPoints(left.Value);
        if (margin.Right is { HasValue: true } right) section.RightMargin = Units.TwipsToPoints(right.Value);
        if (margin.Header is { HasValue: true } header) section.HeaderDistance = Units.TwipsToPoints(header.Value);
        if (margin.Footer is { HasValue: true } footer) section.FooterDistance = Units.TwipsToPoints(footer.Value);
    }

    static void RenderHeaders(WordRenderContext ctx, Section section, W.SectionProperties? sectionProperties)
    {
        foreach (var reference in sectionProperties?.Elements<W.HeaderReference>() ?? [])
        {
            var id = reference.Id?.Value;
            if (string.IsNullOrEmpty(id)) continue;
            if (ctx.MainPart.GetPartById(id) is not HeaderPart { Header: not null } part) continue;

            var container = Container(section, reference.Type?.InnerText, isHeader: true);
            RenderHeaderPart(ctx, section, part.Header, container);
        }
    }

    static void RenderFooters(WordRenderContext ctx, Section section, W.SectionProperties? sectionProperties)
    {
        foreach (var reference in sectionProperties?.Elements<W.FooterReference>() ?? [])
        {
            var id = reference.Id?.Value;
            if (string.IsNullOrEmpty(id)) continue;
            if (ctx.MainPart.GetPartById(id) is not FooterPart { Footer: not null } part) continue;

            var container = Container(section, reference.Type?.InnerText, isHeader: false);
            foreach (var paragraph in part.Footer.Elements<W.Paragraph>())
                RenderHeaderFooterParagraph(ctx, paragraph, container);
        }
    }

    static HeaderFooter Container(Section section, string? referenceType, bool isHeader)
    {
        if (string.Equals(referenceType, "First", StringComparison.OrdinalIgnoreCase))
        {
            section.DifferentFirstPage = true;
            return isHeader ? section.FirstPageHeader : section.FirstPageFooter;
        }
        if (string.Equals(referenceType, "Even", StringComparison.OrdinalIgnoreCase))
        {
            section.OddAndEvenPages = true;
            return isHeader ? section.EvenPageHeader : section.EvenPageFooter;
        }
        return isHeader ? section.PrimaryHeader : section.PrimaryFooter;
    }

    static void RenderHeaderPart(WordRenderContext ctx, Section section, W.Header header, HeaderFooter container)
    {
        var pageWidth = section.PageWidth;
        var pageHeight = section.PageHeight;
        double tallestImage = 0;

        bool IsArt(WordImage image) => IsPageArt(image, pageWidth, pageHeight);

        foreach (var paragraph in header.Elements<W.Paragraph>())
        {
            foreach (var image in ctx.ImagesIn(paragraph))
            {
                if (IsArt(image)) CollectBackground(ctx, section, image);
                else tallestImage = Math.Max(tallestImage, image.HeightPoints);
            }

            RenderHeaderFooterParagraph(ctx, paragraph, container,
                new InlineOptions(MaxImageWidthPoints: pageWidth, SkipImage: IsArt));
        }

        ReserveHeaderSpace(section, tallestImage);
    }

    // A header picture sitting behind the text, or covering most of the page, is page art rather
    // than header content: Word paints it under everything.
    static bool IsPageArt(WordImage image, double pageWidth, double pageHeight)
    {
        if (image.IsBackground) return true;
        if (!image.IsAnchor || !image.ExtentCxEmu.HasValue || !image.ExtentCyEmu.HasValue) return false;
        return image.WidthPoints > pageWidth * LayoutDefaults.FullPageImageRatio
            && image.HeightPoints > pageHeight * LayoutDefaults.FullPageImageRatio;
    }

    static void CollectBackground(WordRenderContext ctx, Section section, WordImage image)
    {
        var source = ctx.MaterializeImage(image);
        if (source == null) return;
        if (section.Backgrounds.All(art => !ReferenceEquals(art.Image, source)))
            section.Backgrounds.Add(new PageArt(source, image.WidthPoints, image.HeightPoints));

        var pageHeight = section.PageHeight;
        if (!image.ExtentCyEmu.HasValue || pageHeight <= 0) return;

        // A band across the top is header art, not a full-page backdrop: keep body text clear of it.
        var imageHeight = image.HeightPoints;
        if (imageHeight / pageHeight >= LayoutDefaults.HeaderBackgroundPageRatio) return;

        var available = Math.Max(0, pageHeight - section.TopMargin - section.BottomMargin - 10.0);
        var headerHeight = Math.Min(imageHeight, available);
        if (section.HeaderDistance < headerHeight) section.HeaderDistance = headerHeight;
    }

    static void ReserveHeaderSpace(Section section, double tallestImage)
    {
        if (tallestImage <= 0) return;
        var needed = section.HeaderDistance + tallestImage + LayoutDefaults.HeaderBodyGapPoints;
        if (section.TopMargin < needed) section.TopMargin = needed;
    }

    static void RenderHeaderFooterParagraph(WordRenderContext ctx, W.Paragraph source,
        HeaderFooter container, InlineOptions options = default)
    {
        var format = WordHelpers.GetParagraphFormatting(source.ParagraphProperties);
        var target = ctx.Prepare(new Paragraph());
        container.Blocks.Add(target);
        target.Alignment = format.Alignment;
        if (format.SpacingBefore > 0) target.SpaceBefore = format.SpacingBefore;
        if (format.SpacingAfter > 0) target.SpaceAfter = format.SpacingAfter;

        WordContentRenderer.Render(ctx, source, source, target, options);
    }

    static IEnumerable<OpenXmlElement> Flatten(OpenXmlElement container)
    {
        foreach (var element in container.Elements())
        {
            if (element is W.SdtBlock { SdtContentBlock: not null } sdt)
            {
                foreach (var inner in Flatten(sdt.SdtContentBlock))
                    yield return inner;
            }
            else
            {
                yield return element;
            }
        }
    }

    static void RenderBodyParagraph(WordRenderContext ctx, Section section, W.Paragraph source)
    {
        var format = ResolveParagraphFormat(ctx, source);

        foreach (var image in ctx.ImagesIn(source))
        {
            if (image.IsPositionedAnchor) AddFloatingImage(ctx, section, image);
            else if (image.IsBackground) CollectBackground(ctx, section, image);
        }

        foreach (var line in WordLineExtractor.Extract(source))
            AddFloatingLine(section, line);

        var target = ctx.Prepare(section.AddParagraph());
        ApplyParagraphFormat(target, format);

        var label = ListLabel(ctx, source);
        if (label != null) target.Add(new TextRun(label, target.DefaultStyle));

        WordContentRenderer.Render(ctx, source, source, target,
            new InlineOptions(SkipImage: image => image.IsPositionedAnchor || image.IsBackground));

        var pictures = source.Elements<W.Run>()
            .Select(run => run.GetFirstChild<W.Picture>())
            .Where(picture => picture != null)
            .ToList();
        if (pictures.Count > 0) VmlTextBoxRenderer.Render(ctx, pictures!, section);

        // An empty paragraph is as tall as its mark, which carries the size of the surrounding text.
        if (target.IsEmpty) target.DefaultStyle = target.DefaultStyle with { Size = format.FontSize };
    }

    static RelativeVertical VerticalAnchor(string? relativeFrom) => relativeFrom?.ToLowerInvariant() switch
    {
        "page" => RelativeVertical.Page,
        "margin" => RelativeVertical.Margin,
        _ => RelativeVertical.Paragraph,
    };

    static RelativeHorizontal HorizontalAnchor(string? relativeFrom) =>
        relativeFrom?.ToLowerInvariant() == "page" ? RelativeHorizontal.Page : RelativeHorizontal.Margin;

    static void AddFloatingLine(Section section, WordLine line)
    {
        section.Blocks.Add(new FloatingBlock
        {
            LineColor = ColorUtils.TryParse(line.Color, out var color) ? color : SKColors.Black,
            LineWidth = LayoutDefaults.BorderWidthPoints,
            Left = line.LeftPoints,
            Top = line.TopPoints,
            Width = line.LengthPoints,
            RelativeHorizontal = HorizontalAnchor(line.HorizontalRelativeFrom),
            RelativeVertical = VerticalAnchor(line.VerticalRelativeFrom),
        });
    }

    static void AddFloatingImage(WordRenderContext ctx, Section section, WordImage info)
    {
        var image = ctx.MaterializeImage(info);
        if (image == null) return;

        var width = info.WidthPoints;
        var height = WordContentRenderer.ScaledHeight(info, width) ?? image.HeightFor(width);

        section.Blocks.Add(new FloatingBlock
        {
            Image = image,
            Url = info.HyperlinkUrl,
            Left = Units.EmuToPoints(info.OffsetXEmu ?? 0),
            Top = Units.EmuToPoints(info.OffsetYEmu ?? 0),
            Width = width,
            Height = height,
            RelativeHorizontal = HorizontalAnchor(info.HorizontalRelativeFrom),
            RelativeVertical = VerticalAnchor(info.VerticalRelativeFrom),
            ReservesSpace = info.WrapKind == "wrapTopAndBottom",
            WrapSide = WrapSideOf(info),
            DistanceTop = info.DistanceTop,
            DistanceBottom = info.DistanceBottom,
            DistanceLeft = info.DistanceLeft,
            DistanceRight = info.DistanceRight,
        });
    }

    static WrapSide WrapSideOf(WordImage info)
    {
        if (info.WrapKind is not ("wrapSquare" or "wrapTight" or "wrapThrough")) return WrapSide.None;
        return info.WrapText?.ToLowerInvariant() switch
        {
            "left" => WrapSide.Left,
            "right" => WrapSide.Right,
            "largest" => WrapSide.Largest,
            _ => WrapSide.Both,
        };
    }

    internal static string? ListLabel(WordRenderContext ctx, W.Paragraph source)
    {
        var numbering = source.ParagraphProperties?.NumberingProperties;
        var numId = numbering?.NumberingId?.Val?.Value;
        if (numId == null) return null;

        var level = (int?)numbering!.NumberingLevelReference?.Val?.Value ?? 0;
        return ctx.Numbering.NextLabel(numId.Value.ToString(), level);
    }

    readonly record struct EffectiveParagraphFormat(
        ParagraphAlignment Alignment,
        double LeftIndent,
        double RightIndent,
        double FirstLineIndent,
        double? SpaceBefore,
        double? SpaceAfter,
        double? LineSpacing,
        string? LineRule,
        bool PageBreakBefore,
        string? ShadingColor,
        double FontSize,
        W.Tabs? Tabs,
        BorderInfo? Borders,
        bool KeepNext,
        bool KeepLines,
        bool WidowControl,
        bool RightToLeft);

    static EffectiveParagraphFormat ResolveParagraphFormat(WordRenderContext ctx, W.Paragraph source)
    {
        var pPr = source.ParagraphProperties;
        var styleId = pPr?.ParagraphStyleId?.Val?.Value;

        var direct = WordHelpers.GetParagraphFormatting(pPr);
        var style = WordHelpers.GetParagraphFormatting(
            WordHelpers.GetStyleParagraphProperties(ctx.MainPart, styleId ?? string.Empty));
        var normal = WordHelpers.GetParagraphFormatting(
            WordHelpers.GetStyleParagraphProperties(ctx.MainPart, "Normal"));
        var docDefaults = WordHelpers.GetParagraphFormatting(ctx.Styles.DocDefaultsParagraphProperties);

        var rightToLeft = direct.Bidi || style.Bidi;
        var alignment = pPr?.Justification != null ? direct.Alignment : style.Alignment;
        if (rightToLeft) alignment = pPr?.Justification == null && style.Alignment == ParagraphAlignment.Left
            ? ParagraphAlignment.Right
            : MirrorForRightToLeft(alignment);

        var leftIndent = style.LeftIndent;
        var rightIndent = style.RightIndent;
        var firstLine = style.FirstLineIndent;
        if (pPr?.Indentation != null)
        {
            leftIndent = direct.LeftIndent;
            firstLine = direct.FirstLineIndent;
            if (pPr.Indentation.Right != null) rightIndent = direct.RightIndent;
        }

        double? spaceBefore =
            direct.HasExplicitSpacingBefore ? direct.SpacingBefore :
            style.HasExplicitSpacingBefore ? style.SpacingBefore :
            normal.HasExplicitSpacingBefore ? normal.SpacingBefore :
            docDefaults.HasExplicitSpacingBefore ? docDefaults.SpacingBefore :
            null;

        double? spaceAfter =
            direct.HasExplicitSpacingAfter ? direct.SpacingAfter :
            style.HasExplicitSpacingAfter ? style.SpacingAfter :
            normal.HasExplicitSpacingAfter ? normal.SpacingAfter :
            docDefaults.HasExplicitSpacingAfter ? docDefaults.SpacingAfter :
            null;

        var lineSpacing = direct.LineSpacing ?? style.LineSpacing ?? normal.LineSpacing ?? docDefaults.LineSpacing;
        var lineRule = direct.LineRule ?? style.LineRule ?? normal.LineRule ?? docDefaults.LineRule;

        return new EffectiveParagraphFormat(alignment, leftIndent, rightIndent, firstLine,
            spaceBefore, spaceAfter, lineSpacing, lineRule,
            direct.PageBreakBefore || style.PageBreakBefore,
            direct.ShadingColor ?? style.ShadingColor,
            ResolveFontSize(ctx, source, styleId),
            pPr?.GetFirstChild<W.Tabs>(),
            direct.Borders ?? style.Borders,
            direct.KeepNext ?? style.KeepNext ?? false,
            direct.KeepLines ?? style.KeepLines ?? false,
            direct.WidowControl ?? style.WidowControl ?? normal.WidowControl ?? docDefaults.WidowControl ?? true,
            rightToLeft);
    }

    // In a bidi paragraph Word treats left and right as logical start and end.
    internal static ParagraphAlignment MirrorForRightToLeft(ParagraphAlignment alignment) => alignment switch
    {
        ParagraphAlignment.Left => ParagraphAlignment.Right,
        ParagraphAlignment.Right => ParagraphAlignment.Left,
        _ => alignment,
    };

    static double ResolveFontSize(WordRenderContext ctx, W.Paragraph source, string? styleId)
    {
        double largest = 0;
        foreach (var run in source.Elements<W.Run>())
        {
            var size = WordHelpers.ResolveRunFormatting(ctx.MainPart, run, source).Size;
            if (size.HasValue && size.Value > largest) largest = size.Value;
        }
        if (largest > 0) return largest;

        var candidates = new[]
        {
            source.ParagraphProperties?.ParagraphMarkRunProperties?
                .GetFirstChild<W.FontSize>()?.Val?.Value,
            WordHelpers.GetStyleRunProperties(ctx.MainPart, styleId ?? string.Empty)?.FontSize?.Val?.Value,
            WordHelpers.GetStyleRunProperties(ctx.MainPart, "Normal")?.FontSize?.Val?.Value,
            ctx.Styles.DocDefaultsRunProperties?.FontSize?.Val?.Value,
        };
        foreach (var candidate in candidates)
            if (Units.HalfPointsToPoints(candidate) is { } size) return size;

        return LayoutDefaults.FontSizePoints;
    }

    static void ApplyParagraphFormat(Paragraph target, EffectiveParagraphFormat format)
    {
        target.Alignment = format.Alignment;
        if (format.LeftIndent > 0) target.LeftIndent = format.LeftIndent;
        if (format.RightIndent > 0) target.RightIndent = format.RightIndent;
        if (format.FirstLineIndent != 0) target.FirstLineIndent = format.FirstLineIndent;
        if (format.SpaceBefore.HasValue) target.SpaceBefore = format.SpaceBefore.Value;
        if (format.SpaceAfter.HasValue) target.SpaceAfter = format.SpaceAfter.Value;
        target.PageBreakBefore = format.PageBreakBefore;
        target.KeepWithNext = format.KeepNext;
        target.KeepTogether = format.KeepLines;
        target.WidowControl = format.WidowControl;
        target.RightToLeft = format.RightToLeft;
        if (ColorUtils.TryParse(format.ShadingColor, out var shading)) target.Shading = shading;
        if (format.Borders is { } borders) target.Borders = ToBorders(borders);

        if (format.LineSpacing is { } spacing && spacing > 0)
        {
            target.LineSpacing = spacing;
            target.LineSpacingRule = format.LineRule?.ToLowerInvariant() switch
            {
                "exact" => LineSpacingRule.Exactly,
                "atleast" => LineSpacingRule.AtLeast,
                _ => LineSpacingRule.Multiple,
            };
        }

        if (format.Tabs == null) return;
        target.TabStops = [];
        foreach (var tab in format.Tabs.Elements<W.TabStop>())
        {
            if (tab.Position?.HasValue != true) continue;
            var alignment =
                tab.Val?.Value == W.TabStopValues.Center ? TabAlignment.Center :
                tab.Val?.Value == W.TabStopValues.Right ? TabAlignment.Right :
                TabAlignment.Left;
            target.TabStops.Add(new TabStop(Units.TwipsToPoints(tab.Position.Value), alignment));
        }
    }

    static Borders ToBorders(BorderInfo info) => new()
    {
        Top = Edge(info.TopWidth, info.TopColor, info.TopStyle),
        Bottom = Edge(info.BottomWidth, info.BottomColor, info.BottomStyle),
        Left = Edge(info.LeftWidth, info.LeftColor, info.LeftStyle),
        Right = Edge(info.RightWidth, info.RightColor, info.RightStyle),
    };

    static BorderEdge? Edge(double width, string? colorValue, string? style)
    {
        if (width <= 0) return null;
        return new BorderEdge(width, ColorUtils.TryParse(colorValue, out var color) ? color : SKColors.Black,
            BorderPatterns.From(style));
    }
}
