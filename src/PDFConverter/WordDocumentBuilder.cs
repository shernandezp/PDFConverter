using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal readonly record struct PageArt(string Path, double Width, double Height);

internal static class WordDocumentBuilder
{
    public static PdfDocumentRenderer Build(WordprocessingDocument word, TempImageStore images)
    {
        OpenXmlHelpers.EnsureFontResolverInitialized();

        var ctx = new WordRenderContext(word, images);
        var body = ctx.MainPart.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        var document = new Document();
        var section = document.AddSection();
        var sectionProperties = body.Elements<W.SectionProperties>().LastOrDefault();

        ApplyPageSetup(section, sectionProperties);

        var backgrounds = new List<PageArt>();
        RenderHeaders(ctx, section, sectionProperties, backgrounds);
        RenderFooters(ctx, section, sectionProperties);
        RenderBody(ctx, section, body, backgrounds);
        ApplyDefaultFont(ctx, document);

        var renderer = new PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();

        DrawBackgrounds(renderer, backgrounds);
        PdfImageLinks.Apply(renderer.PdfDocument, ctx.HyperlinkedImages);
        return renderer;
    }

    static void ApplyPageSetup(Section section, W.SectionProperties? sectionProperties)
    {
        var pageSetup = section.PageSetup;
        pageSetup.LeftMargin = Unit.FromCentimeter(LayoutDefaults.PageMarginCentimeters);
        pageSetup.RightMargin = Unit.FromCentimeter(LayoutDefaults.PageMarginCentimeters);

        var pageSize = sectionProperties?.GetFirstChild<W.PageSize>();
        if (pageSize?.Width is { HasValue: true } width)
            pageSetup.PageWidth = Unit.FromPoint(Units.TwipsToPoints(width.Value));
        if (pageSize?.Height is { HasValue: true } height)
            pageSetup.PageHeight = Unit.FromPoint(Units.TwipsToPoints(height.Value));

        var margin = sectionProperties?.GetFirstChild<W.PageMargin>();
        if (margin == null) return;

        if (margin.Top is { HasValue: true } top)
            pageSetup.TopMargin = Unit.FromPoint(Units.TwipsToPoints(top.Value));
        if (margin.Bottom is { HasValue: true } bottom)
            pageSetup.BottomMargin = Unit.FromPoint(Units.TwipsToPoints(bottom.Value));
        if (margin.Left is { HasValue: true } left)
            pageSetup.LeftMargin = Unit.FromPoint(Units.TwipsToPoints(left.Value));
        if (margin.Right is { HasValue: true } right)
            pageSetup.RightMargin = Unit.FromPoint(Units.TwipsToPoints(right.Value));
        if (margin.Header is { HasValue: true } header)
            pageSetup.HeaderDistance = Unit.FromPoint(Units.TwipsToPoints(header.Value));
        if (margin.Footer is { HasValue: true } footer)
            pageSetup.FooterDistance = Unit.FromPoint(Units.TwipsToPoints(footer.Value));
    }

    static double ContentWidth(Section section) =>
        section.PageSetup.PageWidth.Point - section.PageSetup.LeftMargin.Point - section.PageSetup.RightMargin.Point;

    static void RenderHeaders(WordRenderContext ctx, Section section,
        W.SectionProperties? sectionProperties, List<PageArt> backgrounds)
    {
        foreach (var reference in sectionProperties?.Elements<W.HeaderReference>() ?? [])
        {
            var id = reference.Id?.Value;
            if (string.IsNullOrEmpty(id)) continue;
            if (ctx.MainPart.GetPartById(id) is not HeaderPart { Header: not null } part) continue;

            var container = Container(section.Headers, section, reference.Type?.InnerText);
            RenderHeaderPart(ctx, section, part.Header, container, backgrounds);
        }
    }

    static void RenderFooters(WordRenderContext ctx, Section section, W.SectionProperties? sectionProperties)
    {
        foreach (var reference in sectionProperties?.Elements<W.FooterReference>() ?? [])
        {
            var id = reference.Id?.Value;
            if (string.IsNullOrEmpty(id)) continue;
            if (ctx.MainPart.GetPartById(id) is not FooterPart { Footer: not null } part) continue;

            var container = Container(section.Footers, section, reference.Type?.InnerText);
            foreach (var paragraph in part.Footer.Elements<W.Paragraph>())
                RenderHeaderFooterParagraph(ctx, paragraph, container);
        }
    }

    static HeaderFooter Container(HeadersFooters collection, Section section, string? referenceType)
    {
        if (string.Equals(referenceType, "First", StringComparison.OrdinalIgnoreCase))
        {
            section.PageSetup.DifferentFirstPageHeaderFooter = true;
            return collection.FirstPage;
        }
        if (string.Equals(referenceType, "Even", StringComparison.OrdinalIgnoreCase))
        {
            section.PageSetup.OddAndEvenPagesHeaderFooter = true;
            return collection.EvenPage;
        }
        return collection.Primary;
    }

    static void RenderHeaderPart(WordRenderContext ctx, Section section, W.Header header,
        HeaderFooter container, List<PageArt> backgrounds)
    {
        var pageWidth = section.PageSetup.PageWidth.Point;
        var pageHeight = section.PageSetup.PageHeight.Point;
        double tallestImage = 0;

        bool IsArt(WordImage image) => IsPageArt(image, pageWidth, pageHeight);

        foreach (var paragraph in header.Elements<W.Paragraph>())
        {
            foreach (var image in ctx.ImagesIn(paragraph))
            {
                if (IsArt(image)) CollectBackground(ctx, section, image, backgrounds);
                else tallestImage = Math.Max(tallestImage, image.HeightPoints);
            }

            RenderHeaderFooterParagraph(ctx, paragraph, container,
                new InlineOptions(MaxImageWidthPoints: pageWidth, SkipImage: IsArt));
        }

        ReserveHeaderSpace(section, tallestImage);
    }

    // A header picture sitting behind the text, or covering most of the page, is page art rather
    // than header content: Word paints it under everything, MigraDoc has no equivalent.
    static bool IsPageArt(WordImage image, double pageWidth, double pageHeight)
    {
        if (image.IsBackground) return true;
        if (!image.IsAnchor || !image.ExtentCxEmu.HasValue || !image.ExtentCyEmu.HasValue) return false;
        return image.WidthPoints > pageWidth * LayoutDefaults.FullPageImageRatio
            && image.HeightPoints > pageHeight * LayoutDefaults.FullPageImageRatio;
    }

    static void CollectBackground(WordRenderContext ctx, Section section, WordImage image, List<PageArt> backgrounds)
    {
        var path = ctx.MaterializeImage(image);
        if (path == null) return;
        if (backgrounds.All(art => art.Path != path))
            backgrounds.Add(new PageArt(path, image.WidthPoints, image.HeightPoints));

        var pageHeight = section.PageSetup.PageHeight.Point;
        if (!image.ExtentCyEmu.HasValue || pageHeight <= 0) return;

        // A band across the top is header art, not a full-page backdrop: keep body text clear of it.
        var imageHeight = image.HeightPoints;
        if (imageHeight / pageHeight >= LayoutDefaults.HeaderBackgroundPageRatio) return;

        var available = Math.Max(0, pageHeight - section.PageSetup.TopMargin.Point
            - section.PageSetup.BottomMargin.Point - 10.0);
        var headerHeight = Math.Min(imageHeight, available);
        if (section.PageSetup.HeaderDistance.Point < headerHeight)
            section.PageSetup.HeaderDistance = Unit.FromPoint(headerHeight);
    }

    static void ReserveHeaderSpace(Section section, double tallestImage)
    {
        if (tallestImage <= 0) return;
        var needed = section.PageSetup.HeaderDistance.Point + tallestImage + LayoutDefaults.HeaderBodyGapPoints;
        if (section.PageSetup.TopMargin.Point < needed)
            section.PageSetup.TopMargin = Unit.FromPoint(needed);
    }

    static void RenderHeaderFooterParagraph(WordRenderContext ctx, W.Paragraph source,
        HeaderFooter container, InlineOptions options = default)
    {
        var format = WordHelpers.GetParagraphFormatting(source.ParagraphProperties);
        var target = container.AddParagraph();
        target.Format.Alignment = format.Alignment;
        if (format.SpacingBefore > 0) target.Format.SpaceBefore = Unit.FromPoint(format.SpacingBefore);
        if (format.SpacingAfter > 0) target.Format.SpaceAfter = Unit.FromPoint(format.SpacingAfter);

        WordContentRenderer.Render(ctx, source, source, target, options);
    }

    static void RenderBody(WordRenderContext ctx, Section section, W.Body body, List<PageArt> backgrounds)
    {
        var atTopOfPage = true;
        foreach (var element in Flatten(body))
        {
            if (element is W.Paragraph paragraph)
            {
                RenderBodyParagraph(ctx, section, paragraph, backgrounds, atTopOfPage);
                atTopOfPage = false;
            }
            else if (element is W.Table table)
            {
                WordTableRenderer.RenderTable(ctx, section, table);
                atTopOfPage = false;
            }
        }
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

    static void RenderBodyParagraph(WordRenderContext ctx, Section section, W.Paragraph source,
        List<PageArt> backgrounds, bool atTopOfPage)
    {
        var format = ResolveParagraphFormat(ctx, source);

        foreach (var image in ctx.ImagesIn(source))
        {
            if (image.IsPositionedAnchor) AddFloatingImage(ctx, section, image, atTopOfPage);
            else if (image.IsBackground) CollectBackground(ctx, section, image, backgrounds);
        }

        foreach (var line in WordLineExtractor.Extract(source))
            AddFloatingLine(section, line, atTopOfPage);

        // MigraDoc drops SpaceBefore on the first paragraph of a section, where Word applies it, so
        // it is added to the top margin instead.
        if (atTopOfPage && format.SpaceBefore is > 0)
            section.PageSetup.TopMargin += Unit.FromPoint(format.SpaceBefore.Value);

        var target = section.AddParagraph();
        ApplyParagraphFormat(target, format, section);

        var label = ListLabel(ctx, source);
        if (label != null) target.AddText(label);

        var added = WordContentRenderer.Render(ctx, source, source, target,
            new InlineOptions(SkipImage: image => image.IsPositionedAnchor || image.IsBackground));

        var pictures = source.Elements<W.Run>()
            .Select(run => run.GetFirstChild<W.Picture>())
            .Where(picture => picture != null)
            .ToList();
        if (pictures.Count > 0) VmlTextBoxRenderer.Render(ctx, pictures!, section);

        if (!added && label == null)
            target.AddFormattedText(" ").Size = format.FontSize;
    }

    // MigraDoc ignores a Top offset for paragraph-relative shapes but honours it for margin-relative
    // ones, and on the first block the paragraph top is the margin top.
    static RelativeVertical VerticalAnchor(WordImage info, bool atTopOfPage) =>
        info.VerticalRelativeFrom?.ToLowerInvariant() switch
        {
            "page" => RelativeVertical.Page,
            "margin" => RelativeVertical.Margin,
            _ => atTopOfPage ? RelativeVertical.Margin : RelativeVertical.Paragraph,
        };

    // MigraDoc has no line shape, so a rule is drawn as the bottom border of an empty frame. The
    // frame's height carries the vertical offset, which also sidesteps the ignored Top on a
    // paragraph-relative anchor.
    static void AddFloatingLine(Section section, WordLine line, bool atTopOfPage)
    {
        try
        {
            var frame = section.AddTextFrame();
            frame.Width = Unit.FromPoint(line.LengthPoints);
            frame.Height = Unit.FromPoint(Math.Max(line.TopPoints, 0) + 4);
            frame.WrapFormat.Style = WrapStyle.Through;
            frame.RelativeHorizontal = line.HorizontalRelativeFrom?.ToLowerInvariant() == "page"
                ? RelativeHorizontal.Page
                : RelativeHorizontal.Margin;
            frame.RelativeVertical = line.VerticalRelativeFrom?.ToLowerInvariant() switch
            {
                "page" => RelativeVertical.Page,
                "margin" => RelativeVertical.Margin,
                _ => atTopOfPage ? RelativeVertical.Margin : RelativeVertical.Paragraph,
            };
            frame.Left = Unit.FromPoint(line.LeftPoints);
            if (frame.RelativeVertical != RelativeVertical.Paragraph)
                frame.Top = Unit.FromPoint(0);

            var rule = frame.AddParagraph();
            rule.Format.Font.Size = 1;
            rule.Format.SpaceBefore = Unit.FromPoint(Math.Max(line.TopPoints, 0));
            rule.Format.SpaceAfter = 0;
            rule.Format.Borders.Top.Width = Unit.FromPoint(LayoutDefaults.BorderWidthPoints);
            rule.Format.Borders.Top.Color =
                ColorUtils.TryParse(line.Color, out var color) ? color : Colors.Black;
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Failed adding connector line: {ex.Message}");
        }
    }

    static void AddFloatingImage(WordRenderContext ctx, Section section, WordImage info, bool atTopOfPage)
    {
        var path = ctx.MaterializeImage(info);
        if (path == null) return;

        try
        {
            var image = section.AddImage(path);
            WordContentRenderer.SetSize(image, info, info.WidthPoints, info.HeightPoints);
            image.RelativeHorizontal = info.HorizontalRelativeFrom?.ToLowerInvariant() switch
            {
                "page" => RelativeHorizontal.Page,
                _ => RelativeHorizontal.Margin,
            };
            image.RelativeVertical = VerticalAnchor(info, atTopOfPage);

            // MigraDoc cannot flow text beside a shape, and reserving a full-width band for a square
            // wrap displaces text far more than floating the shape loses in vertical offset.
            image.WrapFormat.Style = info.WrapKind == "wrapTopAndBottom"
                ? WrapStyle.TopBottom
                : WrapStyle.Through;
            image.Left = Unit.FromPoint(Units.EmuToPoints(info.OffsetXEmu ?? 0));
            image.Top = Unit.FromPoint(Units.EmuToPoints(info.OffsetYEmu ?? 0));

            ctx.TrackHyperlinkedImage(info.HyperlinkUrl, info.WidthPoints, info.HeightPoints);
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Failed adding floating image '{path}': {ex.Message}");
        }
    }

    static string? ListLabel(WordRenderContext ctx, W.Paragraph source)
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
        W.Tabs? Tabs);

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

        var alignment = pPr?.Justification != null ? direct.Alignment : style.Alignment;

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
            pPr?.GetFirstChild<W.Tabs>());
    }

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

    static void ApplyParagraphFormat(Paragraph target, EffectiveParagraphFormat format, Section section)
    {
        var f = target.Format;
        f.Alignment = format.Alignment;
        if (format.LeftIndent > 0) f.LeftIndent = Unit.FromPoint(format.LeftIndent);
        if (format.RightIndent > 0) f.RightIndent = Unit.FromPoint(format.RightIndent);
        if (format.FirstLineIndent != 0) f.FirstLineIndent = Unit.FromPoint(format.FirstLineIndent);
        if (format.SpaceBefore.HasValue) f.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Value);
        if (format.SpaceAfter.HasValue) f.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Value);
        f.PageBreakBefore = format.PageBreakBefore;
        if (ColorUtils.TryParse(format.ShadingColor, out var shading)) f.Shading.Color = shading;

        ApplyLineSpacing(f, format);
        ApplyTabStops(f, format.Tabs, ContentWidth(section));
    }

    static void ApplyLineSpacing(MigraDoc.DocumentObjectModel.ParagraphFormat target, EffectiveParagraphFormat format)
    {
        if (format.LineSpacing is not { } spacing || spacing <= 0) return;

        switch (format.LineRule?.ToLowerInvariant())
        {
            case "exact":
                target.LineSpacing = Unit.FromPoint(spacing);
                target.LineSpacingRule = LineSpacingRule.Exactly;
                break;
            case "atleast":
                target.LineSpacing = Unit.FromPoint(spacing);
                target.LineSpacingRule = LineSpacingRule.AtLeast;
                break;
            default:
                target.LineSpacing = spacing;
                target.LineSpacingRule = LineSpacingRule.Multiple;
                break;
        }
    }

    static void ApplyTabStops(MigraDoc.DocumentObjectModel.ParagraphFormat target, W.Tabs? tabs, double contentWidth)
    {
        if (tabs != null)
        {
            foreach (var tab in tabs.Elements<W.TabStop>())
            {
                if (tab.Position?.HasValue != true) continue;
                var alignment =
                    tab.Val?.Value == W.TabStopValues.Center ? TabAlignment.Center :
                    tab.Val?.Value == W.TabStopValues.Right ? TabAlignment.Right :
                    TabAlignment.Left;
                target.TabStops.AddTabStop(Unit.FromPoint(Units.TwipsToPoints(tab.Position.Value)), alignment);
            }
            return;
        }

        for (var position = LayoutDefaults.TabStopIntervalPoints; position < contentWidth;
             position += LayoutDefaults.TabStopIntervalPoints)
            target.TabStops.AddTabStop(Unit.FromPoint(position));
    }

    static void ApplyDefaultFont(WordRenderContext ctx, Document document)
    {
        var normal = document.Styles["Normal"];
        if (normal == null) return;

        var defaults = ctx.Styles.DocDefaultsRunProperties;
        normal.Font.Name = WordHelpers.GetRunFontFamily(defaults)
            ?? ctx.Styles.ThemeFont
            ?? ctx.UsedFonts.FirstOrDefault()
            ?? "Arial";

        if (Units.HalfPointsToPoints(defaults?.FontSize?.Val?.Value) is { } size)
            normal.Font.Size = size;
    }

    static void DrawBackgrounds(PdfDocumentRenderer renderer, List<PageArt> backgrounds)
    {
        if (backgrounds.Count == 0) return;

        foreach (var page in renderer.PdfDocument.Pages)
        {
            using var graphics = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend);
            foreach (var art in backgrounds)
            {
                try
                {
                    using var image = XImage.FromFile(art.Path);
                    var width = art.Width > 0 ? art.Width : page.Width.Point;
                    var height = art.Height > 0 ? art.Height : page.Height.Point;
                    graphics.DrawImage(image, 0, 0, width, height);
                }
                catch (Exception ex)
                {
                    OpenXmlHelpers.ImageLoadLogger?.Invoke($"Failed drawing background '{art.Path}': {ex.Message}");
                }
            }
        }
    }
}
