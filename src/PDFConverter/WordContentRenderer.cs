using DocumentFormat.OpenXml;
using PDFConverter.Layout;
using SkiaSharp;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal readonly record struct InlineOptions(
    double MaxImageWidthPoints = 0,
    Func<WordImage, bool>? SkipImage = null,
    bool ConditionalBold = false,
    string? ConditionalColor = null);

internal static class WordContentRenderer
{
    public static bool Render(WordRenderContext ctx, OpenXmlElement container, W.Paragraph source,
        Paragraph target, InlineOptions options = default) =>
        RenderChildren(ctx, container, source, target, options, new FieldState());

    static bool RenderChildren(WordRenderContext ctx, OpenXmlElement container, W.Paragraph source,
        Paragraph target, InlineOptions options, FieldState fields)
    {
        var added = false;
        foreach (var child in container.ChildElements)
        {
            switch (child)
            {
                case W.Run run:
                    added |= RenderRun(ctx, run, source, target, options, fields);
                    break;
                case W.Hyperlink hyperlink:
                    added |= RenderHyperlink(ctx, hyperlink, source, target, options, fields);
                    break;
                case W.SdtRun sdtRun when sdtRun.SdtContentRun != null:
                    added |= RenderChildren(ctx, sdtRun.SdtContentRun, source, target, options, fields);
                    break;
                case W.SimpleField simpleField:
                    added |= RenderSimpleField(ctx, simpleField, source, target, options, fields);
                    break;
                case W.BookmarkStart bookmark when !string.IsNullOrEmpty(bookmark.Name?.Value):
                    target.Add(new BookmarkInline(bookmark.Name!.Value!));
                    break;
            }
        }
        return added;
    }

    static bool RenderRun(WordRenderContext ctx, W.Run run, W.Paragraph source, Paragraph target,
        InlineOptions options, FieldState fields)
    {
        var format = ResolveFormat(ctx, run, source, options);

        var added = false;
        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case W.FieldChar fieldChar:
                    added |= fields.OnFieldChar(fieldChar, target, format);
                    break;
                case W.FieldCode code:
                    fields.OnInstruction(code.Text);
                    break;
                case W.Text text when !fields.SuppressResult && !format.Hidden && !string.IsNullOrEmpty(text.Text):
                    AddText(target, format.TransformText(text.Text), format);
                    added = true;
                    break;
                case W.NoBreakHyphen when !fields.SuppressResult && !format.Hidden:
                    AddText(target, "‑", format);
                    added = true;
                    break;
                case W.SymbolChar symbol when !fields.SuppressResult && !format.Hidden:
                    added |= AddSymbol(target, symbol, format);
                    break;
                case W.Break brk when !format.Hidden:
                    AddBreak(target, brk);
                    break;
                case W.TabChar when !format.Hidden:
                    target.Add(new TabInline());
                    break;
                case W.Drawing when !format.Hidden:
                    added |= RenderRunImages(ctx, run, target, options);
                    break;
            }
        }
        return added;
    }

    static RunFormat ResolveFormat(WordRenderContext ctx, W.Run run, W.Paragraph source, InlineOptions options)
    {
        var format = WordHelpers.ResolveRunFormatting(ctx.MainPart, run, source);
        if (!options.ConditionalBold && options.ConditionalColor == null) return format;

        return format with
        {
            Bold = format.BoldSpecified ? format.Bold : format.Bold || options.ConditionalBold,
            Color = format.Color ?? options.ConditionalColor,
        };
    }

    static void AddText(Paragraph target, string text, RunFormat format, string? url = null) =>
        target.Add(new TextRun(text, format.ToTextStyle(target.DefaultStyle), url));

    static bool AddSymbol(Paragraph target, W.SymbolChar symbol, RunFormat format)
    {
        if (!int.TryParse(symbol.Char?.Value, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var codePoint))
            return false;

        // Symbol fonts keep their glyphs in the private use area; drop the offset so a text font shows them.
        if (codePoint is >= 0xF000 and <= 0xF0FF) codePoint -= 0xF000;
        AddText(target, char.ConvertFromUtf32(codePoint), format with { FontFamily = null });
        return true;
    }

    static void AddBreak(Paragraph target, W.Break brk)
    {
        if (brk.Type?.Value == W.BreakValues.Page && target.IsEmpty)
            target.PageBreakBefore = true;
        else
            target.Add(new LineBreakInline());
    }

    static bool RenderRunImages(WordRenderContext ctx, OpenXmlElement scope, Paragraph target, InlineOptions options)
    {
        var added = false;
        foreach (var info in ctx.ImagesIn(scope))
        {
            if (options.SkipImage?.Invoke(info) == true) continue;
            added |= AddImage(ctx, info, target, options.MaxImageWidthPoints, info.HyperlinkUrl);
        }
        return added;
    }

    public static bool AddImage(WordRenderContext ctx, WordImage info, Paragraph target,
        double maxWidthPoints, string? hyperlinkUrl)
    {
        var image = ctx.MaterializeImage(info);
        if (image == null) return false;

        var width = maxWidthPoints > 0 ? Math.Min(info.WidthPoints, maxWidthPoints) : info.WidthPoints;
        target.Add(new InlineImage(image, width, ScaledHeight(info, width), hyperlinkUrl));
        return true;
    }

    // Word stretches a picture to the extent stored in the drawing, which need not match the
    // file's own aspect ratio; only without both extents is the file's ratio authoritative.
    public static double? ScaledHeight(WordImage info, double width) =>
        info.ExtentCxEmu.HasValue && info.ExtentCyEmu.HasValue
            ? info.HeightPoints * (width / Math.Max(info.WidthPoints, 0.01))
            : null;

    static bool RenderHyperlink(WordRenderContext ctx, W.Hyperlink hyperlink, W.Paragraph source,
        Paragraph target, InlineOptions options, FieldState fields)
    {
        var destination = ResolveDestination(ctx, hyperlink);
        if (destination == null)
            return RenderChildren(ctx, hyperlink, source, target, options, fields);

        var added = false;
        foreach (var run in hyperlink.Elements<W.Run>())
        {
            if (run.GetFirstChild<W.Drawing>() != null)
            {
                foreach (var info in ctx.ImagesIn(run))
                    added |= AddImage(ctx, info, target, options.MaxImageWidthPoints, destination);
                continue;
            }

            var format = ResolveFormat(ctx, run, source, options);
            var text = run.InnerText;
            if (string.IsNullOrEmpty(text)) continue;

            var style = format.ToTextStyle(target.DefaultStyle) with { Underline = UnderlineStyle.Single };
            if (format.Color == null) style = style with { Color = SKColors.Blue };
            target.Add(new TextRun(format.TransformText(text), style, destination));
            added = true;
        }
        return added;
    }

    static string? ResolveDestination(WordRenderContext ctx, W.Hyperlink hyperlink)
    {
        var relationshipId = hyperlink.Id?.Value;
        if (!string.IsNullOrEmpty(relationshipId))
        {
            var owner = WordImageExtractor.OwnerOf(hyperlink, ctx.MainPart);
            try
            {
                var uri = owner.HyperlinkRelationships.FirstOrDefault(r => r.Id == relationshipId)?.Uri?.ToString();
                if (!string.IsNullOrEmpty(uri)) return uri;
            }
            catch (Exception ex)
            {
                OpenXmlHelpers.ImageLoadLogger?.Invoke(
                    $"Unresolved hyperlink relationship '{relationshipId}': {ex.Message}");
            }
        }

        var anchor = hyperlink.Anchor?.Value;
        return string.IsNullOrEmpty(anchor) ? null : "#" + anchor;
    }

    static bool RenderSimpleField(WordRenderContext ctx, W.SimpleField field, W.Paragraph source,
        Paragraph target, InlineOptions options, FieldState fields)
    {
        var firstRun = field.Elements<W.Run>().FirstOrDefault();
        var format = firstRun == null ? null : WordHelpers.ResolveRunFormatting(ctx.MainPart, firstRun, source);

        return FieldState.TryEmit(field.Instruction?.Value, target, format)
            || RenderChildren(ctx, field, source, target, options, fields);
    }

    // A complex field spans several runs: begin / instrText / separate / cached result / end.
    // Fields we can evaluate are emitted live and their cached result suppressed.
    sealed class FieldState
    {
        readonly Stack<bool> _suppressed = new();
        string? _instruction;
        bool _readingInstruction;

        public bool SuppressResult => _suppressed.Count > 0 && _suppressed.Peek();

        public void OnInstruction(string? text)
        {
            if (_readingInstruction) _instruction += text;
        }

        public bool OnFieldChar(W.FieldChar fieldChar, Paragraph target, RunFormat format)
        {
            var type = fieldChar.FieldCharType?.Value;

            if (type == W.FieldCharValues.Begin)
            {
                _instruction = string.Empty;
                _readingInstruction = true;
                _suppressed.Push(false);
                return false;
            }

            if (type == W.FieldCharValues.Separate)
            {
                _readingInstruction = false;
                var emitted = TryEmit(_instruction, target, format);
                if (_suppressed.Count > 0)
                {
                    _suppressed.Pop();
                    _suppressed.Push(emitted);
                }
                return emitted;
            }

            if (type == W.FieldCharValues.End)
            {
                _readingInstruction = false;
                if (_suppressed.Count > 0) _suppressed.Pop();
            }
            return false;
        }

        public static bool TryEmit(string? instruction, Paragraph target, RunFormat? format)
        {
            if (string.IsNullOrWhiteSpace(instruction)) return false;

            var keyword = instruction.TrimStart()
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .FirstOrDefault()?.ToUpperInvariant();
            if (keyword is not ("PAGE" or "NUMPAGES")) return false;

            var style = format?.ToTextStyle(target.DefaultStyle) ?? target.DefaultStyle;
            target.Add(new FieldRun(keyword == "PAGE" ? FieldKind.Page : FieldKind.NumPages, style));
            return true;
        }
    }
}
