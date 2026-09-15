using DocumentFormat.OpenXml;
using MigraDoc.DocumentObjectModel;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal readonly record struct InlineOptions(
    double MaxImageWidthPoints = 0,
    Func<WordImage, bool>? SkipImage = null,
    bool ConditionalBold = false,
    string? ConditionalColor = null,
    double MaxTextWidthPoints = 0);

internal static class WordContentRenderer
{
    const string EmojiFontName = "Noto Emoji";

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
                    target.AddBookmark(bookmark.Name!.Value!, false);
                    break;
            }
        }
        return added;
    }

    static bool RenderRun(WordRenderContext ctx, W.Run run, W.Paragraph source, Paragraph target,
        InlineOptions options, FieldState fields)
    {
        var format = ResolveFormat(ctx, run, source, options);
        ctx.TrackFont(format.FontFamily);

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
                case W.Text text when !fields.SuppressResult && !string.IsNullOrEmpty(text.Text):
                    AddText(target, Preserve(text, format.TransformText(text.Text), target),
                        format, options.MaxTextWidthPoints);
                    added = true;
                    break;
                case W.NoBreakHyphen when !fields.SuppressResult:
                    AddText(target, "‑", format);
                    added = true;
                    break;
                case W.SymbolChar symbol when !fields.SuppressResult:
                    added |= AddSymbol(target, symbol, format);
                    break;
                case W.Break brk:
                    AddBreak(target, brk);
                    break;
                case W.TabChar:
                    target.AddTab();
                    break;
                case W.Drawing:
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

    const char NonBreakingSpace = (char)0x00A0;

    // MigraDoc drops blanks at the start of a line, while Word keeps them when the run asks for
    // whitespace to be preserved — which is how these templates indent a heading beside a logo.
    static string Preserve(W.Text source, string text, Paragraph target)
    {
        if (source.Space?.Value != SpaceProcessingModeValues.Preserve) return text;
        if (!text.Contains(' ')) return text;

        var result = new System.Text.StringBuilder(text.Length);
        var atLineStart = target.Elements.Count == 0;
        var index = 0;

        while (index < text.Length)
        {
            if (text[index] != ' ')
            {
                result.Append(text[index++]);
                atLineStart = false;
                continue;
            }

            var blanks = 0;
            while (index < text.Length && text[index] == ' ') { blanks++; index++; }

            if (!atLineStart)
            {
                result.Append(' ');
                blanks--;
            }
            result.Append(NonBreakingSpace, blanks);
            atLineStart = false;
        }
        return result.ToString();
    }

    static void AddText(Paragraph target, string text, RunFormat format, double maxWidthPoints = 0)
    {
        var lines = TextMeasure.SplitOverlongWords(text, format, maxWidthPoints);
        for (var i = 0; i < lines.Count; i++)
        {
            if (i > 0) target.AddLineBreak();
            AddSegments(target, lines[i], format);
        }
    }

    static void AddSegments(Paragraph target, string text, RunFormat format)
    {
        if (!ConverterExtensions.ContainsEmoji(text))
        {
            format.ApplyTo(target.AddFormattedText(text));
            return;
        }

        foreach (var (segment, isEmoji) in ConverterExtensions.SplitEmojiSegments(text))
        {
            var formatted = target.AddFormattedText(segment);
            format.ApplyTo(formatted);
            if (isEmoji) formatted.Font.Name = EmojiFontName;
        }
    }

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

    // MigraDoc can only break before a paragraph, which covers Word's usual encoding of a manual
    // page break: an otherwise empty paragraph holding a single w:br.
    static void AddBreak(Paragraph target, W.Break brk)
    {
        if (brk.Type?.Value == W.BreakValues.Page && target.Elements.Count == 0)
            target.Format.PageBreakBefore = true;
        else
            target.AddLineBreak();
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
        var path = ctx.MaterializeImage(info);
        if (path == null) return false;

        try
        {
            var image = target.AddImage(path);
            var width = maxWidthPoints > 0 ? Math.Min(info.WidthPoints, maxWidthPoints) : info.WidthPoints;
            var height = info.HeightPoints * (width / Math.Max(info.WidthPoints, 0.01));
            SetSize(image, info, width, height);

            ctx.TrackHyperlinkedImage(hyperlinkUrl, width, height);
            return true;
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Failed adding image '{path}': {ex.Message}");
            return false;
        }
    }

    // Word stretches a picture to the extent stored in the drawing, which need not match the
    // file's own aspect ratio; letting MigraDoc derive the height distorts such pictures.
    public static void SetSize(MigraDoc.DocumentObjectModel.Shapes.Image image,
        WordImage info, double widthPoints, double heightPoints)
    {
        var hasBothExtents = info.ExtentCxEmu.HasValue && info.ExtentCyEmu.HasValue;
        image.LockAspectRatio = !hasBothExtents;
        image.Width = Unit.FromPoint(widthPoints);
        if (hasBothExtents) image.Height = Unit.FromPoint(heightPoints);
    }

    static bool RenderHyperlink(WordRenderContext ctx, W.Hyperlink hyperlink, W.Paragraph source,
        Paragraph target, InlineOptions options, FieldState fields)
    {
        var (destination, type) = ResolveDestination(ctx, hyperlink);
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
            ctx.TrackFont(format.FontFamily);

            var text = run.InnerText;
            if (string.IsNullOrEmpty(text)) continue;

            var link = target.AddHyperlink(destination, type);
            var formatted = link.AddFormattedText(format.TransformText(text));
            format.ApplyTo(formatted);
            if (format.Color == null) formatted.Color = Colors.Blue;
            formatted.Underline = Underline.Single;
            added = true;
        }
        return added;
    }

    static (string? Destination, HyperlinkType Type) ResolveDestination(WordRenderContext ctx, W.Hyperlink hyperlink)
    {
        var relationshipId = hyperlink.Id?.Value;
        if (!string.IsNullOrEmpty(relationshipId))
        {
            var owner = WordImageExtractor.OwnerOf(hyperlink, ctx.MainPart);
            try
            {
                var uri = owner.HyperlinkRelationships.FirstOrDefault(r => r.Id == relationshipId)?.Uri?.ToString();
                if (!string.IsNullOrEmpty(uri)) return (uri, HyperlinkType.Web);
            }
            catch (Exception ex)
            {
                OpenXmlHelpers.ImageLoadLogger?.Invoke(
                    $"Unresolved hyperlink relationship '{relationshipId}': {ex.Message}");
            }
        }

        var anchor = hyperlink.Anchor?.Value;
        return string.IsNullOrEmpty(anchor) ? (null, HyperlinkType.Web) : (anchor, HyperlinkType.Bookmark);
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

            var container = target.AddFormattedText();
            format?.ApplyTo(container);
            if (keyword == "PAGE") container.AddPageField();
            else container.AddNumPagesField();
            return true;
        }
    }
}
