using System.Text;
using HarfBuzzSharp;
using SkiaSharp;
using SkiaSharp.HarfBuzz;
using Buffer = HarfBuzzSharp.Buffer;

namespace PDFConverter.Layout;

internal sealed class ShapedText
{
    public required ushort[] Glyphs;
    public required SKPoint[] Points;
    public required uint[] Clusters;
    public required double Width;

    public static ShapedText Empty { get; } = new() { Glyphs = [], Points = [], Clusters = [], Width = 0 };
}

internal enum TextDirection { Neutral, LeftToRight, RightToLeft }

internal sealed class TextMeasurer(FontRegistry fonts) : IDisposable
{
    readonly Dictionary<(SKTypeface, float, bool, bool, float), SKFont> _fonts = [];
    readonly Dictionary<SKTypeface, SKShaper> _shapers = [];
    readonly Dictionary<(SKTypeface, float, bool, float, float, bool, string), ShapedText> _shaped = [];

    public FontRegistry Fonts { get; } = fonts;

    public FontFace Resolve(TextStyle style, TextStyle defaults) =>
        Fonts.Resolve(style.FontFamily ?? defaults.FontFamily, style.Bold, style.Italic);

    public SKFont FontFor(FontFace face, double size, double scaleX = 1)
    {
        var key = (face.Typeface, (float)size, face.FakeBold, face.FakeItalic, (float)scaleX);
        if (_fonts.TryGetValue(key, out var font)) return font;
        font = face.CreateFont(size);
        font.ScaleX = (float)scaleX;
        _fonts[key] = font;
        return font;
    }

    public ShapedText Shape(FontFace face, double size, string text, double scaleX = 1, double spacing = 0, bool rtl = false)
    {
        if (text.Length == 0) return ShapedText.Empty;
        var key = (face.Typeface, (float)size, face.FakeBold, (float)scaleX, (float)spacing, rtl, text);
        if (_shaped.TryGetValue(key, out var cached)) return cached;

        var font = FontFor(face, size, scaleX);
        if (!_shapers.TryGetValue(face.Typeface, out var shaper))
            _shapers[face.Typeface] = shaper = new SKShaper(face.Typeface);

        using var buffer = new Buffer();
        buffer.AddUtf16(text);
        buffer.GuessSegmentProperties();
        buffer.Direction = rtl ? Direction.RightToLeft : Direction.LeftToRight;
        var result = shaper.Shape(buffer, 0, 0, font);

        var count = result.Codepoints.Length;
        var glyphs = new ushort[count];
        var points = new SKPoint[count];
        for (var i = 0; i < count; i++)
        {
            glyphs[i] = (ushort)result.Codepoints[i];
            points[i] = new SKPoint(result.Points[i].X + (float)(spacing * i), result.Points[i].Y);
        }
        var shaped = new ShapedText
        {
            Glyphs = glyphs, Points = points, Clusters = result.Clusters,
            Width = Math.Max(result.Width + spacing * count, 0),
        };
        if (text.Length <= 64) _shaped[key] = shaped;
        return shaped;
    }

    public double Width(FontFace face, double size, string text, double scaleX = 1, double spacing = 0, bool rtl = false) =>
        Shape(face, size, text, scaleX, spacing, rtl).Width;

    public (double Ascent, double Descent, double Leading) Metrics(FontFace face, double size)
    {
        var metrics = FontFor(face, size).Metrics;
        return (-metrics.Ascent, metrics.Descent, Math.Max(metrics.Leading, 0));
    }

    public (double Position, double Thickness) Underline(FontFace face, double size)
    {
        var metrics = FontFor(face, size).Metrics;
        var position = metrics.UnderlinePosition is { } p && p > 0 ? p : size * 0.1;
        var thickness = metrics.UnderlineThickness is { } t && t > 0 ? t : size * 0.07;
        return (position, thickness);
    }

    public bool HasGlyph(FontFace face, int codepoint) => FontFor(face, 10).GetGlyph(codepoint) != 0;

    public void Dispose()
    {
        foreach (var shaper in _shapers.Values) shaper.Dispose();
        foreach (var font in _fonts.Values) font.Dispose();
    }
}

internal enum ItemKind { Text, Space, Tab, Image, Destination }

internal sealed class LineItem
{
    public ItemKind Kind;
    public double X;
    public double Width;
    public string Text = string.Empty;
    public ShapedText Shaped = ShapedText.Empty;
    public TextDirection Direction;
    public FontFace? Face;
    public double Size;
    public SKColor Color;
    public UnderlineStyle? Underline;
    public StrikeStyle Strike;
    public SKColor? Highlight;
    public double LetterSpacing;
    public double WidthScale = 1;
    public string? Url;
    public double Shift;
    public FieldKind? Field;
    public ImageSource? Image;
    public double Height;
    public string? Destination;
    public double Ascent, Descent, Leading;
}

internal readonly record struct LineSegment(double Left, double Width, double GapBefore = 0);

internal delegate LineSegment LineGeometry(double yOffset, double estimatedHeight);

internal sealed class LaidOutLine
{
    public List<LineItem> Items { get; } = [];
    public double Height;
    public double Baseline;
    public double Width;
    public double Indent;
    public double Offset;
    public double AvailableWidth;
    public bool EndsWithBreak;
}

internal sealed class ParagraphLayout
{
    public List<LaidOutLine> Lines { get; } = [];
    public double Height => Lines.Sum(line => line.Height);
}

internal static class TextLayout
{
    const double SuperscriptScale = 0.65;

    sealed class Segment
    {
        public required string Text;
        public required FontFace Face;
        public required double Size;
        public required double BaseSize;
        public required TextStyle Style;
        public required ShapedText Shaped;
        public TextDirection Direction;
        public SKColor Color;
        public string? Url;
        public double Shift;
        public FieldKind? Field;
        public double Width => Shaped.Width;
    }

    enum AtomKind { Word, Space, Tab, Break, Image, Marker }

    sealed class Atom
    {
        public AtomKind Kind;
        public List<Segment> Segments = [];
        public bool Breakable = true;
        public ImageSource? Image;
        public double Width, Height;
        public string? Url;
        public string? Marker;
    }

    public static ParagraphLayout Layout(Paragraph paragraph, double availableWidth, TextMeasurer measurer,
        Func<FieldKind, string> fieldText, LineGeometry? geometry = null)
    {
        var atoms = BuildAtoms(paragraph, measurer, fieldText);
        var defaultFace = measurer.Resolve(paragraph.DefaultStyle, paragraph.DefaultStyle);
        var defaultMetrics = measurer.Metrics(defaultFace, paragraph.DefaultStyle.Size);
        geometry ??= (_, _) => new LineSegment(0, availableWidth);

        var lines = Break(paragraph, atoms, geometry, defaultMetrics, measurer);

        var result = new ParagraphLayout();
        for (var i = 0; i < lines.Count; i++)
        {
            var line = lines[i];
            var first = i == 0;
            if (paragraph.Alignment == ParagraphAlignment.Justify && i < lines.Count - 1 && !line.EndsWithBreak)
                Justify(line, line.AvailableWidth);
            line.Indent = line.Offset + paragraph.LeftIndent + (first ? paragraph.FirstLineIndent : 0)
                + Alignment(paragraph.Alignment, line.AvailableWidth, line.Width, paragraph.NoWrap);
            result.Lines.Add(line);
        }
        return result;
    }

    static double Alignment(ParagraphAlignment alignment, double lineWidth, double width, bool allowOverflow)
    {
        var offset = alignment switch
        {
            ParagraphAlignment.Center => (lineWidth - width) / 2,
            ParagraphAlignment.Right => lineWidth - width,
            _ => 0,
        };
        return allowOverflow ? offset : Math.Max(0, offset);
    }

    static List<Atom> BuildAtoms(Paragraph paragraph, TextMeasurer measurer, Func<FieldKind, string> fieldText)
    {
        var atoms = new List<Atom>();
        foreach (var inline in paragraph.Inlines)
        {
            switch (inline)
            {
                case TextRun run:
                    AddText(atoms, run.Text, run.Style, run.Url, null, paragraph.DefaultStyle, measurer);
                    break;
                case FieldRun field:
                    AddText(atoms, fieldText(field.Kind), field.Style, null, field.Kind, paragraph.DefaultStyle, measurer);
                    break;
                case TabInline:
                    atoms.Add(new Atom { Kind = AtomKind.Tab });
                    break;
                case LineBreakInline:
                    atoms.Add(new Atom { Kind = AtomKind.Break });
                    break;
                case BookmarkInline bookmark:
                    atoms.Add(new Atom { Kind = AtomKind.Marker, Marker = bookmark.Name });
                    break;
                case InlineImage image:
                    atoms.Add(new Atom
                    {
                        Kind = AtomKind.Image, Image = image.Image, Width = image.Width,
                        Height = image.ResolvedHeight, Url = image.Url,
                    });
                    break;
            }
        }
        return atoms;
    }

    public static TextDirection DirectionOf(string text)
    {
        var ltr = false;
        foreach (var rune in text.EnumerateRunes())
        {
            if (IsRightToLeft(rune.Value)) return TextDirection.RightToLeft;
            if (!ltr && (Rune.IsLetter(rune) || Rune.IsDigit(rune))) ltr = true;
        }
        return ltr ? TextDirection.LeftToRight : TextDirection.Neutral;
    }

    static bool IsRightToLeft(int value) =>
        value is >= 0x0590 and <= 0x08FF
        || value is >= 0xFB1D and <= 0xFDFF
        || value is >= 0xFE70 and <= 0xFEFF
        || value is >= 0x10800 and <= 0x10FFF
        || value is >= 0x1E800 and <= 0x1EFFF;

    static void AddText(List<Atom> atoms, string text, TextStyle style, string? url, FieldKind? field,
        TextStyle defaults, TextMeasurer measurer)
    {
        if (string.IsNullOrEmpty(text)) return;

        var primary = measurer.Resolve(style, defaults);
        var baseSize = style.Size > 0 ? style.Size : defaults.Size;
        var size = style.VerticalAlignment == RunVerticalAlignment.Baseline ? baseSize : baseSize * SuperscriptScale;
        var shift = style.VerticalAlignment switch
        {
            RunVerticalAlignment.Superscript => -baseSize * 0.33,
            RunVerticalAlignment.Subscript => baseSize * 0.12,
            _ => 0,
        };
        var color = style.Color ?? defaults.Color ?? SKColors.Black;

        Atom? current = null;
        FontFace? currentFace = null;
        var buffer = new StringBuilder();

        void Flush()
        {
            if (current == null || currentFace == null || buffer.Length == 0) return;
            var segmentText = buffer.ToString();
            var direction = current.Kind == AtomKind.Space ? TextDirection.Neutral : DirectionOf(segmentText);
            var spacing = current.Kind == AtomKind.Space && style.LetterSpacing < 0 ? 0 : style.LetterSpacing;
            current.Segments.Add(new Segment
            {
                Text = segmentText, Face = currentFace, Size = size, BaseSize = baseSize, Style = style, Color = color,
                Url = url, Shift = shift, Field = field, Direction = direction,
                Shaped = measurer.Shape(currentFace, size, segmentText, style.WidthScale, spacing,
                    direction == TextDirection.RightToLeft),
            });
            buffer.Clear();
        }

        void Start(AtomKind kind, bool breakable)
        {
            Flush();
            current = new Atom { Kind = kind, Breakable = breakable };
            atoms.Add(current);
        }

        foreach (var rune in text.EnumerateRunes())
        {
            var value = rune.Value;
            if (value == '\r') continue;
            if (value == '\n')
            {
                Flush();
                current = null;
                atoms.Add(new Atom { Kind = AtomKind.Break });
                continue;
            }
            if (IsIgnorable(value)) continue;
            if (value == 0x200B)
            {
                Flush();
                current = null;
                atoms.Add(new Atom { Kind = AtomKind.Space, Breakable = true });
                continue;
            }

            var isSpace = Rune.IsWhiteSpace(rune) || value == 0x00A0 || value == 0x202F || value == 0x2007;
            var breakable = isSpace && value is not (0x00A0 or 0x202F or 0x2007);
            var kind = isSpace ? AtomKind.Space : AtomKind.Word;

            var face = isSpace || measurer.HasGlyph(primary, value)
                ? primary
                : measurer.Fonts.Fallback(primary, value) ?? primary;

            if (current == null || current.Kind != kind || (kind == AtomKind.Space && current.Breakable != breakable))
            {
                Start(kind, breakable);
                currentFace = face;
            }
            else if (!ReferenceEquals(face, currentFace))
            {
                Flush();
                currentFace = face;
            }

            buffer.Append(rune.ToString());
        }
        Flush();
    }

    static bool IsIgnorable(int value) =>
        value is 0x200C or 0x00AD or 0x2060 or 0xFEFF
        || value is >= 0x180B and <= 0x180D
        || value is >= 0xE0100 and <= 0xE01EF;

    static List<LaidOutLine> Break(Paragraph paragraph, List<Atom> atoms, LineGeometry geometry,
        (double Ascent, double Descent, double Leading) defaults, TextMeasurer measurer)
    {
        var lines = new List<LaidOutLine>();
        var line = new LaidOutLine();
        var x = 0.0;
        var afterWrap = false;
        var yOffset = 0.0;
        var estimate = defaults.Ascent + defaults.Descent + defaults.Leading;
        var gapBefore = StartLine(line);

        double StartLine(LaidOutLine target)
        {
            var segment = geometry(yOffset, estimate);
            var first = lines.Count == 0;
            target.Offset = segment.Left;
            target.AvailableWidth = Math.Max(segment.Width - paragraph.LeftIndent - paragraph.RightIndent
                - (first ? paragraph.FirstLineIndent : 0), 1);
            return segment.GapBefore;
        }

        double LineWidth() => paragraph.NoWrap ? double.MaxValue : line.AvailableWidth;

        void FinishLine(bool endsWithBreak)
        {
            TrimTrailingSpaces(line);
            line.EndsWithBreak = endsWithBreak;
            Reorder(line, paragraph.RightToLeft);
            Measure(paragraph, line, defaults);
            line.Height += gapBefore;
            line.Baseline += gapBefore;
            yOffset += line.Height;
            estimate = line.Height - gapBefore;
            lines.Add(line);
            line = new LaidOutLine();
            x = 0;
            gapBefore = StartLine(line);
        }

        void Place(Atom atom)
        {
            if (atom.Kind == AtomKind.Image)
            {
                line.Items.Add(new LineItem
                {
                    Kind = ItemKind.Image, X = x, Width = atom.Width, Height = atom.Height,
                    Image = atom.Image, Url = atom.Url,
                });
                x += atom.Width;
                return;
            }

            foreach (var segment in atom.Segments)
            {
                line.Items.Add(ToItem(segment, atom.Kind == AtomKind.Space ? ItemKind.Space : ItemKind.Text, x,
                    segment.Shaped, segment.Text, measurer));
                x += segment.Width;
            }
        }

        for (var index = 0; index < atoms.Count; index++)
        {
            var atom = atoms[index];
            switch (atom.Kind)
            {
                case AtomKind.Break:
                    FinishLine(true);
                    afterWrap = false;
                    continue;

                case AtomKind.Marker:
                    line.Items.Add(new LineItem { Kind = ItemKind.Destination, X = x, Destination = atom.Marker });
                    continue;

                case AtomKind.Tab:
                {
                    var target = TabTarget(paragraph, atoms, index, x, LineWidth());
                    if (target == null)
                    {
                        if (line.Items.Count > 0) FinishLine(false);
                        afterWrap = true;
                        continue;
                    }
                    line.Items.Add(new LineItem { Kind = ItemKind.Tab, X = x, Width = target.Value - x });
                    x = target.Value;
                    continue;
                }

                case AtomKind.Space:
                {
                    if (afterWrap && x == 0 && atom.Breakable) continue;
                    Place(atom);
                    continue;
                }
            }

            var width = atom.Kind == AtomKind.Image ? atom.Width : atom.Segments.Sum(s => s.Width);
            if (x + width <= LineWidth() + 0.01 || line.Items.Count == 0)
            {
                if (width > LineWidth() + 0.01 && atom.Kind == AtomKind.Word)
                {
                    PlaceOverlong(atom, ref line, ref x, LineWidth, measurer, FinishLine);
                    afterWrap = true;
                }
                else
                {
                    Place(atom);
                }
                continue;
            }

            FinishLine(false);
            afterWrap = true;
            if (width > LineWidth() + 0.01 && atom.Kind == AtomKind.Word)
                PlaceOverlong(atom, ref line, ref x, LineWidth, measurer, FinishLine);
            else
                Place(atom);
        }

        FinishLine(false);
        return lines;
    }

    static void PlaceOverlong(Atom atom, ref LaidOutLine line, ref double x,
        Func<double> lineWidth, TextMeasurer measurer, Action<bool> finishLine)
    {
        foreach (var segment in atom.Segments)
        {
            var runes = segment.Text.EnumerateRunes().Select(r => r.ToString()).ToList();
            var start = 0;
            while (start < runes.Count)
            {
                var length = 1;
                var text = runes[start];
                var shaped = Shape(segment, text, measurer);
                while (start + length < runes.Count)
                {
                    var candidate = text + runes[start + length];
                    var candidateShaped = Shape(segment, candidate, measurer);
                    if (x + candidateShaped.Width > lineWidth() + 0.01) break;
                    text = candidate;
                    shaped = candidateShaped;
                    length++;
                }

                if (x > 0 && x + shaped.Width > lineWidth() + 0.01 && line.Items.Count > 0)
                {
                    finishLine(false);
                    x = 0;
                    continue;
                }

                line.Items.Add(ToItem(segment, ItemKind.Text, x, shaped, text, measurer));
                x += shaped.Width;
                start += length;

                if (start < runes.Count)
                {
                    finishLine(false);
                    x = 0;
                }
            }
        }
    }

    static ShapedText Shape(Segment segment, string text, TextMeasurer measurer) =>
        measurer.Shape(segment.Face, segment.Size, text, segment.Style.WidthScale, segment.Style.LetterSpacing,
            segment.Direction == TextDirection.RightToLeft);

    static LineItem ToItem(Segment segment, ItemKind kind, double x, ShapedText shaped, string text, TextMeasurer measurer)
    {
        var (ascent, descent, leading) = measurer.Metrics(segment.Face, segment.BaseSize);
        return new LineItem
        {
            Kind = kind, X = x, Width = shaped.Width, Text = text, Shaped = shaped, Direction = segment.Direction,
            Face = segment.Face, Size = segment.Size, Color = segment.Color, Underline = segment.Style.Underline,
            Strike = segment.Style.Strike, Highlight = segment.Style.Highlight,
            LetterSpacing = segment.Style.LetterSpacing, WidthScale = segment.Style.WidthScale,
            Url = segment.Url, Shift = segment.Shift, Field = segment.Field,
            Ascent = ascent, Descent = descent, Leading = leading,
        };
    }

    static void TrimTrailingSpaces(LaidOutLine line)
    {
        while (line.Items.Count > 0 && line.Items[^1].Kind == ItemKind.Space)
            line.Items.RemoveAt(line.Items.Count - 1);
    }

    static void Reorder(LaidOutLine line, bool rightToLeft)
    {
        var items = line.Items;
        if (items.Count == 0) return;
        var baseLevel = rightToLeft ? 1 : 0;
        var levels = new int[items.Count];
        var hasRtl = rightToLeft;

        for (var i = 0; i < items.Count; i++)
        {
            levels[i] = items[i].Direction switch
            {
                TextDirection.RightToLeft => 1,
                TextDirection.LeftToRight => rightToLeft ? 2 : 0,
                _ => -1,
            };
            if (items[i].Direction == TextDirection.RightToLeft) hasRtl = true;
        }
        if (!hasRtl) return;

        for (var i = 0; i < levels.Length; i++)
        {
            if (levels[i] >= 0) continue;
            var before = i > 0 ? levels[i - 1] : baseLevel;
            var next = i + 1;
            while (next < levels.Length && levels[next] < 0) next++;
            var after = next < levels.Length ? levels[next] : baseLevel;
            var resolved = before == after ? before : baseLevel;
            for (var k = i; k < next; k++) levels[k] = resolved;
        }

        var order = Enumerable.Range(0, items.Count).ToList();
        for (var level = levels.Max(); level >= 1; level--)
        {
            var start = -1;
            for (var i = 0; i <= order.Count; i++)
            {
                var inRun = i < order.Count && levels[order[i]] >= level;
                if (inRun && start < 0) start = i;
                if (!inRun && start >= 0)
                {
                    order.Reverse(start, i - start);
                    start = -1;
                }
            }
        }

        var reordered = order.Select(index => items[index]).ToList();
        double x = 0;
        foreach (var item in reordered)
        {
            item.X = x;
            x += item.Width;
        }
        items.Clear();
        items.AddRange(reordered);
    }

    static double? TabTarget(Paragraph paragraph, List<Atom> atoms, int tabIndex, double x, double lineWidth)
    {
        var stops = paragraph.TabStops;
        double extent = 0;
        for (var i = tabIndex + 1; i < atoms.Count; i++)
        {
            var atom = atoms[i];
            if (atom.Kind is AtomKind.Tab or AtomKind.Break) break;
            extent += atom.Kind == AtomKind.Image ? atom.Width : atom.Segments.Sum(s => s.Width);
        }

        if (stops != null)
        {
            foreach (var stop in stops.OrderBy(s => s.Position))
            {
                if (stop.Position <= x + 0.01) continue;
                var target = stop.Alignment switch
                {
                    TabAlignment.Right => stop.Position - extent,
                    TabAlignment.Center => stop.Position - extent / 2,
                    _ => stop.Position,
                };
                if (target < x) target = stop.Position;
                return target <= lineWidth + 0.01 ? target : null;
            }
        }

        var next = (Math.Floor(x / LayoutDefaults.TabStopIntervalPoints + 0.001) + 1) * LayoutDefaults.TabStopIntervalPoints;
        return next <= lineWidth + 0.01 ? next : null;
    }

    static void Measure(Paragraph paragraph, LaidOutLine line, (double Ascent, double Descent, double Leading) defaults)
    {
        double ascent = 0, descent = 0, leading = 0;
        var hasText = false;
        foreach (var item in line.Items)
        {
            if (item.Kind is ItemKind.Text or ItemKind.Space)
            {
                hasText = true;
                ascent = Math.Max(ascent, item.Ascent);
                descent = Math.Max(descent, item.Descent);
                leading = Math.Max(leading, item.Leading);
            }
            else if (item.Kind == ItemKind.Image)
            {
                ascent = Math.Max(ascent, item.Height);
            }
        }
        if (!hasText)
        {
            if (ascent == 0) ascent = defaults.Ascent;
            descent = Math.Max(descent, line.Items.Any(i => i.Kind == ItemKind.Image) ? 0 : defaults.Descent);
            leading = defaults.Leading;
        }

        var natural = ascent + descent + leading;
        var height = natural;
        if (paragraph.LineSpacing is { } spacing && spacing > 0)
        {
            height = paragraph.LineSpacingRule switch
            {
                LineSpacingRule.Exactly => spacing,
                LineSpacingRule.AtLeast => Math.Max(natural, spacing),
                _ => natural * spacing,
            };
        }

        line.Height = height;
        line.Baseline = paragraph.LineSpacingRule == LineSpacingRule.Exactly && paragraph.LineSpacing > 0
            ? Math.Max(height - descent - leading, 0)
            : ascent + (height - natural);

        line.Width = line.Items.Count == 0 ? 0 : line.Items.Max(i => i.X + i.Width);
    }

    static void Justify(LaidOutLine line, double lineWidth)
    {
        var extra = lineWidth - line.Width;
        if (extra <= 0) return;

        var spaces = line.Items.Where(i => i.Kind == ItemKind.Space).ToList();
        if (spaces.Count == 0) return;

        var share = extra / spaces.Count;
        double shift = 0;
        foreach (var item in line.Items)
        {
            item.X += shift;
            if (item.Kind != ItemKind.Space) continue;
            item.Width += share;
            shift += share;
        }
        line.Width = lineWidth;
    }

    public static void Emit(LaidOutLine line, double x, double y, Func<FieldKind, string> fieldText,
        TextMeasurer measurer, List<DrawOp> ops)
    {
        var left = x + line.Indent;
        var baseline = y + line.Baseline;

        LinkOp? pendingLink = null;
        void FlushLink()
        {
            if (pendingLink != null) ops.Add(pendingLink);
            pendingLink = null;
        }

        foreach (var item in line.Items)
        {
            var itemX = left + item.X;
            switch (item.Kind)
            {
                case ItemKind.Text:
                case ItemKind.Space:
                {
                    var text = item.Text;
                    var shaped = item.Shaped;
                    if (item.Field is { } field)
                    {
                        text = fieldText(field);
                        shaped = measurer.Shape(item.Face!, item.Size, text, item.WidthScale, item.LetterSpacing,
                            item.Direction == TextDirection.RightToLeft);
                    }
                    var itemBaseline = baseline + item.Shift;
                    if (item.Highlight is { } highlight)
                        ops.Add(new FillRectOp(itemX, y, item.Width, line.Height, highlight));
                    if (item.Kind == ItemKind.Text || item.Underline != null)
                        ops.Add(new TextOp(text, shaped, item.Face!, item.Size, item.Color, itemX, itemBaseline,
                            item.WidthScale));

                    var (position, thickness) = measurer.Underline(item.Face!, item.Size);
                    if (item.Underline is { } underline && (item.Kind == ItemKind.Text || underline != UnderlineStyle.Words))
                    {
                        var lineY = itemBaseline + position;
                        ops.Add(new LineOp(itemX, lineY, itemX + item.Width, lineY, thickness, item.Color, underline));
                    }
                    if (item.Strike != StrikeStyle.None)
                    {
                        var strikeY = itemBaseline - item.Ascent * 0.3;
                        if (item.Strike == StrikeStyle.Single)
                            ops.Add(new LineOp(itemX, strikeY, itemX + item.Width, strikeY, thickness, item.Color));
                        else
                        {
                            ops.Add(new LineOp(itemX, strikeY - thickness, itemX + item.Width, strikeY - thickness, thickness, item.Color));
                            ops.Add(new LineOp(itemX, strikeY + thickness, itemX + item.Width, strikeY + thickness, thickness, item.Color));
                        }
                    }
                    break;
                }
                case ItemKind.Image:
                    ops.Add(new ImageOp(item.Image!, itemX, baseline - item.Height, item.Width, item.Height));
                    break;
                case ItemKind.Destination:
                    ops.Add(new DestinationOp(item.Destination!, itemX, y));
                    break;
            }

            if (item.Url == null)
            {
                FlushLink();
                continue;
            }

            var adjacent = pendingLink != null && pendingLink.Url == item.Url
                && Math.Abs(itemX - (pendingLink.X + pendingLink.Width)) < 0.5;
            if (adjacent)
                pendingLink = pendingLink! with { Width = itemX + item.Width - pendingLink.X };
            else
            {
                FlushLink();
                pendingLink = new LinkOp(itemX, y, item.Width, line.Height, item.Url);
            }
        }
        FlushLink();
    }
}
