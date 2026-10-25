using System.Text;
using SkiaSharp;

namespace PDFConverter.Layout;

internal static class SkiaPdfWriter
{
    public static void Write(IReadOnlyList<Page> pages, Stream output, Document? document)
    {
        var subsets = BuildSubsets(pages);
        try
        {
            using var stream = new SKManagedWStream(output, false);
            var metadata = SKDocumentPdfMetadata.Default;
            metadata.Title = document?.Title ?? string.Empty;
            metadata.Author = document?.Author ?? string.Empty;
            metadata.Subject = document?.Subject ?? string.Empty;
            metadata.Keywords = document?.Keywords ?? string.Empty;
            metadata.Creator = "PDFConverter";
            metadata.Creation = DateTime.Now;
            metadata.Modified = DateTime.Now;
            using var pdf = SKDocument.CreatePdf(stream, metadata)
                ?? throw new InvalidOperationException("Could not create the PDF document.");

            foreach (var page in pages)
            {
                var canvas = pdf.BeginPage((float)page.Width, (float)page.Height);
                using var painter = new Painter(canvas, subsets);
                painter.Paint(page.Background);
                painter.Paint(page.Overlay);
                painter.Paint(page.Body);
                pdf.EndPage();
            }
            pdf.Close();
        }
        finally
        {
            foreach (var (original, subset) in subsets)
                if (!ReferenceEquals(original, subset)) subset.Dispose();
        }
    }

    static Dictionary<SKTypeface, SKTypeface> BuildSubsets(IReadOnlyList<Page> pages)
    {
        var glyphs = new Dictionary<SKTypeface, HashSet<ushort>>();
        var faces = new Dictionary<SKTypeface, FontFace>();
        foreach (var page in pages)
        {
            Collect(page.Background, glyphs, faces);
            Collect(page.Overlay, glyphs, faces);
            Collect(page.Body, glyphs, faces);
        }

        var subsets = new Dictionary<SKTypeface, SKTypeface>();
        foreach (var (typeface, used) in glyphs)
        {
            var source = faces[typeface].ReadBytes();
            var stripped = source is { } font ? FontSubsetter.StripUnusedGlyphs(font.Data, used, font.CollectionIndex) : null;
            var subset = stripped == null ? null : SKTypeface.FromData(SKData.CreateCopy(stripped));
            if (subset == null)
                OpenXmlHelpers.FontLoadLogger?.Invoke($"Embedding '{typeface.FamilyName}' without subsetting.");
            subsets[typeface] = subset ?? typeface;
        }
        return subsets;
    }

    static void Collect(List<DrawOp> ops, Dictionary<SKTypeface, HashSet<ushort>> glyphs, Dictionary<SKTypeface, FontFace> faces)
    {
        foreach (var op in ops)
        {
            if (op is ClipOp clip)
            {
                Collect(clip.Children, glyphs, faces);
                continue;
            }
            if (op is not TextOp text) continue;

            var typeface = text.Face.Typeface;
            if (!glyphs.TryGetValue(typeface, out var used))
            {
                glyphs[typeface] = used = [];
                faces[typeface] = text.Face;
            }
            foreach (var glyph in text.Shaped.Glyphs) used.Add(glyph);
        }
    }

    sealed class Painter(SKCanvas canvas, Dictionary<SKTypeface, SKTypeface> subsets) : IDisposable
    {
        readonly Dictionary<(SKTypeface, float, bool, bool, float), SKFont> _fonts = [];
        readonly SKPaint _fill = new() { IsAntialias = true, Style = SKPaintStyle.Fill };
        readonly SKPaint _stroke = new() { IsAntialias = true, Style = SKPaintStyle.Stroke };
        readonly SKSamplingOptions _sampling = new(SKFilterMode.Linear, SKMipmapMode.Linear);

        public void Paint(List<DrawOp> ops)
        {
            foreach (var op in ops) Paint(op);
        }

        void Paint(DrawOp op)
        {
            switch (op)
            {
                case TextOp text:
                    PaintText(text);
                    break;

                case ImageOp image:
                {
                    var bitmap = image.Image.Image;
                    if (bitmap == null) break;
                    var destination = new SKRect((float)image.X, (float)image.Y,
                        (float)(image.X + image.Width), (float)(image.Y + image.Height));
                    if (image.Image.SourceRect is { } source)
                        canvas.DrawImage(bitmap, source, destination, _sampling);
                    else
                        canvas.DrawImage(bitmap, destination, _sampling);
                    break;
                }

                case FillRectOp rect:
                    _fill.Color = rect.Color;
                    canvas.DrawRect((float)rect.X, (float)rect.Y, (float)rect.Width, (float)rect.Height, _fill);
                    break;

                case LineOp line when line.Width > 0:
                    _stroke.Color = line.Color;
                    _stroke.StrokeWidth = (float)line.Width;
                    _stroke.PathEffect = Dash(line.Pattern, (float)line.Width);
                    canvas.DrawLine((float)line.X1, (float)line.Y1, (float)line.X2, (float)line.Y2, _stroke);
                    _stroke.PathEffect?.Dispose();
                    _stroke.PathEffect = null;
                    break;

                case LinkOp link when link.Url.StartsWith('#'):
                    canvas.DrawLinkDestinationAnnotation(Rect(link.X, link.Y, link.Width, link.Height), link.Url[1..]);
                    break;

                case LinkOp link:
                    canvas.DrawUrlAnnotation(Rect(link.X, link.Y, link.Width, link.Height), link.Url);
                    break;

                case DestinationOp destination:
                    canvas.DrawNamedDestinationAnnotation(new SKPoint((float)destination.X, (float)destination.Y), destination.Name);
                    break;

                case ClipOp clip:
                    canvas.Save();
                    canvas.ClipRect(Rect(clip.X, clip.Y, clip.Width, clip.Height));
                    Paint(clip.Children);
                    canvas.Restore();
                    break;
            }
        }

        void PaintText(TextOp text)
        {
            var shaped = text.Shaped;
            if (shaped.Glyphs.Length == 0) return;

            _fill.Color = text.Color;
            var font = FontFor(text.Face, text.Size, text.WidthScale);
            var utf8 = Encoding.UTF8.GetBytes(text.Text);
            var clusters = Utf8Clusters(text.Text, shaped.Clusters);

            using var builder = new SKTextBlobBuilder();
            var run = builder.AllocatePositionedTextRun(font, shaped.Glyphs.Length, utf8.Length);
            run.SetGlyphs(shaped.Glyphs);
            run.SetPositions(shaped.Points);
            run.SetText(utf8);
            run.SetClusters(clusters);
            using var blob = builder.Build();
            if (blob != null) canvas.DrawText(blob, (float)text.X, (float)text.Baseline, _fill);
        }

        static uint[] Utf8Clusters(string text, uint[] utf16Clusters)
        {
            var offsets = new uint[text.Length + 1];
            uint bytes = 0;
            for (var i = 0; i < text.Length; i++)
            {
                offsets[i] = bytes;
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    bytes += 4;
                    offsets[++i] = bytes;
                    continue;
                }
                bytes += text[i] < 0x80 ? 1u : text[i] < 0x800 ? 2u : 3u;
            }
            offsets[text.Length] = bytes;

            var clusters = new uint[utf16Clusters.Length];
            for (var i = 0; i < clusters.Length; i++)
                clusters[i] = offsets[Math.Min(utf16Clusters[i], (uint)text.Length)];
            return clusters;
        }

        static SKRect Rect(double x, double y, double width, double height) =>
            new((float)x, (float)y, (float)(x + width), (float)(y + height));

        static SKPathEffect? Dash(UnderlineStyle pattern, float width)
        {
            var unit = Math.Max(width, 0.5f);
            return pattern switch
            {
                UnderlineStyle.Dotted => SKPathEffect.CreateDash([unit, unit * 1.5f], 0),
                UnderlineStyle.Dash => SKPathEffect.CreateDash([unit * 4, unit * 2], 0),
                UnderlineStyle.DotDash => SKPathEffect.CreateDash([unit * 4, unit * 2, unit, unit * 2], 0),
                UnderlineStyle.DotDotDash => SKPathEffect.CreateDash([unit * 4, unit * 2, unit, unit * 2, unit, unit * 2], 0),
                _ => null,
            };
        }

        SKFont FontFor(FontFace face, double size, double scaleX)
        {
            var typeface = subsets.TryGetValue(face.Typeface, out var subset) ? subset : face.Typeface;
            var key = (typeface, (float)size, face.FakeBold, face.FakeItalic, (float)scaleX);
            if (_fonts.TryGetValue(key, out var font)) return font;

            font = new SKFont(typeface, (float)size)
            {
                Embolden = face.FakeBold,
                SkewX = face.FakeItalic ? -0.25f : 0f,
                ScaleX = (float)scaleX,
                Hinting = SKFontHinting.None,
                Subpixel = true,
            };
            _fonts[key] = font;
            return font;
        }

        public void Dispose()
        {
            foreach (var font in _fonts.Values) font.Dispose();
            _fill.Dispose();
            _stroke.Dispose();
        }
    }
}
