using System.Collections.Concurrent;
using System.Text;
using PdfSharp.Drawing;

namespace PDFConverter;

internal static class TextMeasure
{
    const double SlackPoints = 2.0;

    static readonly Lock Gate = new();
    static readonly ConcurrentDictionary<string, XFont> Fonts = new(StringComparer.Ordinal);
    static XGraphics? s_context;

    // Word breaks a word wider than its cell at the cell edge; MigraDoc lets it overflow into the
    // neighbouring cell instead, so the break points are measured and inserted here.
    public static List<string> SplitOverlongWords(string text, RunFormat format, double maxWidthPoints)
    {
        if (maxWidthPoints <= 0 || string.IsNullOrEmpty(text)) return [text];

        var font = FontFor(format);
        if (font == null || Width(text, font) <= maxWidthPoints) return [text];

        var lines = new List<string>();
        var buffer = new StringBuilder();
        var index = 0;

        while (index < text.Length)
        {
            if (char.IsWhiteSpace(text[index]))
            {
                buffer.Append(text[index++]);
                continue;
            }

            var start = index;
            while (index < text.Length && !char.IsWhiteSpace(text[index])) index++;
            var word = text[start..index];

            if (Width(word, font) <= maxWidthPoints)
            {
                buffer.Append(word);
                continue;
            }

            var chunks = Chunk(word, font, maxWidthPoints);
            for (var i = 0; i < chunks.Count; i++)
            {
                buffer.Append(chunks[i]);
                if (i == chunks.Count - 1) break;
                lines.Add(buffer.ToString());
                buffer.Clear();
            }
        }

        lines.Add(buffer.ToString());
        return lines;
    }

    static List<string> Chunk(string word, XFont font, double maxWidthPoints)
    {
        var chunks = new List<string>();
        var start = 0;

        while (start < word.Length)
        {
            var length = 1;
            while (start + length < word.Length
                && Width(word.Substring(start, length + 1), font) <= maxWidthPoints)
                length++;

            chunks.Add(word.Substring(start, length));
            start += length;
        }
        return chunks;
    }

    static double Width(string text, XFont font)
    {
        lock (Gate)
        {
            s_context ??= XGraphics.CreateMeasureContext(
                new XSize(1000, 1000), XGraphicsUnit.Point, XPageDirection.Downwards);
            return s_context.MeasureString(text, font).Width + SlackPoints;
        }
    }

    static XFont? FontFor(RunFormat format)
    {
        var family = string.IsNullOrEmpty(format.FontFamily) ? "Arial" : format.FontFamily;
        var size = format.Size ?? LayoutDefaults.FontSizePoints;
        var style = (format.Bold ? XFontStyleEx.Bold : XFontStyleEx.Regular)
            | (format.Italic ? XFontStyleEx.Italic : XFontStyleEx.Regular);

        var key = $"{family}|{size}|{(int)style}";
        if (Fonts.TryGetValue(key, out var cached)) return cached;

        try
        {
            var font = new XFont(family, size, style);
            Fonts[key] = font;
            return font;
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.FontLoadLogger?.Invoke($"Cannot measure text in '{family}': {ex.Message}");
            return null;
        }
    }
}
