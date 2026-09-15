using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal static class ConverterExtensions
{
    const string TempFilePrefix = "pdfconverter_img_";

    internal static string ToRoman(int number)
    {
        if (number < 1) return string.Empty;
        var map = new[]
        {
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
        };
        var result = new StringBuilder();
        foreach (var (value, symbol) in map)
        {
            while (number >= value)
            {
                result.Append(symbol);
                number -= value;
            }
        }
        return result.ToString();
    }

    internal static string GetParagraphText(W.Paragraph? p)
    {
        if (p == null) return string.Empty;
        var sb = new StringBuilder();
        foreach (var node in p.Descendants())
        {
            if (node is W.Text t) sb.Append(t.Text);
            else if (node is W.TabChar) sb.Append('\t');
            else if (node is W.Break) sb.Append('\n');
        }
        return sb.ToString();
    }

    internal static bool ContainsEmoji(string text)
    {
        for (var i = 0; i < text.Length; i++)
            if (IsEmojiChar(text[i])) return true;
        return false;
    }

    internal static List<(string Text, bool IsEmoji)> SplitEmojiSegments(string text)
    {
        var segments = new List<(string, bool)>();
        if (string.IsNullOrEmpty(text)) return segments;

        var current = new StringBuilder();
        var currentIsEmoji = false;

        for (var i = 0; i < text.Length; i++)
        {
            var charIsEmoji = IsEmojiChar(text[i]);

            if (charIsEmoji != currentIsEmoji && current.Length > 0)
            {
                segments.Add((current.ToString(), currentIsEmoji));
                current.Clear();
            }
            currentIsEmoji = charIsEmoji;

            current.Append(text[i]);
            if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                current.Append(text[++i]);
        }

        if (current.Length > 0) segments.Add((current.ToString(), currentIsEmoji));
        return segments;
    }

    static bool IsEmojiChar(char c) =>
        char.IsHighSurrogate(c)
        || c is >= '\u2600' and <= '\u27BF'
        || c is >= '\uFE00' and <= '\uFE0F'
        || c == '\u200D';

    internal static string SaveTempImage(byte[] bytes)
    {
        var extension = DetectImageExtension(bytes);
        var path = Path.Combine(Path.GetTempPath(), $"{TempFilePrefix}{Guid.NewGuid():N}{extension}");

        var normalized = NormalizeForPdfSharp(bytes, extension);
        if (normalized != null)
        {
            File.WriteAllBytes(path, normalized);
            return path;
        }

        File.WriteAllBytes(path, bytes);
        return path;
    }

    // PdfSharp rejects indexed-palette and sub-byte PNGs, and cannot read BMP/TIFF/metafiles at
    // all; re-encoding them as 32-bit ARGB PNG is the only way to keep those pictures visible.
    static byte[]? NormalizeForPdfSharp(byte[] bytes, string extension)
    {
        var needsReencode = extension switch
        {
            ".bmp" or ".tif" or ".emf" or ".wmf" => true,
            ".png" => bytes.Length >= 26 && (bytes[25] == 3 || bytes[24] < 8),
            _ => false,
        };
        if (!needsReencode || !OperatingSystem.IsWindowsVersionAtLeast(6, 1)) return null;

        try
        {
            using var source = new MemoryStream(bytes);
            using var original = Image.FromStream(source);
            using var converted = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb);
            using (var graphics = Graphics.FromImage(converted))
            {
                graphics.DrawImage(original, 0, 0, original.Width, original.Height);
            }
            using var output = new MemoryStream();
            converted.Save(output, ImageFormat.Png);
            return output.ToArray();
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Could not re-encode {extension} image: {ex.Message}");
            return null;
        }
    }

    internal static void TryDeleteTempFile(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Could not delete temporary file '{path}': {ex.Message}");
        }
    }

    internal static string DetectImageExtension(byte[] bytes)
    {
        if (StartsWith(bytes, 0x89, 0x50, 0x4E, 0x47)) return ".png";
        if (StartsWith(bytes, 0xFF, 0xD8)) return ".jpg";
        if (StartsWith(bytes, 0x47, 0x49, 0x46)) return ".gif";
        if (StartsWith(bytes, 0x42, 0x4D)) return ".bmp";
        if (StartsWith(bytes, 0x49, 0x49, 0x2A, 0x00) || StartsWith(bytes, 0x4D, 0x4D, 0x00, 0x2A)) return ".tif";
        if (StartsWith(bytes, 0xD7, 0xCD, 0xC6, 0x9A)) return ".wmf";
        if (IsEnhancedMetafile(bytes)) return ".emf";
        return ".png";
    }

    static bool IsEnhancedMetafile(byte[] bytes) =>
        StartsWith(bytes, 0x01, 0x00, 0x00, 0x00)
        && bytes.Length >= 44
        && bytes[40] == 0x20 && bytes[41] == 0x45 && bytes[42] == 0x4D && bytes[43] == 0x46;

    static bool StartsWith(byte[] bytes, params byte[] signature)
    {
        if (bytes.Length < signature.Length) return false;
        for (var i = 0; i < signature.Length; i++)
            if (bytes[i] != signature[i]) return false;
        return true;
    }

    internal static byte[] ApplySrcRectCrop(byte[] imageBytes, int cropLeft, int cropTop, int cropRight, int cropBottom)
    {
        if (cropLeft == 0 && cropTop == 0 && cropRight == 0 && cropBottom == 0) return imageBytes;
        if (!OperatingSystem.IsWindowsVersionAtLeast(6, 1)) return imageBytes;

        try
        {
            using var source = new MemoryStream(imageBytes);
            using var bitmap = new Bitmap(source);

            var left = (int)(bitmap.Width * cropLeft / Units.ThousandthsOfPercentPerWhole);
            var top = (int)(bitmap.Height * cropTop / Units.ThousandthsOfPercentPerWhole);
            var right = bitmap.Width - (int)(bitmap.Width * cropRight / Units.ThousandthsOfPercentPerWhole);
            var bottom = bitmap.Height - (int)(bitmap.Height * cropBottom / Units.ThousandthsOfPercentPerWhole);

            var rect = new Rectangle(left, top, Math.Max(1, right - left), Math.Max(1, bottom - top));
            using var cropped = bitmap.Clone(rect, bitmap.PixelFormat);
            using var output = new MemoryStream();
            cropped.Save(output, ImageFormat.Png);
            return output.ToArray();
        }
        catch
        {
            return imageBytes;
        }
    }
}
