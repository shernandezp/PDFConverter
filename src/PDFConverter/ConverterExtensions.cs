using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal static class ConverterExtensions
{
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

    // Skia decodes PNG, JPEG, GIF, BMP and WEBP itself; TIFF and Windows metafiles need GDI+.
    internal static byte[]? ReencodeAsPng(byte[] bytes)
    {
        if (!OperatingSystem.IsWindowsVersionAtLeast(6, 1)) return null;

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
            OpenXmlHelpers.ImageLoadLogger?.Invoke(
                $"Could not re-encode {DetectImageExtension(bytes)} image: {ex.Message}");
            return null;
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
}
