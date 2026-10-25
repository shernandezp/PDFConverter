using UglyToad.PdfPig;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Graphics;

namespace PDFConverter.Tests;

internal record PdfText(double X, double Y, string Content);

internal record PdfImage(double Width, double Height, double X, double Y);

internal record PdfLine(double X1, double Y1, double X2, double Y2, double Width);

/// <summary>Reads back generated PDFs; coordinates are PDF user space (origin bottom-left).</summary>
internal static class PdfProbe
{
    public static int PageCount(byte[] pdfBytes)
    {
        using var document = PdfDocument.Open(pdfBytes);
        return document.NumberOfPages;
    }

    public static int PageCount(string pdfPath) => PageCount(File.ReadAllBytes(pdfPath));

    public static (double Width, double Height) PageSize(byte[] pdfBytes, int pageNumber = 1)
    {
        using var document = PdfDocument.Open(pdfBytes);
        var page = document.GetPage(pageNumber);
        return (page.Width, page.Height);
    }

    /// <summary>Text grouped into visual lines: one entry per baseline, words joined with spaces.</summary>
    public static List<PdfText> Texts(byte[] pdfBytes)
    {
        var texts = new List<PdfText>();
        using var document = PdfDocument.Open(pdfBytes);
        foreach (var page in document.GetPages())
        {
            var lines = page.GetWords()
                .Where(word => !string.IsNullOrWhiteSpace(word.Text))
                .GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y, 0))
                .OrderByDescending(group => group.Key);
            foreach (var line in lines)
            {
                var ordered = line.OrderBy(word => word.BoundingBox.Left).ToList();
                var x = ordered[0].BoundingBox.Left;
                var y = ordered[0].Letters[0].StartBaseLine.Y;
                texts.Add(new PdfText(x, y, string.Join(" ", ordered.Select(word => word.Text)).Trim()));
            }
        }
        return texts;
    }

    public static List<string> Words(byte[] pdfBytes)
    {
        using var document = PdfDocument.Open(pdfBytes);
        return document.GetPages().SelectMany(page => page.GetWords()).Select(word => word.Text.Trim())
            .Where(text => text.Length > 0).ToList();
    }

    public static List<PdfImage> Images(byte[] pdfBytes, int? pageNumber = null)
    {
        var images = new List<PdfImage>();
        using var document = PdfDocument.Open(pdfBytes);
        foreach (var page in document.GetPages())
        {
            if (pageNumber != null && page.Number != pageNumber) continue;
            foreach (var image in page.GetImages())
            {
                var bounds = image.BoundingBox;
                images.Add(new PdfImage(bounds.Width, bounds.Height, bounds.Left, bounds.Bottom));
            }
        }
        return images;
    }

    public static List<PdfLine> Lines(byte[] pdfBytes, int? pageNumber = null)
    {
        var lines = new List<PdfLine>();
        using var document = PdfDocument.Open(pdfBytes);
        foreach (var page in document.GetPages())
        {
            if (pageNumber != null && page.Number != pageNumber) continue;
            foreach (var path in page.Paths)
            {
                if (path.IsClipping || !path.IsStroked) continue;
                foreach (var subpath in path)
                {
                    foreach (var command in subpath.Commands)
                    {
                        switch (command)
                        {
                            case PdfSubpath.Line line:
                                lines.Add(new PdfLine(line.From.X, line.From.Y, line.To.X, line.To.Y, path.LineWidth));
                                break;
                        }
                    }
                }
            }
        }
        return lines;
    }
}
