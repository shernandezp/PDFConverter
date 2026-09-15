using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;

namespace PDFConverter;

internal static class PdfImageLinks
{
    // MigraDoc cannot attach a link to an image, so the placements are recovered from the rendered
    // content stream — a "cm" transform followed by "Do" — and matched back by width.
    public static void Apply(PdfDocument pdf, IReadOnlyList<HyperlinkedImage> images)
    {
        if (images.Count == 0) return;

        foreach (var page in pdf.Pages)
        {
            try
            {
                var placed = new HashSet<int>();
                Walk(ContentReader.ReadContent(page), page, images, placed, new Transform());
            }
            catch (Exception ex)
            {
                OpenXmlHelpers.ImageLoadLogger?.Invoke($"Image hyperlink annotation failed: {ex.Message}");
            }
        }
    }

    sealed class Transform
    {
        public double ScaleX, ScaleY, OffsetX, OffsetY;
        public bool IsSet;
    }

    static void Walk(CSequence sequence, PdfPage page, IReadOnlyList<HyperlinkedImage> images,
        HashSet<int> placed, Transform transform)
    {
        foreach (var item in sequence)
        {
            if (item is CSequence nested)
            {
                Walk(nested, page, images, placed, transform);
                continue;
            }
            if (item is not COperator op) continue;

            if (op.OpCode.Name == "cm" && op.Operands.Count >= 6)
            {
                transform.ScaleX = Operand(op.Operands[0]);
                transform.ScaleY = Operand(op.Operands[3]);
                transform.OffsetX = Operand(op.Operands[4]);
                transform.OffsetY = Operand(op.Operands[5]);
                transform.IsSet = true;
            }
            else if (op.OpCode.Name == "Do" && transform.IsSet)
            {
                AddLink(page, images, placed, transform);
                transform.IsSet = false;
            }
        }
    }

    static void AddLink(PdfPage page, IReadOnlyList<HyperlinkedImage> images,
        HashSet<int> placed, Transform transform)
    {
        for (var i = 0; i < images.Count; i++)
        {
            if (placed.Contains(i)) continue;
            if (Math.Abs(transform.ScaleX - images[i].WidthPoints) >= LayoutDefaults.ImageLinkWidthTolerancePoints)
                continue;

            var rect = new PdfRectangle(
                new XPoint(transform.OffsetX, transform.OffsetY),
                new XPoint(transform.OffsetX + transform.ScaleX,
                    transform.OffsetY + Math.Abs(transform.ScaleY)));
            page.AddWebLink(rect, images[i].Url);
            placed.Add(i);
            return;
        }
    }

    static double Operand(CObject value) => value switch
    {
        CReal real => real.Value,
        CInteger integer => integer.Value,
        _ => 0,
    };
}
