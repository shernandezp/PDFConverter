using UglyToad.PdfPig;

Console.OutputEncoding = System.Text.Encoding.UTF8;

if (args.Length < 1)
{
    Console.WriteLine("Usage: PdfInspector <file.pdf>");
    return;
}

var path = args[0];
if (!File.Exists(path))
{
    Console.WriteLine($"File not found: {path}");
    return;
}

using var document = PdfDocument.Open(path);
Console.WriteLine($"File: {path}");
Console.WriteLine($"Pages: {document.NumberOfPages}");

foreach (var page in document.GetPages())
{
    Console.WriteLine($"\n--- Page {page.Number} ({page.Width:F0}x{page.Height:F0}pt) ---");

    var fonts = page.Letters.Select(letter => letter.FontName).Distinct().ToList();
    Console.WriteLine($"  Fonts: {string.Join(", ", fonts)}");

    var images = page.GetImages().ToList();
    Console.WriteLine($"  Images ({images.Count}):");
    foreach (var image in images)
    {
        var bounds = image.BoundingBox;
        Console.WriteLine($"    {image.WidthInSamples}x{image.HeightInSamples}px at ({bounds.Left:F1}, {bounds.Bottom:F1}) " +
                          $"size {bounds.Width:F1}x{bounds.Height:F1}pt");
    }

    var annotations = page.GetAnnotations().ToList();
    if (annotations.Count > 0)
    {
        Console.WriteLine($"  Annotations ({annotations.Count}):");
        foreach (var annotation in annotations)
            Console.WriteLine($"    {annotation.Type} at ({annotation.Rectangle.Left:F1}, {annotation.Rectangle.Bottom:F1}) {annotation.Action?.Type}");
    }

    var lines = page.GetWords()
        .GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y))
        .OrderByDescending(group => group.Key);
    foreach (var line in lines)
    {
        var text = string.Join(" ", line.OrderBy(word => word.BoundingBox.Left).Select(word => word.Text));
        if (text.Length > 100) text = text[..100] + "...";
        Console.WriteLine($"  y={line.Key,6:F0}  {text}");
    }
}
