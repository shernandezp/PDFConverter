// PdfInspector — diagnostic tool for inspecting generated PDF files.
// Shows page count, image resources, transforms, and text content per page.

using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;

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

var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import);
Console.WriteLine($"File: {path}");
Console.WriteLine($"Pages: {doc.PageCount}");

for (int i = 0; i < doc.PageCount; i++)
{
    var page = doc.Pages[i];
    Console.WriteLine($"\n--- Page {i + 1} ({page.Width.Point:F0}x{page.Height.Point:F0}pt) ---");

    // List image XObjects
    var resources = page.Elements.GetDictionary("/Resources");
    if (resources != null)
    {
        var xobjects = resources.Elements.GetDictionary("/XObject");
        if (xobjects != null)
        {
            Console.WriteLine($"  Images ({xobjects.Elements.Count}):");
            foreach (var key in xobjects.Elements.Keys)
            {
                try
                {
                    var xref = xobjects.Elements.GetReference(key);
                    if (xref?.Value is PdfDictionary xobj)
                    {
                        var subtype = xobj.Elements.GetString("/Subtype");
                        var w = xobj.Elements.GetInteger("/Width");
                        var h = xobj.Elements.GetInteger("/Height");
                        var filter = xobj.Elements.GetString("/Filter");
                        var len = xobj.Elements.GetInteger("/Length");
                        Console.WriteLine($"    {key}: {subtype} {w}x{h}px filter={filter} {len}bytes");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    {key}: error reading: {ex.Message}");
                }
            }
        }
        else
        {
            Console.WriteLine("  No images");
        }
    }

    // Parse content stream for image placement and text
    try
    {
        var content = ContentReader.ReadContent(page);
        ExtractOps(content, "  ");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  Content parse error: {ex.Message}");
    }
}

static void ExtractOps(CSequence seq, string indent)
{
    string? lastCm = null;
    var textParts = new List<string>();

    foreach (var obj in seq)
    {
        if (obj is COperator op)
        {
            var name = op.OpCode.OpCodeName.ToString();

            // Track transform matrix
            if (name == "cm" && op.Operands.Count >= 6)
            {
                var vals = op.Operands.Select(o => o.ToString()).ToArray();
                lastCm = string.Join(" ", vals);
            }

            // Image draw
            if (name == "Do" && op.Operands.Count > 0)
            {
                var imgName = op.Operands[0].ToString();
                if (lastCm != null)
                    Console.WriteLine($"{indent}DrawImage {imgName} transform=[{lastCm}]");
                else
                    Console.WriteLine($"{indent}DrawImage {imgName}");
                lastCm = null;
            }

            // Text show operators
            if (name == "Tj" && op.Operands.Count > 0)
            {
                textParts.Add(op.Operands[0].ToString()?.Trim('(', ')') ?? "");
            }
            if (name == "TJ" && op.Operands.Count > 0)
            {
                foreach (var item in op.Operands)
                {
                    if (item is CString s)
                        textParts.Add(s.Value);
                    else if (item is CSequence arr)
                    {
                        foreach (var sub in arr)
                        {
                            if (sub is CString ss)
                                textParts.Add(ss.Value);
                        }
                    }
                }
            }

            // End text block — flush collected text
            if (name == "ET" && textParts.Count > 0)
            {
                var line = string.Join("", textParts).Trim();
                if (!string.IsNullOrWhiteSpace(line))
                {
                    if (line.Length > 80) line = line[..80] + "...";
                    Console.WriteLine($"{indent}Text: \"{line}\"");
                }
                textParts.Clear();
            }
        }
        else if (obj is CSequence inner)
        {
            ExtractOps(inner, indent);
        }
    }

    if (textParts.Count > 0)
    {
        var line = string.Join("", textParts).Trim();
        if (!string.IsNullOrWhiteSpace(line))
        {
            if (line.Length > 80) line = line[..80] + "...";
            Console.WriteLine($"{indent}Text: \"{line}\"");
        }
    }
}
